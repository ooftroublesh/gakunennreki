import streamlit as st
from datetime import date, timedelta, datetime
import jpholiday
from pyscipopt import Model, quicksum
from openpyxl import Workbook
from collections import defaultdict
import io
from openpyxl import load_workbook
from datetime import date
from collections import defaultdict
import re
import bisect


from collections import defaultdict
import bisect
import math
import statistics

def build_positions(actual_weekdays):
    """
    各曜日ごとに、その曜日が出現するインデックスを昇順ソートして保持する。
    actual_weekdays: List[int] (0=月曜,…,6=日曜)
    戻り値: defaultdict(list)  key=weekday, value=[idx1, idx2, …]
    """
    pos = defaultdict(list)
    for idx, wd in enumerate(actual_weekdays):
        pos[wd].append(idx)
    return pos

dup_dates_global: set[date] = set()  
def detect_dup_shift_days(events):
    """
    events : [(date_obj, weekday_int), …]  # weekday_int 0=月, …,4=金
    同じ weekday_int が 2 日連続で出たとき “2 日目” だけを set で返す
    """
    events.sort(key=lambda x: x[0])
    dup = set()
    for (d_prev, w_prev), (d_cur, w_cur) in zip(events, events[1:]):
        if d_cur == d_prev + timedelta(days=1) and w_cur == w_prev:
            dup.add(d_cur)
    return dup

def nearest_distance_in_days(idx, target_wd, dates, positions):
    """
    idx番目の日付 dates[idx] を基準にして、
    target_wd の同じ曜日が直前・直後にいつあるかを探し、
    その日数差の合計を返す。見つからなければ math.inf を返す。

    dates: List[date]
    target_wd: int (0=月曜,…,6=日曜)
    positions: build_positions(actual_weekdays) の戻り値
    """
    current = dates[idx]
    pos_list = positions[target_wd]
    # bisect_left で自身の位置を含むように
    j = bisect.bisect_left(pos_list, idx)

    # 前の同曜日までの日数
    if j > 0:
        prev_idx = pos_list[j - 1]
        d_before = (current - dates[prev_idx]).days
    else:
        d_before = math.inf

    # 次の同曜日までの日数
    if j < len(pos_list):
        next_idx = pos_list[j]
        d_after = (dates[next_idx] - current).days
    else:
        d_after = math.inf

    return d_before + d_after


def normalize_run(run, dates, actual_weekdays, assigned_weekdays):
    """
    run: 連続的に「実曜 != 割当曜」になっているインデックスのリスト
    dates: List[date]（最適化対象の日付リスト）
    actual_weekdays: dates[i].weekday() のリスト
    assigned_weekdays: 最適化で割り当てた曜日（整数）のリスト

    各 run 内で swap を繰り返し、
    「同じ曜日同士ができるだけ近くなる」ように平準化する。
    """
    # 各曜日ごとの位置情報はグローバル変数にセットしておく
    # normalize_run を呼ぶ前に必ず
    #   global positions_by_weekday
    #   positions_by_weekday = build_positions(actual_weekdays)
    # を行ってください。

    # 現在の割当候補リストをコピー
    W = [assigned_weekdays[i] for i in run]
    # 各要素の初期距離
    D = [nearest_distance_in_days(i, w, dates, positions_by_weekday)
         for i, w in zip(run, W)]

    improved = True
    while improved:
        improved = False
        # 全ペアを試す
        for a in range(len(run)):
            for b in range(a + 1, len(run)):
                # swap 後の距離
                Da = nearest_distance_in_days(run[a], W[b], dates, positions_by_weekday)
                Db = nearest_distance_in_days(run[b], W[a], dates, positions_by_weekday)
                old_var = statistics.pvariance([D[a], D[b]])
                new_var = statistics.pvariance([Da, Db])
                if new_var < old_var:
                    # スワップ実行
                    W[a], W[b] = W[b], W[a]
                    D[a], D[b] = Da, Db
                    improved = True

    # 最後に assigned_weekdays に書き戻し
    for idx, new_w in zip(run, W):
        assigned_weekdays[idx] = new_w



# ── キャッシュ定義 ──
@st.cache_data(show_spinner=False)
def get_holidays(year):
    """年度中の祝日リストを返す"""
    start = date(year, 4, 1)
    end = date(year + 1, 3, 31)
    return jpholiday.between(start, end)

@st.cache_resource(show_spinner=False)
def load_calendar_template(path: str):
    """Excel テンプレートを一度だけロード"""
    return load_workbook(path)

# ── 再実行ショートカット ──
try:
    rerun = st.experimental_rerun
except AttributeError:
    from streamlit.runtime.scriptrunner import RerunException, RerunData
    def rerun():
        raise RerunException(RerunData())
    
st.set_page_config(page_title="筑波大学  学年暦作成ツール", layout="centered")
st.title("\U0001F4C5 筑波大学  学年暦作成ツール")


# st.markdown("---")
# st.markdown("初期設定")
st.title("初期設定")
# === 年度の選択 ===
current_year = datetime.now().year
year_options = [current_year + i for i in range(-5, 6)]
year_start = st.selectbox(
     "作成する年度を選択してください",
     options=year_options,
     index=5,
     key="year_start"      # ← これで st.session_state.year_start が自動登録される
     )

# ①── 学年度の最小／最大日を定義
min_academic = date(year_start, 4, 1)
max_academic = date(year_start + 1, 3, 31)


# === 春学期開始日 & 終了日 ===
default_start_summer = date(year_start, 4, 14)
start_date_summer = st.date_input(
    "春学期の授業開始日",
    value=default_start_summer,
    key="start_summer",
    min_value=min_academic,      # ← 4/1 以降のみ選べる
    max_value=max_academic       # ← 翌年3/31 以前のみ選べる
)


# 春学期の終了日は、秋学期開始日の前日に自動設定するため、ここでは入力不要
# === 秋学期開始日 ===
default_start_autumn = date(year_start, 10, 1)
start_date_autumn = st.date_input(
    "秋学期の授業開始日",
    value=default_start_autumn,
    key="start_autumn",
    min_value=min_academic,      # ← 4/1 以降のみ選べる
    max_value=max_academic       # ← 翌年3/31 以前のみ選べる
)


# 自動計算した「春学期終了日」と「秋学期終了日」を画面に表示
end_date_summer = start_date_autumn - timedelta(days=1)
#st.write(f"春学期の授業終了日（自動設定）：{end_date_summer}")

end_date_autumn = date(year_start + 1, 3, 31)
#st.write(f"秋学期の授業終了日（固定）：{end_date_autumn}")
autumn3_start = st.date_input(
    "秋Cモジュールの開始日",
    value=date(year_start+1, 1, 1),
    min_value=min_academic,    # 2025/4/1 以降しか選べない
    max_value=max_academic,    # 2026/3/31 以前しか選べない
    key=f"autumn3_start_{year_start}"
)
st.session_state["autumn3_start"] = autumn3_start


st.markdown("---")



from datetime import date, timedelta
from collections import defaultdict
import streamlit as st


# ── セッション初期化 ──
if "manual_holidays_all" not in st.session_state:
    st.session_state.manual_holidays_all = set()
if "event_periods" not in st.session_state:
    # { イベント名: [date1, date2, ...] }
    st.session_state.event_periods = {}

if "event_labels_remark" not in st.session_state:
    # 日付キー → 備考リスト のマッピングとして defaultdict を用意
    st.session_state["event_labels_remark"] = defaultdict(list)

# ── 備考再構築関数 ──
def rebuild_event_remarks():
    """event_periods から event_labels_remark を再生成する"""
    # まず既存の備考辞書をリセット
    st.session_state.event_labels_remark = defaultdict(list)
    # 1) base_name ごとに日付をまとめる
    grouped = defaultdict(set)
    for name_with_suffix, dates in st.session_state.event_periods.items():
        base = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)
        grouped[base].update(dates)
        # if not dates:
        #     continue

        # ── 末尾に " (数字)" がついていたら除去して base_name を作る
        #    例: "共通テストに伴う休講 (2)" → base_name = "共通テストに伴う休講"
        
        # 番号付きキーならスキップ（すでに番号なしで処理済みのはずなので）
        # if base_name != name_with_suffix:
        #     continue

        # 日付リストをソートして「連続区間 (runs)」を検出する
        # ds = sorted(dates)
        # runs = []
        # run_start = run_end = ds[0]
        # for d in ds[1:]:
        #     if d == run_end + timedelta(days=1):
        #         run_end = d
        #     else:
        #         runs.append((run_start, run_end))
        #         run_start = run_end = d
        # runs.append((run_start, run_end))

        # # ── 連続区間ごとに一行分の備考文字列をつくり、月ごとに１回だけ登録する
        # for start, end in runs:
        #     if start == end:
        #         remark = f"{start.month}月{start.day}日　{base_name}"
        #     else:
        #         remark = (f"{start.month}月{start.day}日～"
        #                   f"{end.month}月{end.day}日　{base_name}")

        #     # 「start から end まで」の区間内に含まれる月を列挙し、
        #     #  それぞれの月の最初の日付をキーにして１回だけ登録する
        #     months = sorted({d.month for d in ds if start <= d <= end})
        #     for m in months:
        #         first_of_month = min(d for d in ds if start <= d <= end and d.month == m)
        #         lst = st.session_state.event_labels_remark[first_of_month]
        #         if remark not in lst:
        #             lst.append(remark)

    # 2) まとめた日付リストで備考を生成
    # 3) 連続区間を検出して備考文字列を作成
    for base_name, dates in grouped.items():
        if not dates:
            continue
        ds = sorted(dates)

        # 3-1) 連続する区間 (runs) を作る
        runs: list[tuple[date, date]] = []
        run_start = run_end = ds[0]
        for d in ds[1:]:
            if d == run_end + timedelta(days=1):
                run_end = d
            else:
                runs.append((run_start, run_end))
                run_start = run_end = d
        runs.append((run_start, run_end))

        # ④ 各連続区間ごとに一度だけ備考を追加
        for start, end in runs:
            if start == end:
                remark = f"{start.month}月{start.day}日　{base_name}"
            else:
                remark = (
                    f"{start.month}月{start.day}日～"
                    f"{end.month}月{end.day}日　{base_name}"
                )

            # 区間内の各月の最初の日付をキーにして一度だけ append
            months = sorted({d.month for d in ds if start <= d <= end})
            for m in months:
                first_of_month = min(
                    d for d in ds
                    if start <= d <= end and d.month == m
                )
                lst = st.session_state.event_labels_remark[first_of_month]
                if remark not in lst:
                    lst.append(remark)


st.title("📝 イベントと休講日の登録")

st.markdown("**イベントごとに、イベント名・イベント期間・休講扱いの有無を選択し、"
            "「📌 イベントを追加」を押してください。"
            "長期休業期間、予備日、曜日変更の案内は自動作成されます**")
st.markdown(
    """
    <style>
    /* ── クリック確定後の開始／間／終了タイル ── */
    .react-calendar__tile.react-calendar__tile--rangeStart,
    .react-calendar__tile.react-calendar__tile--rangeEnd {
        background: #f63366 !important;
        color: white      !important;
    }
    .react-calendar__tile.react-calendar__tile--range {
        background: #ffe3e8 !important;  /* 薄いピンク */
        color: black   !important;
    }

    /* ── ドラッグ（ホバー）中の開始／間／終了タイル ── */
    .react-calendar__tile.react-calendar__tile--hoverRangeStart,
    .react-calendar__tile.react-calendar__tile--hoverRangeEnd {
        background: #f63366 !important;
        color: white      !important;
    }
    .react-calendar__tile.react-calendar__tile--hoverRange {
        background: #ffe3e8 !important;
        color: black   !important;
    }

    /* ── ホバー中さらにマウスオーバーしているときの色 ── */
    .react-calendar__tile.react-calendar__tile--hoverRangeStart:hover,
    .react-calendar__tile.react-calendar__tile--hoverRangeEnd:hover,
    .react-calendar__tile.react-calendar__tile--hoverRange:hover {
        background: #cc2e57 !important;
        color: white      !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ── フォーム ──
with st.form("add_event_form", clear_on_submit=False):
    event_name = st.text_input("イベント名（例：オリエンテーション）")
    default_day = date(year_start, 4, 1)
    event_range = st.date_input(
        "期間（開始日〜終了日）",
        value=(default_day, default_day),
        min_value=min_academic,  # 2025/4/1 以降
        max_value=max_academic   # 2026/3/31 以前
    )
    mode = st.radio(
        "扱いを選択してください",
        options=["休講日とするイベント", "イベントのみ"],
        horizontal=True
    )
    submit_event = st.form_submit_button("📌 イベントを追加")
if submit_event:
    if mode == "休講日とするイベント" and not event_name.strip():
        st.warning("休講日扱いにするならイベント名が必要です。")
    else:
        name  = event_name.strip() or "<匿名イベント>"
        start, end = (event_range if isinstance(event_range, tuple)
                      else (event_range, event_range))
        dates = [start + timedelta(days=i) for i in range((end - start).days + 1)]

        # ── ① 休講扱いなら manual_holidays_all に登録 ──
        if mode == "休講日とするイベント":
            for d in dates:
                st.session_state.manual_holidays_all.add(d)

        
        key = name
        if key in st.session_state.event_periods:
            i = 2
            while True:
                candidate = f"{name} ({i})"
                if candidate not in st.session_state.event_periods:
                    key = candidate
                    break
                i += 1
        # ここで key はユニークになっている（最初は name が未使用→そのまま name、重複時は name (2),(3)…）
        lst = st.session_state.event_periods.setdefault(key, [])
        for d in dates:
            if d not in lst:
                lst.append(d)

        # ③ 備考再構築
        rebuild_event_remarks()
        # 表示用に末尾の " (数字)" を剥がす
        import re
        display_name = re.sub(r"\s*\(\d+\)$", "", key)
        st.success(f"{start}～{end} の『{display_name}』を登録しました。")


# ── イベント削除 ──
def make_delete_entries():
    """
    st.session_state.event_periods と st.session_state.manual_holidays_all から
    (日付, イベント名または None) のリストを作成し、日付順にソートして返す
    """
    entries = []
    # ① event_periods から
    for name, dates in st.session_state.event_periods.items():
        for d in dates:
            entries.append((d, name))
    # ② manual_holidays_all から（すでに event_periods に含まれる日付は除外）
    for d in st.session_state.manual_holidays_all:
        if not any(d in ds for ds in st.session_state.event_periods.values()):
            entries.append((d, None))
    # ③ 日付でソート
    entries.sort(key=lambda x: x[0])
    return entries

st.markdown("### ❌ イベントを削除（複数選択可）")

# delete_entries: [(date, event_name_or_None), …]
delete_entries = make_delete_entries()
# options: ["2025-04-01｜イベントA", "2025-04-02｜（休講日）", …]
options = [
    f"{d:%Y-%m-%d}｜{nm or '（休講日）'}"
    for d, nm in delete_entries
]

if options:
    # 複数選択可能な Multiselect を表示
    selected = st.multiselect(
        "削除したいイベント（日時）を選択してください",
        options
    )

    if st.button("🗑️ 選択したイベントをすべて削除"):
        if not selected:
            st.warning("削除する項目を選択してください。")
        else:
            for opt in selected:
                # opt の文字列から元のインデックスを取得
                idx = options.index(opt)
                d, nm = delete_entries[idx]

                # ① manual_holidays_all から削除
                st.session_state.manual_holidays_all.discard(d)

                # ② event_periods から削除（nm が None でなければ）
                if nm:
                    lst = st.session_state.event_periods.get(nm, [])
                    if d in lst:
                        lst.remove(d)
                    # もしイベントに日付が一つも残らなければ、辞書からキーを削除
                    if nm in st.session_state.event_periods and not st.session_state.event_periods[nm]:
                        del st.session_state.event_periods[nm]

            # ③ 備考を再構築
            # def rebuild_event_remarks():
            #     import re
            #     from datetime import timedelta
            #     from collections import defaultdict

            #     st.session_state.event_labels_remark = defaultdict(list)
            #     for name_with_suffix, dates in st.session_state.event_periods.items():
            #         if not dates:
            #             continue

            #         # "(2)" のような末尾を取り除いて base_name
            #         base_name = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)

            #         ds = sorted(dates)
            #         runs = []
            #         run_start = run_end = ds[0]
            #         for d0 in ds[1:]:
            #             if d0 == run_end + timedelta(days=1):
            #                 run_end = d0
            #             else:
            #                 runs.append((run_start, run_end))
            #                 run_start = run_end = d0
            #         runs.append((run_start, run_end))

            #         for st_day, ed_day in runs:
            #             if st_day == ed_day:
            #                 remark = f"{st_day.month}月{st_day.day}日　{base_name}"
            #             else:
            #                 remark = (
            #                     f"{st_day.month}月{st_day.day}日～"
            #                     f"{ed_day.month}月{ed_day.day}日　{base_name}"
            #                 )

            #             months = sorted({d0.month for d0 in ds if st_day <= d0 <= ed_day})
            #             for m in months:
            #                 first_of_month = min(
            #                     d0 for d0 in ds
            #                     if st_day <= d0 <= ed_day and d0.month == m
            #                 )
            #                 lst0 = st.session_state.event_labels_remark[first_of_month]
            #                 if remark not in lst0:
            #                     lst0.append(remark)

            # rebuild_event_remarks()
            # ③ 汎用の rebuild_event_remarks() を呼び出すだけ
            rebuild_event_remarks()
            st.success("選択したイベントをすべて削除しました。")
            

else:
    st.info("現在、削除できるイベントや休講日はありません。")

# ── 一覧表示 ──

st.markdown("### 📋 イベント一覧")
if st.session_state.event_periods:
    manual = st.session_state.manual_holidays_all
    for name, dates in st.session_state.event_periods.items():
        ds = sorted(dates)
        s, e = ds[0], ds[-1]

        # ── ここを追加 ──
        # 末尾の " (数字)" を除去して表示用の名前を作る
        base_name = re.sub(r"\s*\(\d+\)$", "", name)
        # ─────────────────

        # 番号を剥がした base_name を使って disp_name を組み立て
        disp_name = (
            f"{base_name}（休講日）"
            if any(d in manual for d in dates)
            else base_name
        )

        st.write(
            f"🔖 {s:%Y-%m-%d}"
            + (f"～{e:%Y-%m-%d}" if s != e else "")
            + f"　{disp_name}"
        )
else:
    st.info("現在、イベントは登録されていません。")
# ── 一覧表示 ──
# st.markdown("### 📋 イベント一覧")
# if st.session_state.event_periods:
#     manual = st.session_state.manual_holidays_all

#     # 1) base_name ごとに日付をまとめる
#     grouped = defaultdict(set)
#     for name_with_suffix, dates in st.session_state.event_periods.items():
#         base = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)
#         grouped[base].update(dates)

#     # 2) ソートして表示
#     for base_name, dates in sorted(grouped.items(), key=lambda x: min(x[1])):
#         ds = sorted(dates)
#         start, end = ds[0], ds[-1]
#         # 休講日扱いかどうか
#         disp_name = (
#             f"{base_name}（休講日）"
#             if any(d in manual for d in ds)
#             else base_name
#         )
#         st.write(
#             f"🔖 {start:%Y-%m-%d}"
#             + (f"～{end:%Y-%m-%d}" if start != end else "")
#             + f"　{disp_name}"
#         )
# else:
#     st.info("現在、イベントは登録されていません。")


# === 春・秋に自動分類 ===
# ―― そのままの位置で ――
manual_holidays_all = st.session_state.manual_holidays_all

spring_holidays = {d for d in manual_holidays_all if 4 <= d.month <= 9}
autumn_holidays = {d for d in manual_holidays_all if d.month <= 3 or 10 <= d.month <= 12}



st.markdown("---")
st.markdown("各学期の予備日設定")


# ── 年間（4月 start_date_summer ～ 翌年 3月 end_date_autumn）で祝日名をキャッシュ ──
all_holiday_names = {}
d = start_date_summer
while d <= end_date_autumn:
    if jpholiday.is_holiday(d):
        raw = jpholiday.is_holiday_name(d)
        if raw:
            # 振替休日を「振替休日」とだけ表示
            if "振替休日" in raw:
                all_holiday_names[d] = "振替休日"
            else:
                all_holiday_names[d] = raw
    d += timedelta(days=1)

# セッションに保存
st.session_state.holiday_names = all_holiday_names


date_list = [start_date_summer + timedelta(days=i) for i in range((end_date_summer - start_date_summer).days + 1)]
all_weekdays_summer = [d for d in date_list if d.weekday() < 5 and not jpholiday.is_holiday(d) and d not in spring_holidays ]

if len(all_weekdays_summer) < 25:
    st.error("平日が25日未満のため、最適化を実行できません。")
    st.stop()

# def build_week_penalties(week_dict, jpholiday, spring_holidays, autumn_holidays, reserve_limit=None, dup_dates=None):
#     week_penalties = {}
#     for week_id, days in week_dict.items():
#         year, weeknum = week_id

#         # 1) その週の祝日数を ISO 週（月曜～日曜）で数える
#         #    最大５日までカウント
#         monday = date.fromisocalendar(year, weeknum, 1)
#         holiday_count = 0
#         for i in range(7):
#             d0 = monday + timedelta(days=i)
#             if jpholiday.is_holiday(d0):
#                 holiday_count += 1
#                 if holiday_count >= 5:
#                     break

#         # 2) reserve_limit 超過日の有無は従来通り
#         contains_after_reserve = any(reserve_limit and d > reserve_limit for d in days)

#         # 3) available 日数（候補日に残った授業可能日）も従来通り
#         available = sum(
#             1
#             for d in days
#             if not (jpholiday.is_holiday(d) or d.weekday() >= 5 or d in spring_holidays or d in autumn_holidays)
#         )

#         # 4) ペナルティ計算
#         if contains_after_reserve:
#             penalty = 10_000.0
#         else:
#             if   available >= 5: base = 50.0
#             elif available == 4: base = 2.0
#             elif available == 3: base = 1.0
#             elif available == 2: base = 0.3
#             else:
#                 # ★ 同じ授業が 2 日連続なら 0.4、そうでなければ従来 0.1
#                 is_dup_week = dup_dates and any(d in dup_dates for d in days)
#                 base = 0.4 if is_dup_week else 0.1   # ←★ここだけ変化

#             # 祝日１日につき 0.01 下げる
#             penalty = base - 0.01 * holiday_count

#         week_penalties[week_id] = penalty

#     return week_penalties



spring_buffer_count1 = st.selectbox(
    "春学期Aモジュールの予備日（日数）を選択してください",
    options=[1, 2, 3, 4, 5],
    index=0,
    key="spring_buffer_count1"
)
spring_buffer_count2 = st.selectbox(
    "春学期Bモジュールの予備日（日数）を選択してください",
    options=[1, 2, 3, 4, 5],
    index=0,
    key="spring_buffer_count2"
)
spring_buffer_count3 = st.selectbox(
    "春学期Cモジュールの予備日（日数）を選択してください",
    options=[1, 2, 3, 4, 5],
    index=0,
    key="spring_buffer_count3"
)

autumn_buffer_count1 = st.selectbox(
    "秋学期Aモジュールの予備日（日数）を選択してください",
    options=[1, 2, 3, 4, 5],
    index=0,
    key="autumn_buffer_count1"
)
autumn_buffer_count2 = st.selectbox(
    "秋学期Bモジュールの予備日（日数）を選択してください",
    options=[1, 2, 3, 4, 5],
    index=0,
    key="autumn_buffer_count2"
)
autumn_buffer_count3 = st.selectbox(
    "秋学期Cモジュールの予備日（日数）を選択してください",
    options=[1, 2, 3, 4, 5],
    index=0,
    key="autumn_buffer_count3"
)




# -------------------------------------------
# ▼共通ユーティリティ
# -------------------------------------------
from datetime import date, timedelta
import jpholiday

def build_holiday_week_map_full(
    dates: list[date],
    jpholiday,
    manual_holidays: set[date] = set(),
) -> dict[tuple[int,int], int]:
    """
    dates に含まれる日から ISO 週を一意抽出し、
    その週に祝日 or manual_holidays があれば 1、なければ 0 を返すマップを構築。
    """
    weekids = {d.isocalendar()[:2] for d in dates}
    holiday_week = {}
    for year, weeknum in weekids:
        monday = date.fromisocalendar(year, weeknum, 1)
        flag = False
        for i in range(7):
            d0 = monday + timedelta(days=i)
            if jpholiday.is_holiday(d0) or d0 in manual_holidays:
                flag = True
                break
        holiday_week[(year, weeknum)] = 1 if flag else 0
    return holiday_week


def slide_duplicate_substitutions(dates, assigned, holiday_week_flag):
    week_of = [d.isocalendar()[:2] for d in dates]
    actual  = [d.weekday() for d in dates]

    # 週 × 曜日 → 振替インデックス
    sub_by = defaultdict(lambda: defaultdict(list))
    for j, (w_id, a, g) in enumerate(zip(week_of, actual, assigned)):
        if a != g:                       # 振替コマ
            sub_by[w_id][g].append(j)

    # 余分な振替（2 コマ目以降）を祝日週へスライド
    for w_id, day_map in list(sub_by.items()):
        for g, idxs in list(day_map.items()):
            for extra_idx in idxs[1:]:
                for tgt_w, is_hol in holiday_week_flag.items():
                    if not is_hol or tgt_w == w_id or g in sub_by[tgt_w]:
                        continue
                    # swap 候補：通常コマ(=振替でない)
                    cands = [j for j in range(len(dates))
                             if week_of[j] == tgt_w
                             and actual[j] == assigned[j]
                             and actual[j] != g]
                    if not cands:
                        continue
                    swap_j = cands[0]
                    assigned[swap_j], assigned[extra_idx] = assigned[extra_idx], assigned[swap_j]

                    sub_by[tgt_w][g] = [swap_j]
                    sub_by[w_id][g].remove(extra_idx)
                    break


def add_duplicate_shift_penalty(model, x,           # {(w,j): Var}
                                dates, actual_weekdays, date_to_weekid,
                                pen_dup = 800.0,     # 罰則係数
                                bigM    = 25):
    """
    shift_cnt[k,w] = その週 k に『実曜 ≠ w』なのに w に割り当てたコマ数
    dup[k,w]      = shift_cnt[k,w] が 2 以上なら 1 になるバイナリ
    目的関数に   pen_dup * Σ dup[k,w] を返す
    """
    # ―― 週 ID を 0..K-1 の連番に変換
    week_ids   = sorted({date_to_weekid[d] for d in dates})
    week_index = {wid: i for i, wid in enumerate(week_ids)}
    num_weeks  = len(week_ids)

    shift_cnt, dup = {}, {}
    for k in range(num_weeks):
        for w in range(5):                       # 0=Mon … 4=Fri
            shift_cnt[k, w] = model.addVar(
                vtype="I", lb=0, ub=len(dates),
                name=f"shift_{k}_{w}"
            )
            dup[k, w] = model.addVar(
                vtype="B", name=f"dup_{k}_{w}"
            )

            # === “曜日変更” になったコマの総和を定義 ==================
            js_in_kw = [j for j, d in enumerate(dates)
                        if week_index[date_to_weekid[d]] == k
                        and actual_weekdays[j] != w]  # ← “振替” のみ
            if js_in_kw:      # 週に候補が全く無い場合は作らない
                model.addCons(
                    shift_cnt[k, w] ==
                    quicksum(x[w, j] for j in js_in_kw),
                    name=f"defShift_{k}_{w}"
                )
            else:
                # コマ自体が存在しない週は shift_cnt = 0 と固定
                model.addCons(shift_cnt[k, w] == 0)

            # === 2 コマ以上なら dup=1 になるリンク制約 ================
            model.addCons(
                shift_cnt[k, w] - 1 <= bigM * dup[k, w],
                name=f"dupLink_{k}_{w}"
            )

    # === 目的関数に返す罰則項 ========================================
    return pen_dup * quicksum(dup[k, w]
                              for k in range(num_weeks)
                              for w in range(5))


def add_nonholiday_shift_penalty(model, x,
                                 dates,            # List[date]
                                 actual_wd,        # List[int]
                                 date_to_weekid,   # dict[date → (year,week)]
                                 pen_nonhol=5_000.0,
                                 manual_holidays:set[date]=set()):
    # ← 正しく holidays も渡す
    holiday_week = build_holiday_week_map_full(
        dates,
        jpholiday,
        manual_holidays
    )

    penalties = []
    for j, d in enumerate(dates):
        wid = date_to_weekid[d]
        if holiday_week[wid] == 0:
            for w in range(5):
                if w != actual_wd[j]:
                    penalties.append(x[w, j])
    return pen_nonhol * quicksum(penalties)



def build_holiday_week_map(dates):
    """各 ISO 週 (year, week) が祝日を含むなら 1 を返すマップ"""
    hol_map = {}
    for d in dates:
        wid = d.isocalendar()[:2]
        # 初期化
        hol_map.setdefault(wid, 0)
        if jpholiday.is_holiday(d):
            hol_map[wid] = 1
    return hol_map


# def balance_same_weekday_positions(dates,
#                                    actual_wd,
#                                    assigned_wd):
#     """
#     同じ曜日ラベルが出来るだけ等間隔になるように
#     近くのコマとスワップして調整する。

#     * 実曜≠ラベルの総数が増える交換はスキップ
#     * スワップは “同一 ISO 週” 内だけに限定
#       （週制約や重複チェックを壊さないため）
#     """
#     n = len(dates)
#     idx_by_label = defaultdict(list)
#     for idx, w in enumerate(assigned_wd):
#         idx_by_label[w].append(idx)

#     for w, idx_list in idx_by_label.items():
#         if len(idx_list) < 3:          # 1 区間では均す意味がない
#             continue
#         idx_list.sort()
#         for k in range(1, len(idx_list) - 1):
#             prev_i = idx_list[k-1]
#             cur_i  = idx_list[k]
#             next_i = idx_list[k+1]

#             target_i = (prev_i + next_i) // 2      # 数直線上の真ん中
#             if target_i == cur_i:
#                 continue                           # すでに中央ならスキップ
#             # 週内に限定して swap 候補を探す
#             wid_cur = dates[cur_i].isocalendar()[:2]
#             direction = 1 if target_i > cur_i else -1
#             j = cur_i
#             while 0 <= j < n and dates[j].isocalendar()[:2] == wid_cur:
#                 j += direction
#                 if j < 0 or j >= n:
#                     break
#                 # 「実曜≠ラベル」の総数が増えないかをチェック
#                 before = (actual_wd[cur_i] != assigned_wd[cur_i]) + \
#                          (actual_wd[j]     != assigned_wd[j])
#                 after  = (actual_wd[cur_i] != assigned_wd[j])   + \
#                          (actual_wd[j]     != assigned_wd[cur_i])
#                 if after > before:
#                     continue  # 悪化するスワップはしない
#                 # 交換実行
#                 assigned_wd[cur_i], assigned_wd[j] = assigned_wd[j], assigned_wd[cur_i]
#                 break    # 1 ステップ動かしたら次へ
# ------------------------------------------------------------
# 祝日週 × すでに曜日変更になっているコマだけを使って
# 曜日ごとの“真ん中寄せ”を行うユーティリティ
# ------------------------------------------------------------
def reposition_shifted_lessons_all(
    *,                       # -- 以下キーワード引数のみ --
    dates: list[date],       # 期間 25 日（or 30 日 …）の date 配列
    actual: list[int],       # dates[i].weekday()   (0=Mon … 4=Fri)
    assigned: list[int],     # 既に決まった週ラベル (0..4)
    holiday_week_flag: dict, # {(year,week): 0/1}
    lookback: int = 3,       # 探索ウインドウ（週数）
    lookahead: int = 3):

    # === 祝日を含む週に属するインデックス集合 =========================
    idx_in_holweek = {
        i for i, d in enumerate(dates)
        if holiday_week_flag.get(d.isocalendar()[:2], 0) == 1
    }

    # ---------------------------------------------------------
    # 曜日ごとに処理
    # ---------------------------------------------------------
    for target_wd in range(5):              # 0=月 … 4=金
        pos = [i for i, w in enumerate(assigned) if w == target_wd]
        if len(pos) < 3:
            continue                        # 真ん中が存在しない

        for idx in pos[1:-1]:               # 端を除く
            if actual[idx] == target_wd:
                continue                    # 元々その曜日→動かさない

            prev_i = max(p for p in pos if p < idx)
            next_i = min(p for p in pos if p > idx)
            ideal  = dates[prev_i] + (dates[next_i] - dates[prev_i]) / 2

            lo_day = dates[idx] - timedelta(days=7*lookback)
            hi_day = dates[idx] + timedelta(days=7*lookahead)

            # ---- swap 候補：祝日週かつ既に mismatch かつ別曜日ラベル ----
            cands = [j for j, d in enumerate(dates)
                     if j in idx_in_holweek
                     and lo_day <= d <= hi_day
                     and actual[j] != assigned[j]      # mismatch
                     and assigned[j] != target_wd]     # 別曜日

            if not cands:
                continue

            best = min(cands, key=lambda j: abs((dates[j]-ideal).days))

            # --- swap 実行（曜日変更総数はそのまま） ------------------
            assigned[idx], assigned[best] = assigned[best], assigned[idx]

# ── 全学期（春・秋）を対象に週フラグを一度だけ計算 ──
all_summer_dates = [
    start_date_summer + timedelta(days=i)
    for i in range((end_date_summer - start_date_summer).days + 1)
]
holiday_week_flag_summer = build_holiday_week_map_full(
    dates=all_summer_dates,
    jpholiday=jpholiday,
    manual_holidays=spring_holidays
)

all_autumn_dates = [
    start_date_autumn + timedelta(days=i)
    for i in range((end_date_autumn - start_date_autumn).days + 1)
]
holiday_week_flag_autumn = build_holiday_week_map_full(
    dates=all_autumn_dates,
    jpholiday=jpholiday,
    manual_holidays=autumn_holidays
)



def safe_to_date(cell) -> date:
    """
    Excel セルの値を安全に date 型に変換して返します。

    ・datetime.datetime → .date()
    ・datetime.date     → そのまま返却
    ・"YYYY-MM-DD" 形式の文字列 → datetime.strptime でパースして返却
    ・その他（None や不正フォーマットなど） → date.max を返却

    これにより、セルの型に依存せず比較や辞書キーとして利用可能な
    date 型に統一できます。
    """
    # datetime.datetime → date
    if isinstance(cell, datetime):
        return cell.date()

    # datetime.date → そのまま
    if isinstance(cell, date):
        return cell

    # 文字列 "YYYY-MM-DD" → パース
    if isinstance(cell, str):
        s = cell.strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
            try:
                return datetime.strptime(s, "%Y-%m-%d").date()
            except ValueError:
                pass

    # 上記以外は異常値として最大日付を返却
    return date.max


def build_week_penalties(
    week_dict,
    jpholiday,
    spring_holidays,
    autumn_holidays,
    reserve_limit=None,
    dup_dates=None,
    base_penalty: float = 1000.0,
    red_per_holiday: float = 200.0,
):
    """
    week_dict: {(year, week): [date,...]}
    base_penalty: 祝日ゼロ週のコスト
    red_per_holiday: 1つの祝日・休校日あたりの減少量
    """
    week_penalties = {}
    for week_id, days in week_dict.items():
        year, weeknum = week_id
        monday = date.fromisocalendar(year, weeknum, 1)
        holiday_count = sum(
            1
            for i in range(7)
            if (monday + timedelta(days=i)).weekday() < 5
               and (
                   jpholiday.is_holiday(monday + timedelta(days=i))
                   or (monday + timedelta(days=i)) in spring_holidays
                   or (monday + timedelta(days=i)) in autumn_holidays
               )
        )
        # 予備日超過のチェック
        if reserve_limit and any(d > reserve_limit for d in days):
            penalty = 10_000.0
        else:
            penalty = base_penalty - holiday_count * red_per_holiday
            penalty = max(0.0, penalty)
        week_penalties[week_id] = penalty
    return week_penalties





# === 春1最適化 ===
# def run_spring1_optimization(
#     spring_holidays,
#     autumn_holidays,
#     spring_buffer_count1,
#     base_penalty: float = 1000.0,
#     red_per_holiday: float = 200.0,
#     holiday_weight: float = 1.0,
#     pen_label_dup: float = 300.0,
# ):
#     """
#     春学期Aモジュール（25日分）に対する曜日割当を最適化します。

#     Args:
#         spring_holidays: set[date] 春学期の休講日
#         autumn_holidays: set[date] 秋学期の休講日（本関数内では使用しません）
#         spring_buffer_count1: int   Aモジュールの予備日数（本関数では返値には影響しません）

#     Returns:
#         assigned: List[int]  dates と同順の最適割当曜日 (0=月曜 … 4=金曜)
#     """
#     # # ── 1) 最適化対象の日付リスト（25日分）を取得 ───────────────
#     # # all_weekdays_summer は既に「春学期の平日かつ祝日・manual_holidays除外済み」のリスト
#     # dates = all_weekdays_summer[:25]
#     # num_days = len(dates)

#     # # ── 2) 各日付の実曜日インデックス取得 (0=月曜…6=日曜) ────────
#     # dow = [d.weekday() for d in dates]

#     # # ── 3) ISO 週 ID → 連番(1始まり) に変換 ──────────────────────
#     # week_ids = sorted({d.isocalendar()[:2] for d in dates})   # (年, 週番号) のタプルリスト
#     # week_index = {wid: i+1 for i, wid in enumerate(week_ids)}  # 連番マップ
#     # week_of = [week_index[d.isocalendar()[:2]] for d in dates]  # 各日が属する週連番
#     # num_weeks = len(week_ids)

#     # # ── 4) 各週の休講フラグ h を計算 ────────────────────────────
#     # # build_holiday_week_map_full は (year,week)→0/1 マップを返す既存ユーティリティ
#     # holiday_week = build_holiday_week_map_full(
#     #     dates=dates,
#     #     jpholiday=jpholiday,
#     #     manual_holidays=spring_holidays
#     # )
#     # # 週連番順に並べた 0/1 リスト
#     # h = [holiday_week[wid] for wid in week_ids]

#     # # ── 5) SCIP モデル構築 ──────────────────────────────────────
#     # model = Model("Spring1_ScheduleOpt")

#     # # 決定変数 x[i,j]: 日 j に曜日ラベル i を割当 (i=0..4, j=0..num_days-1)
#     # W = list(range(5))  # 曜日ラベル 0(月曜)～4(金曜)
#     # D = list(range(num_days))
#     # x = {(i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
#     #      for i in W for j in D}

#     # # ミスマッチ指標 y[j]: 実曜日と異なる割当なら 1
#     # y = {j: model.addVar(vtype="B", name=f"y_{j}") for j in D}

#     # # --- 制約 1: 1日1コマ割当 ---
#     # for j in D:
#     #     model.addCons(quicksum(x[i, j] for i in W) == 1)

#     # # --- 制約 2: 各曜日ラベルを必ず5回使う ---
#     # for i in W:
#     #     model.addCons(quicksum(x[i, j] for j in D) == 5)

#     # # --- 制約 3: y[j] の定義 (実曜と異なる割当=1) ---
#     # for j in D:
#     #     # 実曜以外に割当られたら y[j] ≥ x[i,j]
#     #     for i in W:
#     #         if i != dow[j]:
#     #             model.addCons(y[j] >= x[i, j])
#     #     # 実曜以外の合計以上 y[j] ≤ Σ x[i,j]
#     #     model.addCons(
#     #         y[j] <= quicksum(x[i, j] for i in W if i != dow[j])
#     #     )

#     # # --- 制約 4: 同一週内のミスマッチ上限 (週ごと最大1回) ---
#     # for w in range(1, num_weeks + 1):
#     #     idxs = [j for j in D if week_of[j] == w]
#     #     if idxs:
#     #         model.addCons(
#     #             quicksum(
#     #                 x[i, j]
#     #                 for j in idxs
#     #                 for i in W
#     #                 if i != dow[j]
#     #             ) <= 1
#     #         )

#     # # ── 6) 目的関数定義 ───────────────────────────────────────
#     # # Z1 = 全日ミスマッチ数, Z2 = 非祝日週でのミスマッチ数
#     # M = 1000
#     # Z1 = quicksum(y[j] for j in D)
#     # Z2 = quicksum((1 - h[week_of[j] - 1]) * y[j] for j in D)
#     # # 大きな重み M で Z1 を優先しつつ、Z2 で週の祝日有無を考慮
#     # model.setObjective(M * Z1 + Z2, "minimize")

#     # # ── 7) 最適化実行 ─────────────────────────────────────────
#     # model.optimize()
#     # if model.getStatus() != "optimal":
#     #     st.error("❌ 春1の最適化解が見つかりませんでした。")
#     #     return []

#     # # ── 8) 最適解から割当リストを抽出 ───────────────────────────
#     # assigned = [
#     #     next(i for i in W if model.getVal(x[i, j]) > 0.5)
#     #     for j in D
#     # ]


#     # 1) 最適化対象の日付（平日25日）を取得
#     dates = all_weekdays_summer[:25]
#     num_days = len(dates)
#     num_weeks = 5

#     # 2) 実際の曜日インデックスを取得 (0=月曜…6=日曜)
#     actual_weekdays = [d.weekday() for d in dates]

#     # 3) 曜日文字列リスト
#     weekdays_str = ['月曜', '火曜', '水曜', '木曜', '金曜']

#     # 4) 日付→ISO週IDと、週ごとの日付リストを作成
#     date_to_weekid = {}
#     week_dict = defaultdict(list)
#     for d in dates:
#         wid = d.isocalendar()[:2]  # (年, 週番号)
#         date_to_weekid[d] = wid
#         week_dict[wid].append(d)

#     # 5) 連続振替日の検出
#     global dup_dates_global
#     change_events = []
#     for name, ds in st.session_state.event_periods.items():
#         if "曜日変更" in name:
#             m = re.search(r"\((月|火|水|木|金)\)", name)
#             if m:
#                 wd = "月火水木金".index(m.group(1))
#                 change_events += [(d, wd) for d in ds]
#     dup_dates_global = detect_dup_shift_days(change_events)

#     # 6) 週ごとのペナルティを計算
#     reserve_limit = end_date_summer  # 春学期終了日を超える週は高ペナルティ
#     week_penalties = build_week_penalties(
#         week_dict,
#         jpholiday=jpholiday,
#         spring_holidays=spring_holidays,
#         autumn_holidays=autumn_holidays,
#         reserve_limit=reserve_limit,
#         dup_dates=dup_dates_global,
#         base_penalty=base_penalty,
#         red_per_holiday=red_per_holiday,
#     )

#     # 7) SCIPモデル構築
#     model = Model("Spring1_Optimization")
#     x = {(i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
#          for i in range(num_weeks) for j in range(num_days)}

#     # 8) 目的関数項: 実曜≠割当曜 × 週ペナルティ
#     obj_mismatch = quicksum(
#         (1 if i != actual_weekdays[j] else 0)
#         * week_penalties[date_to_weekid[dates[j]]]
#         * x[i, j]
#         for i in range(num_weeks)
#         for j in range(num_days)
#     )
#     obj_holiday = holiday_weight * obj_mismatch

#     # # 9) 隣接日の重複シフトペナルティ
#     # z_adj = {}
#     # for i in range(num_weeks):
#     #     for j in range(num_days - 1):
#     #         if (dates[j+1] - dates[j]).days == 1:
#     #             z_adj[i, j] = model.addVar(vtype="B", name=f"z_adj_{i}_{j}")
#     #             model.addCons(z_adj[i, j] <= x[i, j])
#     #             model.addCons(z_adj[i, j] <= x[i, j+1])
#     #             model.addCons(z_adj[i, j] >= x[i, j] + x[i, j+1] - 1)
#     # penalty_adj = pen_adj * quicksum(z for z in z_adj.values())

#     # 10) 同じ週で同ラベル振替が複数回ある場合のペナルティ
#     v_label = {}
#     # for i in range(num_weeks):
#     for wid, days in week_dict.items():
#         idxs = [j for j, d in enumerate(dates) if date_to_weekid[d] == wid]
#         for label in range(num_weeks):#for i in range(num_weeks):
#             for p in range(len(idxs)):
#                 for q in range(p + 1, len(idxs)):
#                     j1, j2 = idxs[p], idxs[q]
#                     v_label[label, j1, j2] = model.addVar(vtype="B", name=f"v_lbl_{label}_{j1}_{j2}")
#                     model.addCons(v_label[label, j1, j2] <= x[label, j1])
#                     model.addCons(v_label[label, j1, j2] <= x[label, j2])
#                     model.addCons(v_label[label, j1, j2] >= x[label, j1] + x[label, j2] - 1)
#     penalty_label_dup = pen_label_dup * quicksum(v for v in v_label.values())


#     # # 11) 非祝日週ミスマッチペナルティ
#     # hol_map = {
#     #     wid: any(jpholiday.is_holiday(d) or d in spring_holidays or d in autumn_holidays for d in days)
#     #     for wid, days in week_dict.items()
#     # }
#     # penalty_nonhol = quicksum(
#     #     pen_nonhol
#     #     * (1 if i != actual_weekdays[j] else 0)
#     #     * (0 if hol_map[date_to_weekid[dates[j]]] else 1)
#     #     * x[i, j]
#     #     for i in range(num_weeks) for j in range(num_days)
#     # )

#     # # 11) 目的関数の設定
#     # model.setObjective(obj_mismatch +  penalty_label_dup , "minimize")

#     # 12) 制約: 各日jに1回だけ割当
#     for j in range(num_days):
#         model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
#     # 13) 制約: 各週iに5日ずつ割当
#     for i in range(num_weeks):
#         model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)



#     # 11) 目的関数の設定
#     model.setObjective(obj_holiday +  penalty_label_dup , "minimize")

#     # 14) 最適化実行
#     model.optimize()
#     if model.getStatus() != "optimal":
#         st.error("❌ 春1の最適解が見つかりませんでした。")
#         return []

#     # 15) 割当結果の抽出
#     assigned = [
#         next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
#         for j in range(num_days)
#     ]

#     # 16) 距離ベースの平準化
#     global positions_by_weekday
#     positions_by_weekday = build_positions(actual_weekdays)
#     runs, cur = [], []
#     for idx in range(num_days):
#         if assigned[idx] != actual_weekdays[idx]:
#             cur.append(idx)
#         else:
#             if cur:
#                 runs.append(cur)
#                 cur = []
#     if cur:
#         runs.append(cur)
#     for run in runs:
#         normalize_run(run, dates, actual_weekdays, assigned)

def run_spring1_optimization(
    spring_holidays,
    autumn_holidays,
    spring_buffer_count1,
    base_penalty: float = 1000.0,
    red_per_holiday: float = 200.0,
    holiday_weight: float = 1.0,
    pen_label_dup: float = 300.0,
    gap_weight: float = 1,
):
    """
    春学期Aモジュールの曜日割当を最適化
    Returns: assigned list of labels (0=月曜...4=金曜)
    """
    # 1) 対象25日取得
    dates = all_weekdays_summer[:25]
    num_days = len(dates)
    num_weeks = 5

    # 2) 実曜日インデックス
    actual_weekdays = [d.weekday() for d in dates]

    # 曜日文字列リスト（Excel出力などで使用）
    weekdays_str = ['月曜', '火曜', '水曜', '木曜', '金曜']

    # 3) 週IDマップ
    date_to_weekid = {}
    week_dict = defaultdict(list)
    for d in dates:
        wid = d.isocalendar()[:2]
        date_to_weekid[d] = wid
        week_dict[wid].append(d)

    # 4) change_events から mis_idxs (振替対象インデックス) を取得
    change_events = []
    for name, ds in st.session_state.event_periods.items():
        if "曜日変更" in name:
            m = re.search(r"\((月|火|水|木|金)\)", name)
            if m:
                wd = "月火水木金".index(m.group(1))
                change_events += [(d, wd) for d in ds]
    mis_idxs = sorted(
        j for j, d in enumerate(dates)
        if any(ev_d == d for ev_d, _ in change_events)
    )

    # 5) runs (連続区間) に分割
    runs = []
    cur = []
    for j in mis_idxs:
        if not cur or j == cur[-1] + 1:
            cur.append(j)
        else:
            runs.append(cur)
            cur = [j]
    if cur:
        runs.append(cur)

    # 6) 週ペナルティ
    reserve_limit = end_date_summer
    week_penalties = build_week_penalties(
        week_dict,
        jpholiday=jpholiday,
        spring_holidays=spring_holidays,
        autumn_holidays=autumn_holidays,
        reserve_limit=reserve_limit,
        dup_dates=dup_dates_global,
        base_penalty=base_penalty,
        red_per_holiday=red_per_holiday,
    )

    # 7) モデル構築
    model = Model("Spring1_Optimization")
    x = {(i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
         for i in range(num_weeks) for j in range(num_days)}

    # 8) 祝日優先ミスマッチ
    obj_mismatch = quicksum(
        (1 if i != actual_weekdays[j] else 0)
        * week_penalties[date_to_weekid[dates[j]]]
        * x[i, j]
        for i in range(num_weeks) for j in range(num_days)
    )
    obj_holiday = holiday_weight * obj_mismatch

    # 9) 同ラベル重複ペナルティ
    v_label = {}
    for wid, _ in week_dict.items():
        idxs = [j for j, d in enumerate(dates)
                if date_to_weekid[d] == wid]
        for label in range(num_weeks):
            for p in range(len(idxs)):
                for q in range(p + 1, len(idxs)):
                    j1, j2 = idxs[p], idxs[q]
                    v_label[label, j1, j2] = model.addVar(vtype="B", name=f"v_lbl_{label}_{j1}_{j2}")
                    model.addCons(v_label[label, j1, j2] <= x[label, j1])
                    model.addCons(v_label[label, j1, j2] <= x[label, j2])
                    model.addCons(v_label[label, j1, j2] >= x[label, j1] + x[label, j2] - 1)
    penalty_label_dup = pen_label_dup * quicksum(v for v in v_label.values())

    # 10) ギャップ変数 u と制約
    u = {}
    for r, idx_list in enumerate(runs):
        wd = actual_weekdays[idx_list[0]]
        p_r = max(k for k in range(idx_list[0]) if actual_weekdays[k] == wd)
        n_r = min(k for k in range(idx_list[-1]+1, num_days)
                  if actual_weekdays[k] == wd)
        t_r = (p_r + n_r) / 2
        for j in idx_list:
            u[r, j] = model.addVar(vtype="C", name=f"u_{r}_{j}")
            model.addCons(u[r, j] >= j - t_r)
            model.addCons(u[r, j] >= t_r - j)
    gap_cost = gap_weight * quicksum(u[r, j] for r, j in u)

    # 11) 制約: 各日1 & 各週5
    for j in range(num_days):
        model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
    for i in range(num_weeks):
        model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)

    # 12) 目的関数 & solve
    model.setObjective(obj_holiday + penalty_label_dup + gap_cost, "minimize")
    model.optimize()

    # 13) 割当抽出
    if model.getStatus() != "optimal":
        st.error("❌ 春1の最適解が見つかりませんでした。")
        return []
    assigned = [
        next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
        for j in range(num_days)
    ]


    # 17) 結果をExcelに書き込む
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "春学期最適化結果"
    sheet["A1"] = "日付"
    sheet["B1"] = "実曜日"
    sheet["C1"] = ""
    sheet["D1"] = "割り当て"
    sheet["F1"] = "日付"
    sheet["G1"] = "実曜日"
    sheet["R1"] = "Aモジュール予備日"
    sheet["I1"] = "割り当て結果"

    weekday_counts = [0] * num_weeks
    row_map = {}
    for row_idx, d in enumerate(all_weekdays_summer, start=2):
        day_str = d.strftime('%Y-%m-%d')
        weekday_label = weekdays_str[d.weekday()]
        sheet[f"A{row_idx}"] = day_str
        sheet[f"B{row_idx}"] = weekday_label
        row_map[day_str] = row_idx

    for j, d in enumerate(dates):
        w = assigned[j]
        weekday_counts[w] += 1
        label = f"{weekdays_str[w]}{weekday_counts[w]}"
        row = row_map[d.strftime('%Y-%m-%d')]
        sheet[f"D{row}"] = label

    # 未割当日（D列が空欄）
    unassigned_dates = []
    for date_val in all_weekdays_summer:
        day_str = date_val.strftime('%Y-%m-%d')
        row = row_map.get(day_str)
        if row and sheet[f"D{row}"].value in (None, ""):
            unassigned_dates.append((date_val, weekdays_str[date_val.weekday()]))

    if unassigned_dates:
        sorted_unassigned = sorted(unassigned_dates, key=lambda x: x[0])
        top_r_dates = sorted_unassigned[:spring_buffer_count1]
        for idx, (r_date, _) in enumerate(top_r_dates):
            sheet[f"R{idx + 2}"] = r_date.strftime('%Y-%m-%d')
        row_out = 2
        for un_date, un_week in sorted_unassigned[spring_buffer_count1:]:
            sheet[f"F{row_out}"] = un_date.strftime('%Y-%m-%d')
            sheet[f"G{row_out}"] = un_week
            row_out += 1
    else:
        sheet["R2"] = "エラー"
        sheet["F2"] = "エラー"

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    st.session_state["spring1_excel_bytes"] = output

    return assigned



#-----------------------------春2-------------------
def normalize_penalties(raw_penalties: dict[tuple, float]) -> dict[tuple, float]:
    """
    build_week_penalties の結果（週ID→ペナルティ）を
    必要に応じてスケール・正規化します。
    現状はそのまま返します。
    """
    return raw_penalties



def run_spring2_optimization(spring_buffer_count2, spring_holidays, autumn_holidays):
    """
    春学期Bモジュール（25日分）に対する曜日割当を最適化します。

    Args:
        spring_buffer_count2: int   Bモジュールの予備日数（本関数では返値には影響しません）
        spring_holidays:   set[date] 春学期の休講日
        autumn_holidays:   set[date] 秋学期の休講日（本関数内では使用しません）

    Returns:
        assigned: List[int] dates と同順の最適割当曜日 (0=月曜 … 4=金曜)
    """
    # st.markdown("---")

    # ── 1) 春1の結果Excelがなければエラー
    if "spring1_excel_bytes" not in st.session_state:
        st.error("春1の結果がありません。")
        return []

    # ── 2) Excel から候補日の読み込み（F列）─────────────────────────
    data = io.BytesIO(st.session_state["spring1_excel_bytes"].getvalue())
    wb   = load_workbook(data)
    sheet= wb.active

    weekdays = []
    for row in range(2, sheet.max_row + 1):
        val = sheet[f"F{row}"].value
        # _safe_to_date は先頭で定義済みとする
        d = safe_to_date(val)
        if d != date.max:
            weekdays.append(d)

    if len(weekdays) < 25:
        st.warning("春2に使える日が足りません。")
        return []
    dates = weekdays[:25]
    num_days  = len(dates)
    num_weeks = 5

    # ── 3) 実曜日インデックス取得 & 曜日文字列 ───────────────────────
    actual_wd    = [d.weekday() for d in dates]
    weekdays_str = ['月曜','火曜','水曜','木曜','金曜']

    # ── 4) ISO週ID → 週辞書を作成 & reserve_limit を取得 ──────────────
    date_to_wid = {}
    week_dict   = defaultdict(list)
    for d in dates:
        wid = d.isocalendar()[:2]
        date_to_wid[d] = wid
        week_dict[wid].append(d)

    # T2セルから reserve_limit を取得（パース失敗時は date.max）
    rl = sheet["T2"].value
    if isinstance(rl, datetime):
        reserve_limit = rl.date()
    elif isinstance(rl, date):
        reserve_limit = rl
    else:
        try:
            reserve_limit = datetime.strptime(str(rl), "%Y-%m-%d").date()
        except:
            reserve_limit = date.max

    # ── 5) 日付重複振替日の検出（global dup_dates_global）──────────────
    global dup_dates_global
    change_events = []
    for name, ds in st.session_state.event_periods.items():
        if "曜日変更" in name:
            m = re.search(r"\((月|火|水|木|金)\)", name)
            if m:
                wd = "月火水木金".index(m.group(1))
                change_events += [(d, wd) for d in ds]
    dup_dates_global = detect_dup_shift_days(change_events)

    # ── 6) ペナルティ計算 ────────────────────────────────────────
    raw_pen = build_week_penalties(
        week_dict,
        jpholiday=jpholiday,
        spring_holidays=spring_holidays,
        autumn_holidays=autumn_holidays,
        reserve_limit=reserve_limit,
        dup_dates=dup_dates_global
    )
    # normalize_penalties は既存ユーティリティとする
    penalties = normalize_penalties(raw_pen)

    # ── 7) SCIPモデル構築 ────────────────────────────────────────
    model = Model("Spring2_Optimization")
    x = {(i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
         for i in range(num_weeks) for j in range(num_days)}

    # 目的：実曜≠割当曜 × 週ペナルティ
    model.setObjective(
        quicksum(
            (1 if i != actual_wd[j] else 0)
            * penalties[date_to_wid[dates[j]]]
            * x[i, j]
            for i in range(num_weeks)
            for j in range(num_days)
        ),
        "minimize"
    )

    # ── 8) 制約：1日1回割当 & 各週5日ずつ ─────────────────────────
    for j in range(num_days):
        model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
    for i in range(num_weeks):
        model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)

    # ── 9) 最適化実行 ───────────────────────────────────────────
    model.optimize()
    if model.getStatus() != "optimal":
        st.error("❌ 春2の最適解が見つかりませんでした。")
        return []

    # ── 10) 割当結果の抽出 ──────────────────────────────────────
    assigned = [
        next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
        for j in range(num_days)
    ]

    # ── 11) 距離ベースの平準化 ───────────────────────────────────
    global positions_by_weekday
    positions_by_weekday = build_positions(actual_wd)

    runs, cur = [], []
    for idx in range(num_days):
        if assigned[idx] != actual_wd[idx]:
            cur.append(idx)
        else:
            if cur:
                runs.append(cur); cur = []
    if cur:
        runs.append(cur)
    for run in runs:
        normalize_run(run, dates, actual_wd, assigned)

    # ── 12) ラベル書き込み & 未割当日の出力 ───────────────────────
    weekday_counts = [5] * num_weeks  # 春1の続きとして
    # Bモジュール割当ラベルを I列へ
    sheet["T1"] = "Bモジュール予備日"
    sheet["K1"] = "日付"
    sheet["L1"] = "実曜日"
    sheet["N1"] = "割り当て結果"

    for j, d in enumerate(dates):
        w = assigned[j]
        weekday_counts[w] += 1
        label = f"{weekdays_str[w]}{weekday_counts[w]}"
        # F列の該当行を探して I列に書き込む
        for row in range(2, sheet.max_row + 1):
            if safe_to_date(sheet[f"F{row}"].value) == d:
                sheet[f"I{row}"] = label
                break

    # 未割当日のリスト化と書き込み
    unassigned_dates = []
    for row_idx in range(2, sheet.max_row + 1):
        f_val = sheet[f"F{row_idx}"].value
        if f_val and not sheet[f"I{row_idx}"].value:
            dt = safe_to_date(f_val)
            if dt != date.max:
                unassigned_dates.append((dt, weekdays_str[dt.weekday()]))

    if unassigned_dates:
        sorted_un = sorted(unassigned_dates, key=lambda x: x[0])
        top_t = sorted_un[:spring_buffer_count2]
        for idx, (td, _) in enumerate(top_t):
            sheet[f"T{idx+2}"] = td.strftime("%Y-%m-%d")
        row_out = 2
        for dt, wd in sorted_un[spring_buffer_count2:]:
            sheet[f"K{row_out}"] = dt.strftime("%Y-%m-%d")
            sheet[f"L{row_out}"] = wd
            row_out += 1
    else:
        sheet["T2"] = "エラー"
        sheet["K2"] = "エラー"

    # ── 13) Excel をバイトストリームに保存 ───────────────────────
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    st.session_state["spring2_excel_bytes"] = output

    return assigned


#-----------------------春3----------------------------

def run_spring3_optimization(spring_buffer_count3, spring_holidays, autumn_holidays):
    """
    春学期Cモジュール（25日分）に対する曜日割当を最適化し、Excelに書き込みます。

    Args:
        spring_buffer_count3: int   Cモジュールの予備日数（出力セル数に影響）
        spring_holidays: set[date] 春学期の休講日
        autumn_holidays: set[date] 秋学期の休講日（本関数内では使用しません）

    Returns:
        assigned: List[int]  dates と同順の最適割当曜日 (0=月曜 … 4=金曜)
    """
    # st.markdown("---")

    # ── 1) 春2の結果がなければエラー
    if "spring2_excel_bytes" not in st.session_state:
        st.error("春2までの最適化結果が見つかりません。")
        return []

    # ── 2) Excel から未割当日 (K列) を安全に読み込み
    data = st.session_state["spring2_excel_bytes"]
    data.seek(0)
    wb = load_workbook(data)
    sheet = wb.active

    weekdays = []
    for row in range(2, sheet.max_row + 1):
        cell = sheet[f"K{row}"].value
        d = safe_to_date(cell)
        if d != date.max:
            weekdays.append(d)

    if len(weekdays) < 25:
        st.warning("春3に使える未割当日が25日未満です。")
        return []
    dates = weekdays[:25]
    num_days = len(dates)
    num_weeks = 5

    # ── 3) 実曜日インデックスと曜日文字列
    actual_weekdays = [d.weekday() for d in dates]
    weekdays_str = ['月曜', '火曜', '水曜', '木曜', '金曜']

    # ── 4) ISO週ID → 週ごとの日付リストを構築
    date_to_weekid = {}
    week_dict = defaultdict(list)
    for d in dates:
        wid = d.isocalendar()[:2]
        date_to_weekid[d] = wid
        week_dict[wid].append(d)

    # ── 5) 連続振替日の検出
    global dup_dates_global
    change_events = []
    for name, ds in st.session_state.event_periods.items():
        if "曜日変更" in name:
            m = re.search(r"\((月|火|水|木|金)\)", name)
            if m:
                wd = "月火水木金".index(m.group(1))
                change_events += [(d, wd) for d in ds]
    dup_dates_global = detect_dup_shift_days(change_events)

    # ── 6) reserve_limit を V2セルから取得（安全パース）
    rl_cell = sheet["V2"].value
    rl_date = safe_to_date(rl_cell)
    # date.max は学期末まで許可
    reserve_limit = rl_date if rl_date != date.max else date.max

    # ── 7) 週ごとのペナルティ計算
    week_penalties = build_week_penalties(
        week_dict,
        jpholiday=jpholiday,
        spring_holidays=spring_holidays,
        autumn_holidays=autumn_holidays,
        reserve_limit=reserve_limit,
        dup_dates=dup_dates_global
    )

    # ── 8) SCIPモデル構築
    model = Model("Spring3_Optimization")
    x = {(i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
         for i in range(num_weeks) for j in range(num_days)}

    # 目的：ズレ × 週ペナルティの合計を最小化
    model.setObjective(
        quicksum(
            (1 if i != actual_weekdays[j] else 0)
            * week_penalties[date_to_weekid[dates[j]]]
            * x[i, j]
            for i in range(num_weeks)
            for j in range(num_days)
        ),
        "minimize"
    )

    # ── 9) 制約：各日 j に 1 回だけ割当
    for j in range(num_days):
        model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
    # ── 10) 制約：各週 i に 5 日ずつ割当
    for i in range(num_weeks):
        model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)

    # ── 11) 最適化実行
    model.optimize()
    if model.getStatus() != "optimal":
        st.error("❌ 春3の最適解が見つかりませんでした。")
        return []

    # ── 12) 割当結果の抽出
    assigned = [
        next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
        for j in range(num_days)
    ]

    # ── 13) 距離ベースの平準化
    positions_by_weekday = build_positions(actual_weekdays)
    runs, cur = [], []
    for idx in range(num_days):
        if assigned[idx] != actual_weekdays[idx]:
            cur.append(idx)
        else:
            if cur:
                runs.append(cur)
                cur = []
    if cur:
        runs.append(cur)
    for run in runs:
        normalize_run(run, dates, actual_weekdays, assigned)

    # ── 14) 結果をExcelに書き込み
    weekday_counts = [10] * num_weeks  # 春1(5)+春2(5)の続き
    sheet["V1"] = "Cモジュールの予備日"

    # N列に割当ラベルを記入
    for j, d in enumerate(dates):
        w = assigned[j]
        weekday_counts[w] += 1
        label = f"{weekdays_str[w]}{weekday_counts[w]}"
        for row in range(2, sheet.max_row + 1):
            k_cell = sheet[f"K{row}"].value
            if safe_to_date(k_cell) == d:
                sheet[f"N{row}"] = label
                break

    # V列に未割当（N列空欄のK列）を書き込み
    unassigned_dates = []
    for row in range(2, sheet.max_row + 1):
        if not sheet[f"N{row}"].value and sheet[f"K{row}"].value:
            dt = safe_to_date(sheet[f"K{row}"].value)
            if dt != date.max:
                unassigned_dates.append(dt)

    if unassigned_dates:
        for idx, d in enumerate(sorted(unassigned_dates)[:spring_buffer_count3]):
            sheet[f"V{idx+2}"] = d.strftime("%Y-%m-%d")
    else:
        sheet["V2"] = "未割当なし"

    # ── 15) バイトストリームに保存
    output3 = io.BytesIO()
    wb.save(output3)
    output3.seek(0)
    st.session_state["spring3_excel_bytes"] = output3
    st.session_state["spring3_done"] = True

    return assigned








# === 秋学期の平日を抽出 ===
# def run_autumn1_optimization(spring_holidays, autumn_holidays, autumn_buffer_count1):
#     """
#     秋学期Aモジュール（25日分）の曜日割当を最適化し、Excel に書き込みます。

#     Args:
#         spring_holidays: set[date] 春学期の休講日（本関数内では使用しません）
#         autumn_holidays: set[date] 秋学期の休講日
#         autumn_buffer_count1: int Aモジュールの予備日数（Excel出力に影響）
#     Returns:
#         assigned: List[int] dates と同順の最適割当曜日 (0=月曜 … 4=金曜)
#     """
#     # 1) 秋学期平日リストを作成（祝日・土日・manual_holidays除外）
#     autumn_date_list = [
#         start_date_autumn + timedelta(days=i)
#         for i in range((end_date_autumn - start_date_autumn).days + 1)
#     ]
#     all_weekdays_autumn = [
#         d for d in autumn_date_list
#         if d.weekday() < 5
#         and not jpholiday.is_holiday(d)
#         and d not in autumn_holidays
#     ]
#     if len(all_weekdays_autumn) < 25:
#         st.warning("秋1に使える日が25日未満です。")
#         return []

#     # 2) 最適化対象の先頭25日を切り出し
#     dates     = all_weekdays_autumn[:25]
#     num_days  = len(dates)
#     num_weeks = 5
#     actual_weekdays = [d.weekday() for d in dates]
#     weekdays_str   = ['月曜','火曜','水曜','木曜','金曜']

#     # 3) ISO週ID→辞書と、週ごとの日付リストを作成
#     date_to_weekid = {}
#     week_dict      = defaultdict(list)
#     for d in dates:
#         wid = d.isocalendar()[:2]  # (year, week)
#         date_to_weekid[d] = wid
#         week_dict[wid].append(d)

#     # 4) 連続振替日の検出
#     global dup_dates_global
#     change_events = []
#     for name, ds in st.session_state.event_periods.items():
#         if "曜日変更" in name:
#             m = re.search(r"\((月|火|水|木|金)\)", name)
#             if m:
#                 wd = "月火水木金".index(m.group(1))
#                 change_events += [(d, wd) for d in ds]
#     dup_dates_global = detect_dup_shift_days(change_events)

#     # 5) reserve_limit（秋学期末）を設定
#     reserve_limit = end_date_autumn

#     # 6) 週ごとのペナルティを計算
#     week_penalties = build_week_penalties(
#         week_dict,
#         jpholiday=jpholiday,
#         spring_holidays=spring_holidays,
#         autumn_holidays=autumn_holidays,
#         reserve_limit=reserve_limit,
#         dup_dates=dup_dates_global
#     )

#     # 7) SCIP モデル構築
#     model = Model("Autumn1_Optimization")
#     x = {
#         (i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
#         for i in range(num_weeks) for j in range(num_days)
#     }

#     # 8) 目的関数：実曜≠割当曜 × 週ペナルティ ＋ 連続罰
#     base_obj = quicksum(
#         (1 if i != actual_weekdays[j] else 0)
#         * week_penalties[date_to_weekid[dates[j]]]
#         * x[i, j]
#         for i in range(num_weeks) for j in range(num_days)
#     )

#     # 連続罰（隣接日で同ラベルなら z=1）
#     ADJ_PENALTY = 5.0
#     z = {}
#     for i in range(num_weeks):
#         for j in range(num_days - 1):
#             if (dates[j+1] - dates[j]).days == 1:
#                 z[i, j] = model.addVar(vtype="B", name=f"z_{i}_{j}")
#                 model.addCons(z[i, j] <= x[i, j])
#                 model.addCons(z[i, j] <= x[i, j+1])
#                 model.addCons(z[i, j] >= x[i, j] + x[i, j+1] - 1)
#     adj_obj = quicksum(ADJ_PENALTY * z_var for z_var in z.values())

#     model.setObjective(base_obj + adj_obj, "minimize")

#     # 9) 制約：各日1回割当 & 各週5日ずつ
#     for j in range(num_days):
#         model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
#     for i in range(num_weeks):
#         model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)

#     # 10) 最適化実行
#     model.optimize()
#     if model.getStatus() != "optimal":
#         st.error("❌ 秋1の最適解が見つかりませんでした。")
#         return []

#     # 11) 割当取得
#     assigned = [
#         next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
#         for j in range(num_days)
#     ]

#     # 12) 距離ベースの平準化
#     global positions_by_weekday
#     positions_by_weekday = build_positions(actual_weekdays)
#     runs, cur = [], []
#     for idx in range(num_days):
#         if assigned[idx] != actual_weekdays[idx]:
#             cur.append(idx)
#         else:
#             if cur:
#                 runs.append(cur); cur = []
#     if cur:
#         runs.append(cur)
#     for run in runs:
#         normalize_run(run, dates, actual_weekdays, assigned)


def run_autumn1_optimization(
    spring_holidays,
    autumn_holidays,
    autumn_buffer_count1,
    base_penalty: float = 1000.0,
    red_per_holiday: float = 200.0,
    adj_weight: float = 5.0,
    gap_weight: float = 1.0,
):
    """
    秋学期Aモジュール（25日分）の曜日割当を最適化し、Excel に書き込みます。

    Args:
        spring_holidays: set[date] 春学期の休講日（本関数内では使用しません）
        autumn_holidays: set[date] 秋学期の休講日
        autumn_buffer_count1: int Aモジュールの予備日数（Excel出力に影響）
        base_penalty: float 祝日ゼロ週のコスト
        red_per_holiday: float 1つの祝日・休校日あたり減少量
        adj_weight: float 隣接シフト重複のペナルティ重み
        gap_weight: float 前後授業間の中点からのズレに対する重み
    Returns:
        assigned: List[int] dates と同順の最適割当曜日 (0=月曜 … 4=金曜)
    """
    # 1) 秋学期平日リストを作成（祝日・土日・manual_holidays除外）
    autumn_date_list = [
        start_date_autumn + timedelta(days=i)
        for i in range((end_date_autumn - start_date_autumn).days + 1)
    ]
    all_weekdays_autumn = [
        d for d in autumn_date_list
        if d.weekday() < 5
        and not jpholiday.is_holiday(d)
        and d not in autumn_holidays
    ]
    if len(all_weekdays_autumn) < 25:
        st.warning("秋1に使える日が25日未満です。")
        return []

    # 2) 最適化対象の先頭25日を切り出し
    dates = all_weekdays_autumn[:25]
    num_days = len(dates)
    num_weeks = 5

    # 3) 実際の曜日インデックス & 曜日文字列リスト
    actual_weekdays = [d.weekday() for d in dates]
    weekdays_str = ['月曜','火曜','水曜','木曜','金曜']

    # 4) ISO週ID→辞書と、週ごとの日付リストを作成
    date_to_weekid = {}
    week_dict = defaultdict(list)
    for d in dates:
        wid = d.isocalendar()[:2]
        date_to_weekid[d] = wid
        week_dict[wid].append(d)

    # 5) 連続振替日の検出
    change_events = []
    for name, ds in st.session_state.event_periods.items():
        if "曜日変更" in name:
            m = re.search(r"\((月|火|水|木|金)\)", name)
            if m:
                wd = "月火水木金".index(m.group(1))
                change_events += [(d, wd) for d in ds]

    # mis_idxs: 振替予定日インデックス
    mis_idxs = sorted(
        j for j, d in enumerate(dates)
        if any(ev_d == d for ev_d, _ in change_events)
    )

    # runs: 連続したミスマッチ区間
    runs = []
    cur = []
    for j in mis_idxs:
        if not cur or j == cur[-1] + 1:
            cur.append(j)
        else:
            runs.append(cur)
            cur = [j]
    if cur:
        runs.append(cur)

    # 6) 週ペナルティ計算
    reserve_limit = end_date_autumn
    week_penalties = build_week_penalties(
        week_dict,
        jpholiday=jpholiday,
        spring_holidays=spring_holidays,
        autumn_holidays=autumn_holidays,
        reserve_limit=reserve_limit,
        dup_dates=detect_dup_shift_days(change_events),
        base_penalty=base_penalty,
        red_per_holiday=red_per_holiday,
    )

    # 7) SCIP モデル構築
    model = Model("Autumn1_Optimization")
    x = {
        (i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
        for i in range(num_weeks) for j in range(num_days)
    }

    # 8) obj_base: ミスマッチ×週ペナルティ
    obj_base = quicksum(
        (1 if i!=actual_weekdays[j] else 0)
        * week_penalties[ date_to_weekid[dates[j]] ]
        * x[i,j]
        for i in range(num_weeks) for j in range(num_days)
    )

    # 9) obj_nonhol: 非祝日週に置かれたミスマッチを直接罰
    holiday_map = {
        wid: any(jpholiday.is_holiday(d) or d in autumn_holidays for d in days)
        for wid, days in week_dict.items()
    }
    C_nonhol = 800.0  # 1コマあたりの罰則、調整可
    obj_nonhol = quicksum(
        C_nonhol
        * (1 if i!=actual_weekdays[j] else 0)
        * (1 - holiday_map[ date_to_weekid[dates[j]] ])
        * x[i,j]
        for i in range(num_weeks) for j in range(num_days)
    )

    # 10) obj_dup: 同じ週で同ラベルの振替が複数回ある場合のペナルティ
    v_label = {}
    for wid, days in week_dict.items():
        idxs = [j for j,d in enumerate(dates) if date_to_weekid[d]==wid]
        for label in range(num_weeks):
            for p in range(len(idxs)):
                for q in range(p+1, len(idxs)):
                    j1, j2 = idxs[p], idxs[q]
                    v_label[label,j1,j2] = model.addVar(vtype="B", name=f"v_lbl_{label}_{j1}_{j2}")
                    model.addCons(v_label[label,j1,j2] <= x[label,j1])
                    model.addCons(v_label[label,j1,j2] <= x[label,j2])
                    model.addCons(v_label[label,j1,j2] >= x[label,j1] + x[label,j2] - 1)
    pen_label_dup = 300.0  # 調整可
    obj_dup = pen_label_dup * quicksum(v for v in v_label.values())

    # 11) obj_gap: ギャップペナルティ（前後授業間中点からのズレ）
    u = {}
    for r, run in enumerate(runs):
        wd = actual_weekdays[run[0]]
        p_r = max(k for k in range(run[0]) if actual_weekdays[k]==wd)
        n_r = min(k for k in range(run[-1]+1, num_days) if actual_weekdays[k]==wd)
        t_r = (p_r + n_r)/2
        for j in run:
            u[r,j] = model.addVar(vtype="C", name=f"u_{r}_{j}")
            model.addCons(u[r,j] >= j - t_r)
            model.addCons(u[r,j] >= t_r - j)
    gamma = 0.1  # 調整可
    obj_gap = gamma * quicksum(u_var for u_var in u.values())

    # 11) 制約：各日1回 & 各週5日ずつ
    for j in range(num_days):
        model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
    for i in range(num_weeks):
        model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)

    # 12) 目的関数設定 & 最適化
    model.setObjective(obj_base + obj_nonhol + obj_dup + obj_gap, "minimize")
    model.optimize()

    # 13) 割当取得
    if model.getStatus() != "optimal":
        st.error("❌ 秋1の最適解が見つかりませんでした。")
        return []
    assigned = [
        next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
        for j in range(num_days)
    ]
    
    # 13) Excel 書き込み（ループ外で一度だけ）
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "秋1最適化結果"

    # A〜B列：日付と実曜日、D列：割当、F〜G：未割当、R列：早い未割当
    sheet["A1"] = "日付"
    sheet["B1"] = "実曜日"
    sheet["D1"] = "割り当て結果"
    sheet["F1"] = "日付"
    sheet["G1"] = "実曜日"
    sheet["R1"] = "Aモジュール予備日"
    sheet["I1"] = "割り当て結果"

    row_map = {}
    for idx, d in enumerate(all_weekdays_autumn, start=2):
        day_str = d.strftime('%Y-%m-%d')
        row_map[day_str] = idx
        sheet[f"A{idx}"] = day_str
        sheet[f"B{idx}"] = weekdays_str[d.weekday()]

    weekday_counts = [0] * num_weeks
    for j, d in enumerate(dates):
        i = assigned[j]
        weekday_counts[i] += 1
        label = f"{weekdays_str[i]}{weekday_counts[i]}"
        row = row_map[d.strftime('%Y-%m-%d')]
        sheet[f"D{row}"] = label

    # 未割当の検出と出力
    unassigned = [
        (d, weekdays_str[d.weekday()])
        for d in all_weekdays_autumn
        if sheet[f"D{row_map[d.strftime('%Y-%m-%d')]}"].value in (None, "")
    ]
    if unassigned:
        sorted_un = sorted(unassigned, key=lambda x: x[0])
        for k, (d, w) in enumerate(sorted_un[:autumn_buffer_count1]):
            sheet[f"R{k+2}"] = d.strftime('%Y-%m-%d')
        row_out = 2
        for d, w in sorted_un[autumn_buffer_count1:]:
            sheet[f"F{row_out}"] = d.strftime('%Y-%m-%d')
            sheet[f"G{row_out}"] = w
            row_out += 1
    else:
        sheet["R2"] = "未割当なし"
        sheet["F2"] = "すべて割り当て済み"

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    st.session_state["autumn1_excel_bytes"] = output

    # 14) 割当リストを返す
    return assigned
    # st.download_button(
    #     #label="📥 秋1結果付きExcelファイルをダウンロード",
    #     data=output,
    #     file_name="optimized_calendar_autumn1.xlsx",
    #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    # )
# else:
#     st.error("❌ 秋1の最適解が見つかりませんでした。")





#--------------秋2-------------------------------------

def run_autumn2_optimization(autumn_buffer_count2, spring_holidays, autumn_holidays):
    """
    秋学期Bモジュール（25日分）の曜日割当を最適化し、Excel に書き込みます。
    返り値は dates と同順の最適割当曜日リスト (0=月曜…4=金曜)。
    """
    # ── 1) 秋1結果の存在チェック ──
    if "autumn1_excel_bytes" not in st.session_state:
        st.error("秋1の最適化結果が見つかりません。")
        return []

    # ── 2) Excel から未割当候補日（F列）を読み込み ──
    data = io.BytesIO(st.session_state["autumn1_excel_bytes"].getvalue())
    wb   = load_workbook(data)
    sheet = wb.active

    all_days = []
    for row in range(2, sheet.max_row + 1):
        d = safe_to_date(sheet[f"F{row}"].value)
        if d != date.max:
            all_days.append(d)

    if len(all_days) < 25:
        st.warning("秋2に使える未割当日が25日未満です。")
        return []
    dates     = all_days[:25]
    num_days  = len(dates)     # == 25
    num_weeks = 5

    # ── 3) 実曜日インデックス & 曜日文字列 ──
    actual_weekdays = [d.weekday() for d in dates]
    weekdays_str    = ["月曜","火曜","水曜","木曜","金曜"]

    # ── 4) ISO週ID → 週ごとの日付リスト ──
    date_to_weekid = {}
    week_dict      = defaultdict(list)
    for d in dates:
        wid = d.isocalendar()[:2]
        date_to_weekid[d] = wid
        week_dict[wid].append(d)

    # ── 5) 連続振替日の検出 ──
    global dup_dates_global
    change_events = []
    for name, ds in st.session_state.event_periods.items():
        if "曜日変更" in name:
            m = re.search(r"\((月|火|水|木|金)\)", name)
            if m:
                wd = "月火水木金".index(m.group(1))
                change_events += [(d, wd) for d in ds]
    dup_dates_global = detect_dup_shift_days(change_events)

    # ── 6) reserve_limit を T2セルから取得 ──
    rl = safe_to_date(sheet["T2"].value)
    reserve_limit = rl if rl != date.max else date.max

    # ── 7) ペナルティ計算 ──
    week_penalties = build_week_penalties(
        week_dict,
        jpholiday=jpholiday,
        spring_holidays=spring_holidays,
        autumn_holidays=autumn_holidays,
        reserve_limit=reserve_limit,
        dup_dates=dup_dates_global
    )

    # ── 8) SCIP モデル構築 ──
    model = Model("Autumn2_Optimization")
    x = {
        (i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
        for i in range(num_weeks)
        for j in range(num_days)
    }

    # ── 9) 目的関数設定 ──
    model.setObjective(
        quicksum(
            (1 if i != actual_weekdays[j] else 0)
            * week_penalties[date_to_weekid[dates[j]]]
            * x[i, j]
            for i in range(num_weeks)
            for j in range(num_days)
        ),
        "minimize"
    )

    # ──10) 制約：各日１回、各週５日ずつ ──
    for j in range(num_days):
        model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
    for i in range(num_weeks):
        model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)

    # ──11) 最適化実行 ──
    model.optimize()
    if model.getStatus() != "optimal":
        st.error("❌ 秋2の最適解が見つかりませんでした。")
        return []

    # ──12) 割当抽出 ──
    assigned = [
        next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
        for j in range(num_days)
    ]

    # ──13) 距離ベース平準化 ──
    global positions_by_weekday
    positions_by_weekday = build_positions(actual_weekdays)
    runs, cur = [], []
    for idx in range(num_days):
        if assigned[idx] != actual_weekdays[idx]:
            cur.append(idx)
        else:
            if cur:
                runs.append(cur); cur = []
    if cur:
        runs.append(cur)
    for run in runs:
        normalize_run(run, dates, actual_weekdays, assigned)

    
     # === Excel への書き込み（以降は変更なし）===
    weekday_counts = [5] * num_weeks  # 秋1からの続きと仮定

    # === 曜日ラベルを I列 に書き込み
    for j, d in enumerate(dates):
        wd = assigned[j]
        weekday_counts[wd] += 1
        label = f"{weekdays_str[wd]}{weekday_counts[wd]}"
        # F列の日付と突き合わせて I列に書き込む
        for row in range(2, sheet.max_row + 1):
            if safe_to_date(sheet[f"F{row}"].value) == d:
                sheet[f"I{row}"] = label
                break

    sheet["T1"] = "Bモジュール予備日"
    sheet["K1"] = "日付"
    sheet["L1"] = "実曜日"
    sheet["N1"] = "割り当て結果"

    unassigned_dates = []
    for row in range(2, sheet.max_row + 1):
        fcell = safe_to_date(sheet[f"F{row}"].value)
        if fcell != date.max and not sheet[f"I{row}"].value:
            unassigned_dates.append((fcell, weekdays_str[fcell.weekday()]))

    if unassigned_dates:
        sorted_un = sorted(unassigned_dates, key=lambda x: x[0])
        for idx, (d, _) in enumerate(sorted_un[:autumn_buffer_count2]):
            sheet[f"T{idx+2}"] = d.strftime("%Y-%m-%d")
        row_out = 2
        for d, w in sorted_un[autumn_buffer_count2:]:
            sheet[f"K{row_out}"] = d.strftime("%Y-%m-%d")
            sheet[f"L{row_out}"] = w
            row_out += 1
    else:
        sheet["T2"] = "未割当なし"
        sheet["K2"] = "すべて割り当て済み"

    # === 書き出し
    output2 = io.BytesIO()
    wb.save(output2)
    output2.seek(0)
    st.session_state["autumn2_excel_bytes"] = output2
    return assigned
        # st.download_button(
        #     label="📥 秋2結果付きExcelファイルをダウンロード",
        #     data=output2,
        #     file_name="optimized_calendar_autumn2.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # )
    # else:
    #     st.error("❌ 秋2の最適解が見つかりませんでした。")






#-----------------------秋3----------------------------

def run_autumn3_optimization(autumn_buffer_count3, spring_holidays, autumn_holidays, autumn3_start_date):
    """
    秋学期Cモジュール（25日分）の曜日割当を最適化し、Excel に書き込みます。
    Args:
        autumn_buffer_count3: int   Cモジュールの予備日数（Excel出力に影響）
        spring_holidays:      set[date] 春学期の休講日（本関数内では未使用）
        autumn_holidays:      set[date] 秋学期の休講日
        autumn3_start_date:   date   Cモジュール開始日
    Returns:
        assigned: List[int] 最適割当曜日リスト (0=月曜…4=金曜)
    """
    # st.markdown("---")
    # 1) 秋2結果がなければエラー
    if "autumn2_excel_bytes" not in st.session_state:
        st.error("秋2までの最適化結果が見つかりません。先に秋1・秋2を完了してください。")
        return []

    # 2) Excel から未割当候補日（K列）を読み込み
    data = io.BytesIO(st.session_state["autumn2_excel_bytes"].getvalue())
    wb   = load_workbook(data)
    sheet = wb.active

    candidates = []
    for row in range(2, sheet.max_row + 1):
        d = safe_to_date(sheet[f"K{row}"].value)
        if d != date.max and d >= autumn3_start_date:
            candidates.append(d)

    if len(candidates) < 25:
        st.warning("秋3に使える未割当日が25日未満です。")
        st.stop()

    # 3) 最初の25日を対象に
    dates     = sorted(candidates)[:25]
    num_days  = len(dates)     # == 25
    num_weeks = 5

    # 4) 実曜日インデックス取得 & 曜日文字列
    actual_weekdays = [d.weekday() for d in dates]
    weekdays_str    = ["月曜","火曜","水曜","木曜","金曜"]

    # 5) ISO週ID → 週ごとの日付リスト
    date_to_weekid = {}
    week_dict      = defaultdict(list)
    for d in dates:
        wid = d.isocalendar()[:2]
        date_to_weekid[d] = wid
        week_dict[wid].append(d)

    # 6) 連続振替日の検出
    global dup_dates_global
    change_events = []
    for name, ds in st.session_state.event_periods.items():
        if "曜日変更" in name:
            m = re.search(r"\((月|火|水|木|金)\)", name)
            if m:
                wd = "月火水木金".index(m.group(1))
                change_events += [(d, wd) for d in ds]
    dup_dates_global = detect_dup_shift_days(change_events)

    # 7) reserve_limit を V2セルから取得
    rl = safe_to_date(sheet["V2"].value)
    reserve_limit = rl if rl != date.max else date.max

    # 8) 週ごとのペナルティ計算
    week_penalties = build_week_penalties(
        week_dict,
        jpholiday=jpholiday,
        spring_holidays=spring_holidays,
        autumn_holidays=autumn_holidays,
        reserve_limit=reserve_limit,
        dup_dates=dup_dates_global
    )

    # 9) SCIP モデル構築
    model = Model("Autumn3_Optimization")
    x = {
        (i, j): model.addVar(vtype="B", name=f"x_{i}_{j}")
        for i in range(num_weeks)
        for j in range(num_days)
    }

    # 10) 目的関数設定
    model.setObjective(
        quicksum(
            (1 if i != actual_weekdays[j] else 0)
            * week_penalties[date_to_weekid[dates[j]]]
            * x[i, j]
            for i in range(num_weeks)
            for j in range(num_days)
        ),
        "minimize"
    )

    # 11) 制約：各日１回、各週５日ずつ
    for j in range(num_days):
        model.addCons(quicksum(x[i, j] for i in range(num_weeks)) == 1)
    for i in range(num_weeks):
        model.addCons(quicksum(x[i, j] for j in range(num_days)) == 5)

    # 12) 最適化実行
    model.optimize()
    if model.getStatus() != "optimal":
        st.error("❌ 秋3の最適解が見つかりませんでした。")
        return []

    # 13) 割当抽出
    assigned = [
        next(i for i in range(num_weeks) if model.getVal(x[i, j]) > 0.5)
        for j in range(num_days)
    ]

    # 14) 距離ベース平準化
    global positions_by_weekday
    positions_by_weekday = build_positions(actual_weekdays)
    runs, cur = [], []
    for idx in range(num_days):
        if assigned[idx] != actual_weekdays[idx]:
            cur.append(idx)
        else:
            if cur:
                runs.append(cur)
                cur = []
    if cur:
        runs.append(cur)
    for run in runs:
        normalize_run(run, dates, actual_weekdays, assigned)

    # 15) Excel への書き込み（以下は変更しない）
    weekday_counts = [10] * num_weeks  # 秋1＋秋2＝10件ずつ仮定

    for j, d in enumerate(dates):
        wd = assigned[j]
        weekday_counts[wd] += 1
        label = f"{weekdays_str[wd]}{weekday_counts[wd]}"
        # K列の日付に対応する行を探し、N列に書き込む
        for row in range(2, sheet.max_row + 1):
            if safe_to_date(sheet[f"K{row}"].value) == d:
                sheet[f"N{row}"] = label
                break


    # === V列に最も早い未割当日を書き出し
    sheet["V1"] = "Cモジュール予備日"
    unassigned = []
    for row in range(2, sheet.max_row + 1):
        d = safe_to_date(sheet[f"K{row}"].value)
        if d != date.max and not sheet[f"N{row}"].value and d >= autumn3_start_date:
            unassigned.append(d)

    if unassigned:
        for idx, d in enumerate(sorted(unassigned)[:autumn_buffer_count3]):
            sheet[f"V{idx+2}"] = d.strftime("%Y-%m-%d")
    else:
        sheet["V2"] = "未割当なし"

    # 16) 保存してセッションに格納
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    st.session_state["autumn3_excel_bytes"] = output
    st.session_state["autumn3_done"] = True

    return assigned
        # st.download_button(
        #     label="  秋学期のExcelファイルをダウンロード",
        #     data=output4,
        #     file_name="optimized_calendar_autumn3.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # )
    # else:
    #     st.error("❌ 秋3の最適解が見つかりませんでした。")






if st.button("📘 学年暦を作成"):
    run_spring1_optimization(spring_holidays, autumn_holidays, spring_buffer_count1)
    run_spring2_optimization(spring_buffer_count2, spring_holidays, autumn_holidays)
    run_spring3_optimization(spring_buffer_count3, spring_holidays, autumn_holidays)
    run_autumn1_optimization(spring_holidays, autumn_holidays, autumn_buffer_count1)
    run_autumn2_optimization(autumn_buffer_count2, spring_holidays, autumn_holidays)
    run_autumn3_optimization(autumn_buffer_count3, spring_holidays, autumn_holidays, autumn3_start)





# ———— 完了メッセージ ————
# if st.session_state.get("spring3_done"):
#     st.success("✨ 春学期の割り当てが完了しました！")

# if st.session_state.get("autumn3_done"):
#     st.success("✨ 秋学期の割り当てが完了しました！")
# ————————————————
# —————— ダウンロードボタン ——————
create_container = st.container()
with create_container:
    st.markdown("---")
    st.header("📥 詳細な割り当て結果のダウンロード")
if "spring3_excel_bytes" in st.session_state:
    st.download_button("🌸 春学期の詳細な割り当てExcelをダウンロード",
                       st.session_state["spring3_excel_bytes"].getvalue(),
                       "spring3.xlsx", key="dl_s3")

if "autumn3_excel_bytes" in st.session_state:
    st.download_button("🍂 秋学期の詳細な割り当てExcelをダウンロード",
                       st.session_state["autumn3_excel_bytes"].getvalue(),
                       "autumn3.xlsx", key="dl_a3")




#-----------------------カレンダー表示生成----------------------------

# 🔻make_calendar_* の冒頭で spring_ws / autumn_ws を開いた直後に置く
def collect_busy_dates(*worksheets) -> set[date]:
    """
    春3・秋3 のシートから『実際に授業が割り当てられている日』を
    set[date] で返す。D/I/N 列に何か書かれている行を採用。
    """
    busy = set()
    for ws in worksheets:
        for dcol, acol in (("A", "D"), ("F", "I"), ("K", "N")):
            for r in range(2, ws.max_row + 1):
                if ws[f"{acol}{r}"].value:                # ← 割当が入っている行
                    raw = ws[f"{dcol}{r}"].value          # ← 対応する日付
                    if not raw:
                        continue
                    d = raw.date() if isinstance(raw, datetime) \
                        else datetime.strptime(str(raw), "%Y-%m-%d").date()
                    busy.add(d)
    return busy




#-------------------------------------------------------4か月1枚----------------------------------------------------


import io
import re
import calendar
from datetime import date, timedelta, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.pagebreak import Break
from dateutil.relativedelta import relativedelta
import jpholiday
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# ★ 追加: セル座標 (row,col) → datetime.date のマッピングを構築するユーティリティ
# ──────────────────────────────────────────────────────────────────────────────

def build_date_map(ws, year_start):
    """
    ワークシート ``ws`` について、
      ・マージされた「〇月」ヘッダー (A,Q 列) を検出し
      ・その下の 1〜31 の数字セルを走査して
    (row,col) → datetime.date の辞書を返す。

    例: ``date_map[(5, 4)] == datetime.date(2026, 4, 3)``
    """
    # 1) 行番号 → 月 を作る
    row2month = {}
    for mrange in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = mrange.bounds
        col_letter = get_column_letter(min_col)
        if col_letter not in ("A", "Q"):
            continue
        label = str(ws.cell(min_row, min_col).value or "")
        m = re.match(r"(\d{1,2})月", label)
        if not m:
            continue
        month = int(m.group(1))
        for r in range(min_row, max_row + 1):
            row2month[r] = month

    # 2) 個々のセルを調べ (row,col)→date を埋める
    date_map = {}
    for r, mon in row2month.items():
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, int):
                continue
            dnum = v
            if 1 <= dnum <= 31:
                yr = year_start if mon >= 4 else year_start + 1
                try:
                    date_map[(r, c)] = date(yr, mon, dnum)
                except ValueError:
                    # 2月30日など不正日は無視
                    pass
    return date_map

# ──────────────────────────────────────────────────────────────────────────────
# 以降の関数は既存ロジックを極力変更せずに維持しつつ、
# build_date_map を活用できるよう最小限の改修を加える。
# ──────────────────────────────────────────────────────────────────────────────

def collect_event_labels_remark(
    year_start: int,
    spring3_bytes: io.BytesIO,
    autumn3_bytes: io.BytesIO,
    autumn2_bytes: io.BytesIO,
    start_date_summer: date,
    start_date_autumn: date,
    event_periods: dict,
    # manual_remark: dict,
    
) -> dict:
    
    # event_labels_remark = {}
    event_labels_remark: dict[date, list[str]] = {}

    def add_remark(d: date, txt: str):
        lst = event_labels_remark.setdefault(d, [])
        if txt not in lst:
            lst.append(txt)

    # 1) 春学期・秋学期開始日
    add_remark(start_date_summer,
               f"{start_date_summer.month}月{start_date_summer.day}日　春学期授業開始")
    add_remark(start_date_autumn,
               f"{start_date_autumn.month}月{start_date_autumn.day}日　秋学期授業開始")

    # 2) 手動で登録された備考をマージ
    # for d, txts in (manual_remark or {}).items():
    #     for txt in txts:
    #         add_remark(d, txt)
    grouped = defaultdict(set)   # base_name → set(すべての dates)
    for name_with_suffix, dates in (event_periods or {}).items():
        base = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)
        grouped[base].update(dates)

    for base_name, dates in grouped.items():
        ds = sorted(dates)
        # run 検出
        runs: list[tuple[date,date]] = []
        start = prev = ds[0]
        for d in ds[1:]:
            if d == prev + timedelta(days=1):
                prev = d
            else:
                runs.append((start, prev))
                start = prev = d
        runs.append((start, prev))

        # 各 run ごとに「○月△日～○月◇日　base_name」を
        # 区間内の各月最初の日付キーで一度だけ登録
        for s, e in runs:            
            if s == e:
                remark = f"{s.month}月{s.day}日　{base_name}"
            else:
                remark = f"{s.month}月{s.day}日～{e.month}月{e.day}日　{base_name}"
            months = sorted({d.month for d in ds if s <= d <= e})
            for m in months:
                first_of_month = min(d for d in ds if s <= d <= e and d.month == m)
                add_remark(first_of_month, remark)

    # 3) イベント期間を連続区間検出してマージ
    for name_with_suffix, dates in (event_periods or {}).items():
        if not dates:
            continue
        base_name = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)
        ds = sorted(dates)
        runs = []
        run_start = run_end = ds[0]
        for d0 in ds[1:]:
            if d0 == run_end + timedelta(days=1):
                run_end = d0
            else:
                runs.append((run_start, run_end))
                run_start = run_end = d0
        runs.append((run_start, run_end))
        for st_day, ed_day in runs:
            if st_day == ed_day:
                remark = f"{st_day.month}月{st_day.day}日　{base_name}"
            else:
                remark = f"{st_day.month}月{st_day.day}日～{ed_day.month}月{ed_day.day}日　{base_name}"
            add_remark(st_day, remark)

    # 4) モジュール予備日を読み込んでマージ
    spring_ws = load_workbook(io.BytesIO(spring3_bytes.getvalue())).active
    autumn_ws = load_workbook(io.BytesIO(autumn3_bytes.getvalue())).active
    def add_module_days(ws_src, module_label):
        col_map = {"A": "R", "B": "T", "C": "V"}
        col_letter = col_map[module_label]
        for r in range(2, ws_src.max_row + 1):
            v = ws_src[f"{col_letter}{r}"].value
            if not v:
                continue
            d0 = v.date() if isinstance(v, datetime) else datetime.strptime(str(v), "%Y-%m-%d").date()
            remark = f"{d0.month}月{d0.day}日　{module_label}モジュール予備日"
            add_remark(d0, remark)

    for sheet, module in ((spring_ws, "A"), (spring_ws, "B"), (spring_ws, "C"),
                          (autumn_ws, "A"), (autumn_ws, "B"), (autumn_ws, "C")):
        add_module_days(sheet, module)

    # 5) 振替授業を読み込んでマージ
    def extract_compensations(ws_src):
        out = []
        for dcol, ocol, acol in (("A","B","D"), ("F","G","I"), ("K","L","N")):
            for r in range(2, ws_src.max_row + 1):
                dv = ws_src[f"{dcol}{r}"].value
                ov = ws_src[f"{ocol}{r}"].value
                av = ws_src[f"{acol}{r}"].value
                if not (dv and ov and av):
                    continue
                m = re.match(r"[月火水木金]", str(av))
                if not m or ov == m.group(0) + "曜":
                    continue
                dt0 = dv.date() if isinstance(dv, datetime) else datetime.strptime(str(dv), "%Y-%m-%d").date()
                out.append((dt0, m.group(0) + "曜"))
        return out

    for sh in (spring_ws, autumn_ws):
        for dt0, wd in extract_compensations(sh):
            remark = f"{dt0.month}月{dt0.day}日　{wd}授業に変更"
            add_remark(dt0, remark)

    # 6) 休業期間をマージ
    def month_starts_between(st_day, ed_day):
        current = date(st_day.year, st_day.month, 1)
        while current <= ed_day:
            yield current
            current += relativedelta(months=1)

    def add_period_remarks(st_day: date, ed_day: date, label: str):
        for ms in month_starts_between(st_day, ed_day):
            display_start = max(st_day, ms)
            next_month = ms + relativedelta(months=1)
            last_of_month = next_month - timedelta(days=1)
            display_end = min(ed_day, last_of_month)
            add_remark(display_start,
                       f"{display_start.month}月{display_start.day}日～{display_end.month}月{display_end.day}日　{label}")

    # A) 春季休業
    enroll = (event_periods or {}).get("入学式")
    if enroll:
        enroll_day = min(enroll)
        st_day, ed_day = date(year_start, 4, 1), enroll_day - timedelta(days=1)
        if st_day <= ed_day:
            add_period_remarks(st_day, ed_day, "春季休業")

    # B) 夏季休業
    if spring3_bytes:
        ws3 = load_workbook(io.BytesIO(spring3_bytes.getvalue())).active
        dates3 = [
            (c.date() if isinstance(c, datetime) else datetime.strptime(str(c), "%Y-%m-%d").date())
            for r in range(2, ws3.max_row + 1)
            if (c := ws3[f"V{r}"].value)
        ]
        if dates3:
            last3 = max(dates3)
            st3, ed3 = last3 + timedelta(days=1), start_date_autumn - timedelta(days=1)
            if st3 <= ed3:
                add_period_remarks(st3, ed3, "夏季休業")

    # C) 冬期休業 (秋2予備日の翌日～秋C開始日前日)
    if autumn2_bytes:
        ws2 = load_workbook(io.BytesIO(autumn2_bytes.getvalue())).active
        dates2 = [
            (c.date() if isinstance(c, datetime) else
            datetime.strptime(str(c), "%Y-%m-%d").date())
            for r in range(2, ws2.max_row + 1)
            if (c := ws2[f"T{r}"].value)                  # 秋Bモジュール予備日 (T列)
        ]
        if dates2:
            last2 = max(dates2)

            # ① フォームで選択された秋C開始日を最優先
            autumn3_start = st.session_state.get("autumn3_start")

            # ② フォームが空(None)なら、従来どおり V 列の最初の日付で代替
            if autumn3_start is None:
                autumn3_start = min(
                    [
                        (c.date() if isinstance(c, datetime) else
                        datetime.strptime(str(c), "%Y-%m-%d").date())
                        for r in range(2, autumn_ws.max_row + 1)
                        if (c := autumn_ws[f"V{r}"].value)          # 秋C予備日 (V列)
                    ],
                    default=None
                )

            # ③ 冬休み区間を作成
            if autumn3_start:
                st2 = last2 + timedelta(days=1)           # 休業開始 = 秋B予備日の翌日
                ed2 = autumn3_start - timedelta(days=1)   # 休業終了 = 秋C開始日の前日
                if st2 <= ed2:
                    add_period_remarks(st2, ed2, "冬期休業")


    # D) 年度末春季休業
    if autumn3_bytes:
        wsa3 = load_workbook(io.BytesIO(autumn3_bytes.getvalue())).active
        dates_a3 = [
            (c.date() if isinstance(c, datetime) else datetime.strptime(str(c), "%Y-%m-%d").date())
            for r in range(2, wsa3.max_row + 1)
            if (c := wsa3[f"V{r}"].value)
        ]
        if dates_a3:
            last_a3 = max(dates_a3)
            st4, ed4 = last_a3 + timedelta(days=1), date(year_start + 1, 3, 31)
            if st4 <= ed4:
                add_period_remarks(st4, ed4, "春季休業")

    return event_labels_remark

def make_calendar_4months(year_start: int) -> bytes:
    # ───────────── 既存 import（関数内で再 import されていた分）───────────
    import calendar
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.worksheet.pagebreak import Break

    # ─── セッションステートから必要な情報を取得 ───
    spring3_bytes     = st.session_state["spring3_excel_bytes"]
    autumn3_bytes     = st.session_state["autumn3_excel_bytes"]
    autumn2_bytes     = st.session_state.get("autumn2_excel_bytes", None)
    start_date_summer = st.session_state["start_summer"]
    start_date_autumn = st.session_state["start_autumn"]
    event_periods     = st.session_state.get("event_periods", {})
    manual_remark     = st.session_state.get("event_labels_remark", {})
    # 祝日名辞書（session_state から取得）
    holiday_names    = st.session_state.get("holiday_names", {})
    manual_holidays   = st.session_state.get("manual_holidays_all", set())
    # 祝日名辞書（session_state から取得）
    holiday_names    = st.session_state.get("holiday_names", {})

    # ─── 春３・秋３ ワークシートを事前読み込み ───
    spring_ws = load_workbook(io.BytesIO(spring3_bytes.getvalue())).active
    autumn_ws = load_workbook(io.BytesIO(autumn3_bytes.getvalue())).active
    #busy_dates = collect_busy_dates(spring_ws, autumn_ws)

    # ➀ 備考データをまとめる（既存関数をそのまま呼ぶ）
    event_labels_remark = collect_event_labels_remark(
        year_start,
        spring3_bytes,
        autumn3_bytes,
        autumn2_bytes,
        start_date_summer,
        start_date_autumn,
        event_periods,
        manual_remark
        
    )

    # ➁ Workbook 初期化 & シート作成（以降、既存ロジックを保持）
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet()

    # タイトル行
    ws.insert_rows(1)
    ws.merge_cells("A1:O1")
    t = ws["A1"]
    t.value = f"{year_start}年度 筑波大学 学年暦"
    t.font = Font(size=16, bold=True)
    t.alignment = Alignment("center", "center")
    ws.row_dimensions[1].height = 30

    ws["P1"] = "学群/大学院(筑波キャンパス)"
    sub = ws["P1"]
    sub.font = Font(size=10)
    sub.alignment = Alignment("right", "center")

    # 印刷設定
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.print_title_rows       = '1:2'

    # 曜日ヘッダー
    days = ['日','月','火','水','木','金','土']
    for i, d in enumerate(days):
        ccol = 2 + 2*i
        ws.merge_cells(start_row=2, start_column=ccol, end_row=2, end_column=ccol+1)
        cell = ws.cell(row=2, column=ccol)
        cell.value = d
        cell.font = Font(size=11, bold=True)
        cell.alignment = Alignment("center","center")

    # 備考ヘッダー
    ws.merge_cells("P2")
    r2 = ws["P2"]
    r2.value = "備考"
    r2.font = Font(size=11, bold=True)
    r2.alignment = Alignment("center","center")

    # 日付描画＋月名＋備考欄
    row_ptr     = 3
    red_font    = Font(color="9C0006", size=11)
    std_font    = Font(size=11)
    sm10        = Font(size=10)
    max_end_row = 0

    for m_off in range(12):
        m = ((4-1 + m_off) % 12) + 1
        y = year_start if m >= 4 else year_start + 1

        first   = date(y, m, 1)
        wd0     = first.isoweekday() % 7
        length  = calendar.monthrange(y, m)[1]
        dlist   = [ first + timedelta(days=i) for i in range(length) ]

        start_row   = row_ptr
        current_row = row_ptr

        # ① 空白
        for i in range(wd0):
            ws.cell(row=current_row, column=2+2*i, value="")
            ws.cell(row=current_row, column=3+2*i, value="")

        # ② 日付
        for d in dlist:
            ccol = d.isoweekday() % 7
            cell = ws.cell(row=current_row, column=2+2*ccol, value=d.day)
            cell.alignment = Alignment("center","center")
            cell.font = red_font if (jpholiday.is_holiday(d) or d.weekday()==6) else std_font
            if ccol == 6:
                current_row += 1

        end_row = current_row

        # ③ 末尾空行削除
        while end_row>=start_row and all(
            ws.cell(row=end_row, column=c).value in (None,"")
            for c in range(1,17)
        ):
            ws.delete_rows(end_row)
            end_row -= 1

        # ④ 月名縦マージ
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        mon_c = ws.cell(row=start_row, column=1)
        mon_c.value = f"{m}月"
        mon_c.font = std_font
        mon_c.alignment = Alignment("center","center")

        # ⑤ 備考欄縦マージ＋書き込み
        ws.merge_cells(start_row=start_row, start_column=16, end_row=end_row, end_column=16)
        rc   = ws.cell(row=start_row, column=16)
        rc.font = sm10
        rc.alignment = Alignment("left","top", wrap_text=True)
        items = [
            (dobj, txt)
            for dobj, lst in event_labels_remark.items()
            if dobj.year == y and dobj.month == m
            for txt in lst
        ]
        items.sort(key=lambda x: x[0])
        if items:
            rc.value = "\n".join(txt for (_,txt) in items)

        # ⑥ 改ページ
        if m_off in (3,7):
            ws.row_breaks.append(Break(id=end_row))

        row_ptr = end_row + 1 if not (m==3 and y==year_start+1) else end_row
        max_end_row = max(max_end_row, end_row)

    # シート末尾の空行を削除
    while True:
        last = ws.max_row
        if all(ws.cell(row=last, column=c).value in (None,"") for c in range(1,17)):
            for mr in list(ws.merged_cells.ranges):
                if mr.min_row <= last <= mr.max_row:
                    ws.unmerge_cells(str(mr))
            ws.delete_rows(last)
        else:
            break

    # 罫線
    thin = Side(border_style="thin", color="FF000000")
    bd   = Border(top=thin,bottom=thin,left=thin,right=thin)
    for r in range(2, max_end_row+1):
        for c in range(1,17):
            ws.cell(row=r, column=c).border   = bd

    # ── ⑥ モジュール色付け、振替授業、予備日、休講日、祝日名、罫線 ──

    cell2date = build_date_map(ws, year_start)
    module_labels = {}
    for sh in (spring_ws, autumn_ws):
        for dcol, acol in (("A", "D"), ("F", "I"), ("K", "N")):
            for r in range(2, sh.max_row + 1):
                dv = sh[f"{dcol}{r}"].value
                av = sh[f"{acol}{r}"].value
                if not (dv and av):
                    continue
                # 日付を date 型へ
                if isinstance(dv, datetime):
                    dt0 = dv.date()
                else:
                    try:
                        dt0 = datetime.strptime(str(dv), "%Y-%m-%d").date()
                    except ValueError:
                        continue
                # '月曜3' → '月3' に整形（好みで '月曜3' のままでも可）
                lbl = str(av).replace("曜", "")
                module_labels[(dt0.year, dt0.month, dt0.day)] = lbl

    # モジュール色と予備日・振替授業を収集
    y_fill = PatternFill("solid", fgColor="FFFFE599")
    b_fill = PatternFill("solid", fgColor="FFB8E1FC")
    g_fill = PatternFill("solid", fgColor="FFC6E0B4")

    def extract_module(ws_src, cond_col, date_col, fill):
        out = []
        for r in range(2, ws_src.max_row + 1):
            if ws_src[f"{cond_col}{r}"].value:
                dv = ws_src[f"{date_col}{r}"].value
                d0 = dv.date() if isinstance(dv, datetime) else datetime.strptime(str(dv), "%Y-%m-%d").date()
                out.append((d0, fill))
        return out

    module_marks = []
    module_marks += extract_module(spring_ws, "D", "A", y_fill)
    module_marks += extract_module(spring_ws, "I", "F", b_fill)
    module_marks += extract_module(spring_ws, "N", "K", g_fill)
    module_marks += extract_module(autumn_ws, "D", "A", y_fill)
    module_marks += extract_module(autumn_ws, "I", "F", b_fill)
    module_marks += extract_module(autumn_ws, "N", "K", g_fill)

    reserve_dates = []
    for ws_src, col in ((spring_ws,"R"),(spring_ws,"T"),(spring_ws,"V"),
                       (autumn_ws,"R"),(autumn_ws,"T"),(autumn_ws,"V")):
        for r in range(2, ws_src.max_row + 1):
            raw = ws_src[f"{col}{r}"].value
            if raw:
                d0 = raw.date() if isinstance(raw, datetime) else datetime.strptime(str(raw), "%Y-%m-%d").date()
                reserve_dates.append(d0)

    def extract_comp(ws_src):
        out = []
        for dcol, ocol, acol in (("A","B","D"), ("F","G","I"), ("K","L","N")):
            for r in range(2, ws_src.max_row+1):
                dv = ws_src[f"{dcol}{r}"].value   # 日付
                ov = ws_src[f"{ocol}{r}"].value   # “月曜” など実曜日
                av = ws_src[f"{acol}{r}"].value   # “火2” “月1” … 割当ラベル
                if not (dv and ov and av):
                    continue

                # 先頭の曜日文字だけを取り出す
                m_ov = re.match(r"[月火水木金]", str(ov))
                m_av = re.match(r"[月火水木金]", str(av))
                if not (m_ov and m_av):
                    continue

                # 同じ曜日なら振替ではない
                if m_ov.group(0) == m_av.group(0):
                    continue

                # --- ここまで来れば確実に振替授業 ---
                try:
                    dt = dv.date() if isinstance(dv, datetime) else datetime.strptime(str(dv), "%Y-%m-%d").date()
                except Exception:
                    continue
                out.append((dt, str(av).replace("曜", "")))   # “火曜13” → “火13” に
        return out

    comp_marks = extract_comp(spring_ws) + extract_comp(autumn_ws)
    

    mod_dict  = {(d.year,d.month,d.day):f for d,f in module_marks}
    res_set   = {(d.year,d.month,d.day) for d in reserve_dates}
    comp_dict = {(d.year,d.month,d.day):lbl for d,lbl in comp_marks}
    manual_holidays_set = set(manual_holidays)

    thin = Side(border_style="thin", color="FF000000")
    bd   = Border(top=thin,bottom=thin,left=thin,right=thin)
    reserve_fill = PatternFill("solid", fgColor="FFBFBFBF")
    purple       = PatternFill("solid", fgColor="FFD3B8F5")
    red_font2    = Font(color="9C0006", size=10)
    # 曜日カウント用
    lesson_count: dict[int,int] = {}

    for (r, c), dt in cell2date.items():
        key = (dt.year, dt.month, dt.day)
        tgt = ws.cell(r, c+1)

        # モジュール日 → 「曜日＋回数」を書き込み
        if key in comp_dict:                       # ① 振替授業を最優先
            tgt.fill = purple
            tgt.value = comp_dict[key]
            tgt.alignment = Alignment("center", "center")

        elif key in mod_dict:                      # ② モジュール日
            tgt.fill = mod_dict[key]
            if key in module_labels:
                tgt.value = module_labels[key]
            else:
                wd = dt.weekday()
                lesson_count.setdefault(wd, 0)
                lesson_count[wd] += 1
                wd_names = ['月','火','水','木','金','土','日']
                tgt.value = f"{wd_names[wd]}曜{lesson_count[wd]}"
            tgt.alignment = Alignment("center", "center")


        # # 振替授業
        # elif key in comp_dict:
        #     tgt.fill = purple
        #     tgt.value = comp_dict[key]
        #     tgt.alignment = Alignment("center","center")

        #予備日
        elif key in res_set:                       
            tgt.fill = reserve_fill
            tgt.value = "予備日"
            tgt.alignment = Alignment("center", "center")

        # 手動休講日
        elif dt in manual_holidays_set and dt.weekday() < 5 and not jpholiday.is_holiday(dt):
            tgt.value = "休講日"
            tgt.font = red_font2 
            tgt.alignment = Alignment("center","center")

        # 祝日名
        elif dt in holiday_names:
            tgt.value = holiday_names[dt]
            tgt.font = red_font2
            tgt.alignment = Alignment("center","center", wrap_text=True)

        # 罫線
        ws.cell(r, c).border   = bd
        ws.cell(r, c+1).border = bd

    try:
        date_map = build_date_map(ws, year_start)
        st.session_state["calendar_date_map"] = date_map
    except Exception as e:
        # マッピング生成でエラーになってもメイン処理は壊さない
        st.warning(f"build_date_map 実行で例外が発生しました: {e}")
    

    # ── フォントサイズ／列幅・行高・罫線 ──
    def px_to_col(px): return (px - 5) / 7
    ws.column_dimensions['A'].width = px_to_col(49)
    for c in ('B','D','F','H','J','L','N'):
        ws.column_dimensions[c].width = px_to_col(33)
    for c in ('C','E','G','I','K','M','O'):
        ws.column_dimensions[c].width = px_to_col(66)
    ws.column_dimensions['P'].width = px_to_col(270)

    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = (30 * 0.75)

    # 列幅・行高調整、省略行などの後… 最終的に Workbook を返す
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()












#-------------------------------------------------------6か月1枚----------------------------------------------------



import io
import re
import calendar
from datetime import date, timedelta, datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.pagebreak import Break
from dateutil.relativedelta import relativedelta
import jpholiday
import streamlit as st


cell2date: dict[tuple[int, int], date] | None = (
        st.session_state.get("calendar_date_map_master")
)

def collect_event_labels_remark(year_start: int, 
                                spring3_bytes: io.BytesIO, 
                                autumn3_bytes: io.BytesIO, 
                                autumn2_bytes: io.BytesIO,
                                start_date_summer: date, 
                                start_date_autumn: date,
                                event_periods: dict, 
                                manual_remark: dict,
                                ) -> dict:
    """
    『春3・秋3 の Excel』バイト列（BytesIO）や手動イベント情報を受け取り、
    ①春学期／秋学期開始日
    ②手動で登録された event_labels_remark（session_state["event_labels_remark"] 相当）
    ③イベント期間（event_periods）
    ④モジュール予備日（春3・秋3）
    ⑤振替授業（春3・秋3）
    ⑥休業期間（春季・夏季・冬季・年度末春季）
    をすべて month-start-once-per-month の形でマージし、辞書形式で返す。
    戻り値: { date_obj: [remark1, remark2, …], … }
    """
    event_labels_remark = {}
    def add_remark(d, txt):
        lst = event_labels_remark.setdefault(d, [])
        if txt not in lst:
            lst.append(txt)

    # 1) 春学期・秋学期開始日
    add_remark(start_date_summer, f"{start_date_summer.month}月{start_date_summer.day}日　春学期授業開始")
    add_remark(start_date_autumn, f"{start_date_autumn.month}月{start_date_autumn.day}日　秋学期授業開始")

    # # 2) 手動で登録された備考をマージ
    # for d, txts in (manual_remark or {}).items():
    #     for txt in txts:
    #         add_remark(d, txt)
    grouped = defaultdict(set)   # base_name → set(すべての dates)
    for name_with_suffix, dates in (event_periods or {}).items():
        base = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)
        grouped[base].update(dates)

    for base_name, dates in grouped.items():
        ds = sorted(dates)
        # run 検出
        runs: list[tuple[date,date]] = []
        start = prev = ds[0]
        for d in ds[1:]:
            if d == prev + timedelta(days=1):
                prev = d
            else:
                runs.append((start, prev))
                start = prev = d
        runs.append((start, prev))
        for s, e in runs:            
            if s == e:
                remark = f"{s.month}月{s.day}日　{base_name}"
            else:
                remark = f"{s.month}月{s.day}日～{e.month}月{e.day}日　{base_name}"
            months = sorted({d.month for d in ds if s <= d <= e})
            for m in months:
                first_of_month = min(d for d in ds if s <= d <= e and d.month == m)
                add_remark(first_of_month, remark)

    # 3) イベント期間を連続区間検出してマージ
    for name_with_suffix, dates in (event_periods or {}).items():
        if not dates:
            continue
        base_name = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)
        ds = sorted(dates)
        runs = []
        run_start = run_end = ds[0]
        for d0 in ds[1:]:
            if d0 == run_end + timedelta(days=1):
                run_end = d0
            else:
                runs.append((run_start, run_end))
                run_start = run_end = d0
        runs.append((run_start, run_end))
        for st_day, ed_day in runs:
            if st_day == ed_day:
                remark = f"{st_day.month}月{st_day.day}日　{base_name}"
            else:
                remark = f"{st_day.month}月{st_day.day}日～{ed_day.month}月{ed_day.day}日　{base_name}"
            add_remark(st_day, remark)


    # # 3) event_periods（イベント名→日付リスト）から連続区間を検出してマージ
    # for name, dates in (event_periods or {}).items():
    #     if not dates:
    #         continue
    #     ds = sorted(dates)
    #     runs = []
    #     run_start = run_end = ds[0]
    #     for d in ds[1:]:
    #         if d == run_end + timedelta(days=1):
    #             run_end = d
    #         else:
    #             runs.append((run_start, run_end))
    #             run_start = run_end = d
    #     runs.append((run_start, run_end))

    #     for st_day, ed_day in runs:
    #         if st_day == ed_day:
    #             remark = f"{st_day.month}月{st_day.day}日　{name}"
    #         else:
    #             remark = f"{st_day.month}月{st_day.day}日～{ed_day.month}月{ed_day.day}日　{name}"
    #         add_remark(st_day, remark)
    

    # 4) モジュール予備日をマージ（春3・秋3）
    def add_module_days(ws, module_label):
        col_map = {"A": "R", "B": "T", "C": "V"}
        col_letter = col_map[module_label]
        for r in range(2, ws.max_row + 1):
            v = ws[f"{col_letter}{r}"].value
            if not v:
                continue
            d = v.date() if isinstance(v, datetime) else datetime.strptime(str(v), "%Y-%m-%d").date()
            remark = f"{d.month}月{d.day}日　{module_label}モジュール予備日"
            add_remark(d, remark)

    spring_ws = load_workbook(io.BytesIO(spring3_bytes.getvalue())).active
    autumn_ws = load_workbook(io.BytesIO(autumn3_bytes.getvalue())).active
    for sheet, module in ((spring_ws, "A"), (spring_ws, "B"), (spring_ws, "C"),
                          (autumn_ws, "A"), (autumn_ws, "B"), (autumn_ws, "C")):
        add_module_days(sheet, module)

    # 5) 振替授業をマージ
    def extract_compensations(ws):
        out = []
        for dcol, ocol, acol in (("A","B","D"), ("F","G","I"), ("K","L","N")):
            for r in range(2, ws.max_row + 1):
                dv = ws[f"{dcol}{r}"].value
                ov = ws[f"{ocol}{r}"].value
                av = ws[f"{acol}{r}"].value
                if not (dv and ov and av):
                    continue
                m = re.match(r"[月火水木金]", str(av))
                if not m or ov == m.group(0) + "曜":
                    continue
                dt = dv.date() if isinstance(dv, datetime) else datetime.strptime(str(dv), "%Y-%m-%d").date()
                out.append((dt, m.group(0) + "曜"))
        return out

    for sh in (spring_ws, autumn_ws):
        for dt, wd in extract_compensations(sh):
            remark = f"{dt.month}月{dt.day}日　{wd}授業に変更"
            add_remark(dt, remark)

    # 6) 休業期間をマージ
    def month_starts_between(st_day: date, ed_day: date):
        current = date(st_day.year, st_day.month, 1)
        while current <= ed_day:
            yield current
            current += relativedelta(months=1)

    def add_period_remarks(st_day: date, ed_day: date, label: str):
        for ms in month_starts_between(st_day, ed_day):
            display_start = max(st_day, ms)
            next_month = ms + relativedelta(months=1)
            last_of_month = next_month - timedelta(days=1)
            display_end = min(ed_day, last_of_month)
            add_remark(display_start,
                       f"{display_start.month}月{display_start.day}日～{display_end.month}月{display_end.day}日　{label}")

    # A) 春季休業 (4/1～入学式前日)
    enroll = (event_periods or {}).get("入学式")
    if enroll:
        enroll_day = min(enroll)
        st_day, ed_day = date(year_start, 4, 1), enroll_day - timedelta(days=1)
        if st_day <= ed_day:
            add_period_remarks(st_day, ed_day, "春季休業")

    # B) 夏季休業 (春3予備日の翌日～秋学期開始日前日)
    if spring3_bytes:
        ws3 = load_workbook(io.BytesIO(spring3_bytes.getvalue())).active
        dates3 = [
            (c.date() if isinstance(c, datetime) 
             else datetime.strptime(str(c), "%Y-%m-%d").date())
            for r in range(2, ws3.max_row + 1)
            if (c := ws3[f"V{r}"].value)
        ]
        if dates3:
            last3 = max(dates3)
            st3, ed3 = last3 + timedelta(days=1), start_date_autumn - timedelta(days=1)
            if st3 <= ed3:
                add_period_remarks(st3, ed3, "夏季休業")

    # C) 冬期休業 (秋2予備日の翌日～秋C開始日前日)
    if autumn2_bytes:
        ws2 = load_workbook(io.BytesIO(autumn2_bytes.getvalue())).active
        dates2 = [
            (c.date() if isinstance(c, datetime) else
            datetime.strptime(str(c), "%Y-%m-%d").date())
            for r in range(2, ws2.max_row + 1)
            if (c := ws2[f"T{r}"].value)                  # 秋Bモジュール予備日 (T列)
        ]
        if dates2:
            last2 = max(dates2)

            # ① フォームで選択された秋C開始日を最優先
            autumn3_start = st.session_state.get("autumn3_start")

            # ② フォームが空(None)なら、従来どおり V 列の最初の日付で代替
            if autumn3_start is None:
                autumn3_start = min(
                    [
                        (c.date() if isinstance(c, datetime) else
                        datetime.strptime(str(c), "%Y-%m-%d").date())
                        for r in range(2, autumn_ws.max_row + 1)
                        if (c := autumn_ws[f"V{r}"].value)          # 秋C予備日 (V列)
                    ],
                    default=None
                )

            # ③ 冬休み区間を作成
            if autumn3_start:
                st2 = last2 + timedelta(days=1)           # 休業開始 = 秋B予備日の翌日
                ed2 = autumn3_start - timedelta(days=1)   # 休業終了 = 秋C開始日の前日
                if st2 <= ed2:
                    add_period_remarks(st2, ed2, "冬期休業")


    # D) 年度末春季休業 (秋3予備日の翌日～3/31)
    if autumn3_bytes:
        wsa3 = load_workbook(io.BytesIO(autumn3_bytes.getvalue())).active
        dates_a3 = [
            (c.date() if isinstance(c, datetime)
             else datetime.strptime(str(c), "%Y-%m-%d").date())
            for r in range(2, wsa3.max_row + 1)
            if (c := wsa3[f"V{r}"].value)
        ]
        if dates_a3:
            last_a3 = max(dates_a3)
            st4, ed4 = last_a3 + timedelta(days=1), date(year_start + 1, 3, 31)
            if st4 <= ed4:
                add_period_remarks(st4, ed4, "春季休業")

    return event_labels_remark





def make_calendar_6months(year_start: int) -> bytes:
    """
    6か月×1枚 のフォーマット通りに学年暦を生成して bytes で返す。
    （例: 4月～9月を一枚に収めるレイアウト）
    """
    global cell2date
    import re    
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.worksheet.pagebreak import Break
    

    # Streamlit 側の session_state から必要なバイト列や日付情報を取り出す
    spring3_bytes   = st.session_state["spring3_excel_bytes"]
    autumn3_bytes   = st.session_state["autumn3_excel_bytes"]
    autumn2_bytes   = st.session_state.get("autumn2_excel_bytes", None)
    start_date_summer = st.session_state["start_summer"]
    start_date_autumn  = st.session_state["start_autumn"]
    event_periods     = st.session_state.get("event_periods", {})
    manual_remark     = st.session_state.get("event_labels_remark", {})

    spring_ws = load_workbook(io.BytesIO(spring3_bytes.getvalue())).active
    autumn_ws = load_workbook(io.BytesIO(autumn3_bytes.getvalue())).active
    #busy_dates = collect_busy_dates(spring_ws, autumn_ws)

    # ❶ session_state から祝日名辞書を取得
    holiday_names = st.session_state.get("holiday_names", {})   
    
    # ➀ 備考データをまとめる
    event_labels_remark = collect_event_labels_remark(
        year_start,
        spring3_bytes,
        autumn3_bytes,
        autumn2_bytes,
        start_date_summer,
        start_date_autumn,
        event_periods,
        manual_remark
        #busy_dates
        )

    

    # ➁ 新規 Workbook＋ワークシート作成
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet()

    # ── タイトル行 ──
    # タイトル
    ws.insert_rows(1)
    ws.merge_cells("A1:N1")
    ws["A1"].value = f"{year_start}年度 筑波大学 学年暦"
    ws["A1"].font  = Font(size=16, bold=True)
    ws["A1"].alignment = Alignment("center", "center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("O1:P1")
    ws["O1"].value = "学群/大学院(筑波キャンパス)"
    ws["O1"].font  = Font(size=10)
    ws["O1"].alignment = Alignment(horizontal="right", vertical="center")

    # ── 印刷設定 ──
    ws.page_setup.orientation   = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth    = 1
    ws.page_setup.fitToHeight   = 2  # ※6か月は縦に２枚分で印刷させる
    ws.page_setup.paperSize     = ws.PAPERSIZE_A4
    ws.print_title_rows         = '2:2'

    # ── 曜日ヘッダー ──
    #days = ['日','月','火','水','木','金','土']
    for i, d in enumerate("日月火水木金土"):
        c = 2 + 2*i
        ws.merge_cells(start_row=2, start_column=c, end_row=2, end_column=c+1)
        hdr = ws.cell(2, c); hdr.value = d
        hdr.font = Font(size=11, bold=True)
        hdr.alignment = Alignment("center", "center")

    # ── 備考ヘッダー ──
    ws.merge_cells("P2")
    rhead = ws["P2"]
    rhead.value = "備考"
    rhead.font = Font(size=11, bold=True)
    rhead.alignment = Alignment("center","center")

    # ── 日付セル描画＋月名＋備考欄 ──
    row_ptr = 3
    red  = Font(color="9C0006", size=11)
    std  = Font(size=11)
    sm10 = Font(size=10)
    holiday_bold = Font(color="9C0006", size=11, bold=True)

    # 【ポイント】6か月だけループさせる → 4月～9月
    for m_off in range(12):                       # 4 月 → 翌年 3 月
        m = (3 + m_off) % 12 + 1                  # 4→5→…→12→1→…→3
        y = year_start if m >= 4 else year_start + 1
        first = date(y, m, 1)
        start_wd = first.isoweekday() % 7
        last_day = calendar.monthrange(y, m)[1]
        dates = [first + timedelta(days=i) for i in range(last_day)]

        sr = row_ptr
        ptr = row_ptr

        # 空白セル（その月の初日の曜日まで）
        for i in range(start_wd):
            ws.cell(row=ptr, column=2 + 2*i, value="")
            ws.cell(row=ptr, column=3 + 2*i, value="")

        # 日付セル描画
        for d in dates:
            ccol = d.isoweekday() % 7
            cell = ws.cell(row=ptr, column=2 + 2*ccol, value=d.day)
            cell.alignment = Alignment("center","center")

            if jpholiday.is_holiday(d):
                # 祝日の場合 → 赤・太字フォント
                cell.font = holiday_bold
            elif d.weekday() == 6:
                # d.weekday()==6 → 日曜日
                cell.font = red
            # elif d.weekday() == 0:
            #     # 日曜日も赤フォント（お好みで日曜も太字にしたい場合は条件を追加してください）
            #     cell.font = red
            else:
                # それ以外の平日は標準フォント
                cell.font = std
            
            # ❷ 祝日名を書き込む  ────────────────
            if d in holiday_names:
                lbl = ws.cell(ptr, 2 + 2*ccol + 1)
                lbl.value = holiday_names[d]
                lbl.font  = Font(size=9, color="9C0006")
                lbl.alignment = Alignment("center", "center", wrap_text=True)

            if ccol == 6:
                    ptr += 1

        # ───────────────────────────────────────────────────
        # ここで、一旦 er = ptr としますが……
        er = ptr

        # ★ 9 月が終わる行で改ページ
        if m == 9:
            ws.row_breaks.append(Break(id=er))


        # # 改ページ：9月月末（m == 9 のときだけ行ブレークを挿入）
        # if m == 9:
        #     ws.row_breaks.append(Break(id=er))

        # 月名セル
        ws.merge_cells(
            start_row=sr, start_column=1,
            end_row=er,   end_column=1
        )
        mon_cell = ws.cell(sr, 1); mon_cell.value = f"{m}月"
        mon_cell.alignment = Alignment("center", "center")
        mon_cell.font = std


        # 備考欄をマージして書き込む
        ws.merge_cells(start_row=sr, start_column=16, end_row=er, end_column=16)
        rc = ws.cell(row=sr, column=16)
        rc.font = sm10
        rc.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        items = [
            (d, txt)
            for d, txts in event_labels_remark.items()
            if d.month == m
            for txt in txts
        ]
        items.sort(key=lambda x: x[0])
        if items:
            rc.value = "\n".join(txt for _, txt in items)

        row_ptr = er + 1
        


    # ── モジュール色付け / 振替授業（紫） / 予備日（灰色塗り） ──
    # ── モジュール色付けの対象日取得 ──
        y_fill = PatternFill("solid", fgColor="FFFFE599")
        b_fill = PatternFill("solid", fgColor="FFB8E1FC")
        g_fill = PatternFill("solid", fgColor="FFC6E0B4")

        def extract(col_cond, col_date, fill, sheet):
            out = []
            for r in range(2, sheet.max_row+1):
                cv = sheet[f"{col_cond}{r}"].value
                dv = sheet[f"{col_date}{r}"].value
                if cv in (None, "None", ""): continue
                if isinstance(dv, datetime):
                    dt = dv.date()
                else:
                    try:
                        dt = datetime.strptime(str(dv), "%Y-%m-%d").date()
                    except:
                        continue
                out.append((dt.year, dt.month, dt.day, fill))
            return out

        all_dates = []
        for cond, date_col, fill in [
            ("D","A", y_fill), ("I","F", b_fill), ("N","K", g_fill)
        ]:
            all_dates += extract(cond, date_col, fill, spring_ws)
            all_dates += extract(cond, date_col, fill, autumn_ws)

        # ── 割当ラベルと色付け ──
        def match_mon(v):
            if isinstance(v, str):
                m = re.match(r"(\d{1,2})月", v)
                if m: return int(m.group(1))
            return None

        def fill_cells(styles):
            # ① 振替授業ラベルを作る辞書(assign) ──
            def build_dict(sh, dcol, acol):
                d = {}
                for r in range(2, sh.max_row + 1):
                    dv = sh[f"{dcol}{r}"].value
                    av = sh[f"{acol}{r}"].value
                    if not av:
                        continue
                    # 日付を取得
                    if isinstance(dv, datetime):
                        dt0 = dv.date()
                    else:
                        try:
                            dt0 = datetime.strptime(str(dv), "%Y-%m-%d").date()
                        except:
                            continue

                    # “火曜13” のような文字列から “曜” を削除して「火13」にする
                    label = str(av).replace("曜", "")
                    d[(dt0.year, dt0.month, dt0.day)] = label
                return d

            assign = {}
            for sh in (spring_ws, autumn_ws):
                for dcol, acol in (("A", "D"), ("F", "I"), ("K", "N")):
                    assign.update(build_dict(sh, dcol, acol))

            # ② 本体：月ラベル行を探して 1～6 週目を着色／文字入れ
            # （styles は [(yy, mm, dd, fill), …] というリストを想定）
            for r in range(1, ws.max_row + 1):
                for base_col in ("A", "Q"):
                    header = ws[f"{base_col}{r}"].value
                    if not isinstance(header, str):
                        continue
                    m = re.match(r"(\d{1,2})月", header)
                    if not m:
                        continue

                    # ── この行 r が「月ラベル行」 ──
                    mon = int(m.group(1))
                    year_for_mon = year_start if mon >= 4 else year_start + 1

                    # 日付セルが並んでいる列を決定
                    if base_col == "A":
                        day_cols = ['B', 'D', 'F', 'H', 'J', 'L', 'N']
                    else:
                        day_cols = ['T', 'V', 'X', 'Z', 'AB']

                    # 週ごとに off=0～5（＝1週目～6週目）を処理
                    for off in range(6):
                        rr = r + off  # 実際に着色すべき行番号

                        # 【変更点】6週目(off==5) の場合は前週との連続性だけチェック
                        if off == 5:
                            skip_this_week = True
                            for chk_col in day_cols:
                                cur_val = ws[f"{chk_col}{rr}"].value
                                if isinstance(cur_val, (int, float)):
                                    cur_d = int(cur_val)
                                    prev_cell = ws[f"{chk_col}{r + 4}"].value  # off=4 のセル
                                    if isinstance(prev_cell, (int, float)):
                                        prev_d = int(prev_cell)
                                        if prev_d + 7 == cur_d:
                                            skip_this_week = False
                                            break
                            if skip_this_week:
                                # ６週目だが「前週と連続していない」→ スキップ
                                continue

                        # off＝0～4 または（連続している off=5）の場合、日付と styles を照合して着色＋文字入れ
                        for dc in day_cols:
                            cell = ws[f"{dc}{rr}"]
                            try:
                                dnum = int(cell.value)
                            except:
                                continue

                            for yy, mm0, dd0, fill0 in styles:
                                if mm0 == mon and dd0 == dnum:
                                    tgt_col = get_column_letter(column_index_from_string(dc) + 1)
                                    tgt = ws[f"{tgt_col}{rr}"]
                                    tgt.fill = fill0

                                    # 振替授業ラベルがあれば文字を上書き（“曜”は既に削除済み）
                                    key = (yy, mm0, dd0)
                                    if key in assign:
                                        tgt.value = assign[key]
                                        tgt.alignment = Alignment("center", "center")
                                    break

                    # 同じ行 r で A列 と Q列 の両方に“〇月”は入らない想定なので、
                    # 一度マッチしたらこの base_col のループを抜けて次の行へ
                    break


        # ── 実際に「all_dates (＝styles)」を渡して呼び出す ──
        fill_cells(all_dates)

        # ─── ここから「振替授業（紫塗り＋ラベル書き込み）」を追加する ───
    purple = PatternFill("solid", fgColor="FFD3B8F5")

    def compensations(sh):
        out = []
        for dcol, ocol, acol in (("A","B","D"), ("F","G","I"), ("K","L","N")):
            for r in range(2, sh.max_row + 1):
                dv = sh[f"{dcol}{r}"].value
                ov = sh[f"{ocol}{r}"].value
                av = sh[f"{acol}{r}"].value
                if not (dv and ov and av):
                    continue
                m = re.match(r"[月火水木金]", str(av))
                # 「実際の曜日」と同じなら振替授業ではない
                if not m or ov == m.group(0) + "曜":
                    continue
                try:
                    dt = dv.date() if isinstance(dv, datetime) else datetime.strptime(str(dv), "%Y-%m-%d").date()
                except:
                    continue
                # “火曜13” → “火13” に
                label = str(av).replace("曜", "")
                out.append((dt, label))
        return out

    for sh in (spring_ws, autumn_ws):
        for dt, wd in compensations(sh):
            placed = False
            for r in range(1, ws.max_row + 1):
                if placed:
                    break
                for base_col in ('A','Q'):
                    if match_mon(ws[f"{base_col}{r}"].value) != dt.month:
                        continue
                    cols = ['D','F','H','J','L'] if base_col == 'A' else ['T','V','X','Z','AB']
                    for off in range(6):
                        rr = r + off
                        for dc in cols:
                            val = ws[f"{dc}{rr}"].value
                            if val == dt.day:
                                tgt_col = get_column_letter(column_index_from_string(dc) + 1)
                                tgt = ws[f"{tgt_col}{rr}"]
                                tgt.fill = purple
                                tgt.value = wd
                                tgt.alignment = Alignment("center", "center")
                                placed = True
                                break
                        if placed:
                            break
                    if placed:
                        break
                if placed:
                    break

        # ─── 以下は “予備日（灰色）” を塗る部分 ───
        def paint_reserve_days():
            reserve_fill = PatternFill("solid", fgColor="FFBFBFBF")
            marks = []

            # (1) 予備日のリストを作成
            for sh, c in (
                (spring_ws, "R"), (spring_ws, "T"), (spring_ws, "V"),
                (autumn_ws, "R"), (autumn_ws, "T"), (autumn_ws, "V"),
            ):
                for rr in range(2, sh.max_row + 1):
                    raw = sh[f"{c}{rr}"].value
                    if not raw:
                        continue
                    # datetime 型の場合は .date()、文字列ならパース
                    if isinstance(raw, datetime):
                        d = raw.date()
                    else:
                        try:
                            d = datetime.strptime(str(raw), "%Y-%m-%d").date()
                        except:
                            continue
                    # (年, 月, 日) のタプルとして追加
                    marks.append((d.year, d.month, d.day))

            # (2) カレンダー本体をループして灰色で塗る
            for _, mm, dd in marks:
                for r2 in range(1, ws.max_row + 1):
                    for base_col in ("A", "Q"):
                        header = ws[f"{base_col}{r2}"].value
                        if not isinstance(header, str):
                            continue
                        m2 = re.match(r"(\d{1,2})月", header)
                        if not m2 or int(m2.group(1)) != mm:
                            continue

                        if base_col == "A":
                            day_cols2 = ['B', 'D', 'F', 'H', 'J', 'L', 'N']
                        else:
                            day_cols2 = ['T', 'V', 'X', 'Z', 'AB']

                for off2 in range(6):
                    rr2 = r2 + off2

                    # ６週目(off2==5) のときは前週との連続性をチェック
                    if off2 == 5:
                        skip2 = True
                        for chk_col2 in day_cols2:
                            cur_val2 = ws[f"{chk_col2}{rr2}"].value
                            if isinstance(cur_val2, (int, float)):
                                cur_d2 = int(cur_val2)
                                prev_cell2 = ws[f"{chk_col2}{r2 + 4}"].value
                                if isinstance(prev_cell2, (int, float)):
                                    prev_d2 = int(prev_cell2)
                                    if prev_d2 + 7 == cur_d2:
                                        skip2 = False
                                        break
                        if skip2:
                            continue  # 「６週目だが連続でない」→ スキップ

                    # off2=0～4 または「連続している off2=5」の場合、灰色で塗る
                    for dc2 in day_cols2:
                        ccell = ws[f"{dc2}{rr2}"]
                        try:
                            if int(ccell.value) == dd:
                                tgt_col2 = get_column_letter(column_index_from_string(dc2) + 1)
                                tgt2 = ws[f"{tgt_col2}{rr2}"]
                                tgt2.fill = reserve_fill
                                tgt2.value = "予備日"
                                tgt2.alignment = Alignment("center", "center")
                                painted = True
                                break
                        except:
                            continue
                    if painted:
                        break
                if painted:
                    break
            if painted:
                break

        # まとめた関数を呼び出す
        paint_reserve_days()

    # ── 休講日 を出力（「休講」 という文字列） ──
    from openpyxl.utils import get_column_letter as _get_col, column_index_from_string as _col_idx
    for d in st.session_state.manual_holidays_all:
        # 【変更点】土日・祝日ならスキップする
        #  平日かつ祝日でなければ ↓ の処理を行う
        if d.weekday() >= 5:  # 土曜＝5, 日曜＝6
            continue
        if jpholiday.is_holiday(d):
            continue

        year, month, day = d.year, d.month, d.day

        done = False
        for r in range(1, ws.max_row + 1):
            for base_col in ('A', 'Q'):
                # 月ヘッダーか？（例: "4月" 等が入っているセルを探す）
                label = ws[f"{base_col}{r}"].value
                if match_mon(label) != month:
                    continue

                day_cols = ['B','D','F','H','J','L','N'] if base_col=='A' else ['T','V','X','Z','AB']
                for off in range(5):
                    rr = r + off
                    for dc in day_cols:
                        val = ws[f"{dc}{rr}"].value
                        if isinstance(val, (int, float)) and int(val) == day:
                            tgt_col = _get_col(_col_idx(dc) + 1)
                            tgt = ws[f"{tgt_col}{rr}"]
                            tgt.value = "休講"
                            tgt.font = red
                            tgt.alignment = Alignment("center","center")
                            done = True
                            break
                    if done: break
                if done: break
            if done: break


    if cell2date is None:
        try:
            cell2date = build_date_map(ws, year_start)
            st.session_state["calendar_date_map_master"] = cell2date
        except Exception as e:
            st.warning(f"build_date_map で例外: {e}")



    # ── フォントサイズ／列幅・行高・罫線 ──
    def px_to_col(px): return (px - 5) / 7
    ws.column_dimensions['A'].width = px_to_col(50)
    # B D F H J L N を 24px、C E G I K M O を 48px
    small = px_to_col(28)   # ≒ 2.7
    large = px_to_col(44)   # ≒ 6.1

    for col in "BDFHJLN":
        ws.column_dimensions[col].width = small

    for col in "CEGIKMO":
        ws.column_dimensions[col].width = large

    ws.column_dimensions['P'].width = px_to_col(143)

    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 24
    thin = Side(border_style="thin", color="000000")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(2, ws.max_row + 1):
        for c in range(1, 17):
            ws.cell(row=r, column=c).border = bd

    # ── バイト列を返す ──
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()














#-----------------------------------------------------結果-----------------------------------------------------
import streamlit as st

st.title("📥 学年暦の作成＆ダウンロード")

# （前提として、st.session_state には
#   "spring3_excel_bytes","autumn3_excel_bytes","autumn2_excel_bytes" などが入っているものとします）

col1, col2 = st.columns(2)

with col1:
    if "autumn3_excel_bytes" in st.session_state:
        data4 = make_calendar_4months(st.session_state["year_start"])
        st.download_button(
            label=f"📥 {st.session_state['year_start']}年度 学年暦 (1ページに4か月表示)",
            data=data4,
            file_name=f"{st.session_state['year_start']}_学年暦_4m.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_4m_calendar"
        )

with col2:
    if "autumn3_excel_bytes" in st.session_state:
        data6 = make_calendar_6months(st.session_state["year_start"])
        st.download_button(
            label=f"📥 {st.session_state['year_start']}年度 学年暦 (1ページに6か月表示)",
            data=data6,
            file_name=f"{st.session_state['year_start']}_学年暦_6m.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_6m_calendar"
        )










# #------------------------------------.icsファイルの作成-----------------------------------------

import io
import pandas as pd
from datetime import timedelta
from icalendar import Calendar, Event
from openpyxl import load_workbook
import streamlit as st

# --- 振替が必要な授業日の抽出 ---
def extract_mismatch_assignments(sheet, triplets):
    mismatches = []
    for date_col, actual_col, assign_col in triplets:
        for row in range(2, sheet.max_row + 1):
            date_val   = sheet[f"{date_col}{row}"].value
            actual_val = sheet[f"{actual_col}{row}"].value
            assign_val = sheet[f"{assign_col}{row}"].value
            if not (date_val and actual_val and assign_val):
                continue
            try:
                d = pd.to_datetime(date_val)
            except:
                continue
            if str(actual_val).strip()[:2] != str(assign_val).strip()[:2]:
                mismatches.append((d.date(), str(assign_val).strip()))
    return mismatches

# --- .ics 作成（手動イベント＋振替授業）---
def create_ics(mismatches, event_periods, manual_holidays_all):
    """
    ・mismatches: 振替が必要な授業日リスト [(date, "火曜3"), …]
    ・event_periods: { "イベント名 (2)": [date1, date2,…], … }
    ・manual_holidays_all: 休講日として扱う日（set of date）
    """
    cal = Calendar()
    cal.add('prodid', '-//筑波大 学年暦//')
    cal.add('version', '2.0')

    # まず、手動イベント（休講イベント／イベントのみ）を ICS に追加
    for name_with_suffix, dates in event_periods.items():
        # 末尾に "(数字)" がついていたら除去して base_name を作成
        base_name = re.sub(r"\s*\(\d+\)$", "", name_with_suffix)

        for d in dates:
            ev = Event()
            # 休講日扱いの日は「休講：○○」それ以外は単にベース名だけ
            if d in manual_holidays_all:
                summary = f"休講：{base_name}"
            else:
                summary = base_name
            ev.add('summary', summary)
            ev.add('dtstart', d)
            ev.add('dtend',   d + timedelta(days=1))
            ev.add('transp',  'OPAQUE')
            cal.add_component(ev)

    # 次に、振替授業イベントを追加
    for date_obj, label in mismatches:
        ev = Event()
        summary = f"【振替】{label[:2]}授業"
        ev.add('summary', summary)
        ev.add('dtstart', date_obj)
        ev.add('dtend',   date_obj + timedelta(days=1))
        ev.add('transp',  'OPAQUE')
        cal.add_component(ev)

    return cal.to_ical()


# --- メイン：ICS 生成とダウンロードボタン表示 --- 
spring_bytes  = st.session_state.get("spring3_excel_bytes")
autumn_bytes  = st.session_state.get("autumn3_excel_bytes")

if not spring_bytes or not autumn_bytes:
    st.warning("「学年暦を作成」で生成してください。")
else:
    # spring3, autumn3 のシートを読み込む
    spring_ws = load_workbook(io.BytesIO(spring_bytes.getvalue())).active
    autumn_ws = load_workbook(io.BytesIO(autumn_bytes.getvalue())).active

    triplets = [("A","B","D"), ("F","G","I"), ("K","L","N")]
    mismatches = []
    mismatches += extract_mismatch_assignments(spring_ws, triplets)
    mismatches += extract_mismatch_assignments(autumn_ws, triplets)

    # ICS 作成
    ics_bytes = create_ics(
        mismatches,
        st.session_state["event_periods"],
        st.session_state["manual_holidays_all"]
    )

    st.download_button(
        label=f"📥 {st.session_state['year_start']}年度 カレンダー用 ICS をダウンロード",
        data=ics_bytes,
        file_name=f"{st.session_state['year_start']}_学年暦.ics",
        mime="text/calendar",
        key="download_ics"
    )

