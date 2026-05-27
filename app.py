"""
委員長面談スケジューラー（Streamlit版）
ブラウザ・スマホ対応
"""
import streamlit as st
import os, re, random
from datetime import datetime, date, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── 設定 ──────────────────────────────────────────────────
SHEET_ID    = st.secrets.get("sheet_id", "1ZcD0rOlRUsq-SOgsAzVgufZ7OsqITGNtbMnuL0QtJc0")
SHEET_MEET  = "Sheet1"
SHEET_MEIBO = "OUTPUT (43)"
SCOPES      = ["https://www.googleapis.com/auth/spreadsheets"]
SLOT_MIN    = 25
BREAK_MIN   = 5

_DIR       = os.path.dirname(os.path.abspath(__file__))
SA_FILE = os.path.join(_DIR, "service_account.json")

PREFECTURES = [
    "北海道","青森","岩手","宮城","秋田","山形","福島",
    "茨城","栃木","群馬","埼玉","千葉","東京","神奈川",
    "新潟","富山","石川","福井","山梨","長野","岐阜",
    "静岡","愛知","三重","滋賀","京都","大阪","兵庫",
    "奈良","和歌山","鳥取","島根","岡山","広島","山口",
    "徳島","香川","愛媛","高知","福岡","佐賀","長崎",
    "熊本","大分","宮崎","鹿児島","沖縄",
]

# ── ユーティリティ ─────────────────────────────────────────
def is_pref_honbu(org):
    return any(org.strip().startswith(p) for p in PREFECTURES)

def _normalize(text):
    table = str.maketrans("０１２３４５６７８９：／　", "0123456789:/ ")
    return text.translate(table)

def _extract_date(text):
    m = re.search(r'(\d{1,2})月(\d{1,2})日?', text)
    if not m:
        m = re.search(r'(\d{1,2})[/](\d{1,2})', text)
    if not m:
        return None
    month, day = int(m.group(1)), int(m.group(2))
    try:
        weekdays = "月火水木金土日"
        dow = weekdays[date(2026, month, day).weekday()]
        label = f"{month}月{day}日（{dow}）"
    except Exception:
        label = f"{month}月{day}日"
    return month, day, label

def _extract_times(text):
    found = []
    for m in re.finditer(r'(\d{1,2}):(\d{2})', text):
        h, mi = int(m.group(1)), int(m.group(2))
        if 0 <= h <= 23 and 0 <= mi <= 59:
            found.append((m.start(), datetime(2000, 1, 1, h, mi)))
    prev_ampm = None
    for m in re.finditer(r'(午前|午後)?(\d{1,2})時(\d{1,2}分)?', text):
        if any(abs(pos - m.start()) < 4 for pos, _ in found):
            continue
        ampm = m.group(1)
        h    = int(m.group(2))
        mi   = int(re.search(r'\d+', m.group(3)).group()) if m.group(3) else 0
        if ampm:
            prev_ampm = ampm
        elif prev_ampm:
            ampm = prev_ampm
        if ampm == "午後" and h != 12:
            h += 12
        elif ampm == "午前" and h == 12:
            h = 0
        elif not ampm and 1 <= h <= 7:
            h += 12
        if 0 <= h <= 23:
            found.append((m.start(), datetime(2000, 1, 1, h, mi)))
    found.sort(key=lambda x: x[0])
    return [dt for _, dt in found[:2]]

def parse_time_block(line):
    text = _normalize(line)
    result = _extract_date(text)
    if not result:
        return None
    _, _, label = result
    times = _extract_times(text)
    if len(times) < 2 or times[0] >= times[1]:
        return None
    return label, times[0], times[1]

def parse_filters(line):
    text = _normalize(line.strip())
    if not text:
        return None, None, None, None
    filter_pref = None
    age_min = age_max = None
    for pref in PREFECTURES:
        if pref in text:
            filter_pref = pref
            text = text.replace(pref, '')
            break
    m = re.search(r'(\d+)代', text)
    if m:
        d = int(m.group(1))
        age_min, age_max = d, d + 9
        text = text[:m.start()] + text[m.end():]
    if age_min is None:
        m = re.search(r'(\d+)\s*[〜~\-ー]\s*(\d+)\s*歳?', text)
        if m:
            age_min, age_max = int(m.group(1)), int(m.group(2))
            text = text[:m.start()] + text[m.end():]
    m = re.search(r'(\d+)歳以上', text)
    if m:
        age_min = int(m.group(1)); text = text[:m.start()] + text[m.end():]
    m = re.search(r'(\d+)歳以下', text)
    if m:
        age_max = int(m.group(1)); text = text[:m.start()] + text[m.end():]
    dept = re.sub(r'[からでのをにはが以上以下　\s]', '', text).strip()
    filter_dept = dept if dept else None
    return filter_pref, filter_dept, age_min, age_max

def calc_slots(label, ts, te):
    slots, unit, dur = [], timedelta(minutes=SLOT_MIN + BREAK_MIN), timedelta(minutes=SLOT_MIN)
    cur = ts
    while cur + dur <= te:
        slots.append({
            "date":  label,
            "start": cur.strftime("%H:%M"),
            "end":   (cur + dur).strftime("%H:%M"),
        })
        cur += unit
    return slots

def col_letter(idx):
    r, i = "", idx + 1
    while i > 0:
        i, rem = divmod(i - 1, 26)
        r = chr(65 + rem) + r
    return r

# ── Google認証（サービスアカウント）──────────────────────────
@st.cache_resource
def get_service():
    # クラウド環境：Streamlit Secretsから読む（secrets.tomlがある場合）
    try:
        info = st.secrets["gcp_service_account"]
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES
        )
        return build("sheets", "v4", credentials=creds)
    except (KeyError, FileNotFoundError):
        pass  # ローカル環境へフォールバック

    # ローカル環境：service_account.jsonから読む
    if os.path.exists(SA_FILE):
        creds = service_account.Credentials.from_service_account_file(
            SA_FILE, scopes=SCOPES
        )
        return build("sheets", "v4", credentials=creds)

    st.error("❌ 認証情報が見つかりません（service_account.json を確認してください）")
    st.stop()

# ── データ読み込み（5分キャッシュ）─────────────────────────
@st.cache_data(ttl=300)
def load_data():
    svc = get_service()

    # 名簿
    rows = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"{SHEET_MEIBO}!A:L"
    ).execute().get("values", [])
    meibo = {}
    today = date.today()
    for row in rows[1:]:
        if not row: continue
        name = row[0].strip() if row else ""
        if not name: continue
        birth_str = row[1].strip() if len(row) > 1 else ""
        org       = row[5].strip() if len(row) > 5 else ""
        age = None
        if birth_str:
            try:
                bd  = datetime.strptime(birth_str, "%Y/%m/%d").date()
                age = today.year - bd.year - ((today.month, today.day) < (bd.month, bd.day))
            except Exception:
                pass
        org_type = "都道府県本部" if is_pref_honbu(org) else "党本部"
        meibo[name] = {"org": org, "age": age, "org_type": org_type}

    # 面談シート
    rows2 = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"{SHEET_MEET}!A:ZZ"
    ).execute().get("values", [])
    if not rows2:
        return meibo, [], []
    header = rows2[0]
    data = []
    for i, row in enumerate(rows2[2:], start=3):
        name = row[0].strip() if row else ""
        if not name: continue
        met = any(
            ci < len(row) and row[ci].strip() in ("〇", "○", "o", "O")
            for ci in range(1, len(header))
        )
        data.append({"name": name, "row_idx": i, "met": met})
    return meibo, header, data

# ── スプレッドシート書き込み ────────────────────────────────
def find_or_create_date_col(svc, header, date_label):
    nums = re.findall(r'\d+', date_label)
    for i in range(1, len(header)):
        hn = re.findall(r'\d+', header[i])
        if len(nums) >= 2 and len(hn) >= 2 and nums[0] == hn[0] and nums[1] == hn[1]:
            return i
    new_col = len(header)
    cell = f"{SHEET_MEET}!{col_letter(new_col)}1"
    svc.spreadsheets().values().update(
        spreadsheetId=SHEET_ID, range=cell,
        valueInputOption="RAW", body={"values": [[date_label]]}
    ).execute()
    header.append(date_label)
    return new_col

def write_marks(svc, header, assignments):
    date_label = assignments[0][0]["date"]
    col_idx    = find_or_create_date_col(svc, header, date_label)
    updates    = []
    for slot, person in assignments:
        cell = f"{SHEET_MEET}!{col_letter(col_idx)}{person['row_idx']}"
        updates.append({"range": cell, "values": [["〇"]]})
    if updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"valueInputOption": "RAW", "data": updates}
        ).execute()

def clear_marks(svc, header, assignments):
    date_label = assignments[0][0]["date"]
    nums       = re.findall(r'\d+', date_label)
    col_idx    = None
    for i in range(1, len(header)):
        hn = re.findall(r'\d+', header[i])
        if len(nums) >= 2 and len(hn) >= 2 and nums[0] == hn[0] and nums[1] == hn[1]:
            col_idx = i
            break
    if col_idx is None:
        st.error("列が見つからないため取り消しできません")
        return
    ranges = [f"{SHEET_MEET}!{col_letter(col_idx)}{p['row_idx']}" for _, p in assignments]
    svc.spreadsheets().values().batchClear(
        spreadsheetId=SHEET_ID, body={"ranges": ranges}
    ).execute()

# ── セッション初期化 ───────────────────────────────────────
def init_session():
    defaults = {
        "step": "mode",
        "mode": "new",
        "date_label": None,
        "ts": None, "te": None,
        "slots": None,
        "n_pref": 0, "n_party": 0,
        "filter_pref": None, "filter_dept": None,
        "age_min": None, "age_max": None,
        "assignments": None,
        "marks_written": False,
        "header": None,
        "unmet": None,
        "meibo": None,
        "pref_count": 0,
        "party_count": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

# ── UI ────────────────────────────────────────────────────

def show_mode_select():
    """モード選択画面"""
    meibo, header, data = load_data()
    unmet       = [p for p in data if not p["met"]]
    met_count   = len(data) - len(unmet)
    pref_count  = sum(1 for p in unmet if meibo.get(p["name"], {}).get("org_type") == "都道府県本部")
    party_count = len(unmet) - pref_count

    col1, col2, col3 = st.columns(3)
    col1.metric("全体",    f"{len(data)}名")
    col2.metric("面談済み", f"{met_count}名")
    col3.metric("未面談",   f"{len(unmet)}名")
    st.caption(f"未面談内訳：都道府県本部 {pref_count}名 ／ 党本部 {party_count}名")

    st.divider()
    st.subheader("モードを選択")

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        if st.button("📋 新規スケジュール", use_container_width=True, type="primary"):
            st.session_state.update(
                mode="new", step="time",
                meibo=meibo, header=list(header),
                unmet=unmet, pref_count=pref_count, party_count=party_count,
                marks_written=False, assignments=None
            )
            st.rerun()
    with col_b:
        if st.button("🔄 リスケ\n（〇を削除）", use_container_width=True):
            st.session_state.update(mode="reschedule", step="reschedule")
            st.rerun()
    with col_c:
        if st.button("🧪 テストモード\n（書き込みなし）", use_container_width=True):
            st.session_state.update(
                mode="test", step="time",
                meibo=meibo, header=list(header),
                unmet=unmet, pref_count=pref_count, party_count=party_count,
                marks_written=False, assignments=None
            )
            st.rerun()


def show_time_input():
    """① 空き時間入力"""
    if st.session_state.mode == "test":
        st.info("🧪 テストモード：スプレッドシートへの書き込みは行いません")

    st.subheader("① 空き時間の入力")

    with st.form("time_form"):
        time_input = st.text_input(
            "委員長の空き時間",
            placeholder="例）6月2日 13:00-15:00　/　6/2 午後1時から3時"
        )
        submitted = st.form_submit_button("次へ →", use_container_width=True, type="primary")

    if submitted:
        if not time_input.strip():
            st.error("空き時間を入力してください")
            return
        result = parse_time_block(time_input)
        if not result:
            st.error("日付と開始・終了時刻が読み取れませんでした\n例：6月2日 13:00-15:00")
            return
        date_label, ts, te = result
        slots = calc_slots(date_label, ts, te)
        if not slots:
            st.error("コマを作れませんでした（開始・終了時刻を確認してください）")
            return
        st.session_state.update(
            date_label=date_label, ts=ts, te=te, slots=slots, step="counts"
        )
        st.rerun()

    st.divider()
    if st.button("← トップへ戻る"):
        st.session_state.step = "mode"
        st.rerun()


def show_counts_input():
    """② 人数指定"""
    st.subheader("② 人数の指定")

    slots       = st.session_state.slots
    date_label  = st.session_state.date_label
    ts          = st.session_state.ts
    te          = st.session_state.te
    pref_count  = st.session_state.pref_count
    party_count = st.session_state.party_count

    st.info(
        f"📅 {date_label}　{ts.strftime('%H:%M')}〜{te.strftime('%H:%M')}\n"
        f"→ {len(slots)}コマ（各{SLOT_MIN}分）：" +
        "　".join(f"{s['start']}〜{s['end']}" for s in slots)
    )

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        n_pref = st.number_input(
            f"都道府県本部（未面談 {pref_count}名）",
            min_value=0, max_value=min(pref_count, len(slots)), value=0, step=1
        )
    with col2:
        n_party = st.number_input(
            f"党本部（未面談 {party_count}名）",
            min_value=0, max_value=min(party_count, len(slots)), step=1
        )

    total = n_pref + n_party
    if total > len(slots):
        st.warning(f"⚠️ 合計{total}名がコマ数{len(slots)}を超えています")
    elif total > 0:
        st.success(f"都道府県本部 {n_pref}名 ／ 党本部 {n_party}名　（合計 {total}名）")

    col_back, col_next = st.columns(2)
    with col_back:
        if st.button("← 戻る", use_container_width=True):
            st.session_state.step = "time"
            st.rerun()
    with col_next:
        disabled = (total == 0 or total > len(slots))
        if st.button("次へ →", use_container_width=True, type="primary", disabled=disabled):
            st.session_state.update(n_pref=n_pref, n_party=n_party, step="filters")
            st.rerun()


def show_filters_input():
    """③ 絞り込み条件"""
    st.subheader("③ 絞り込み条件（任意）")
    st.caption("都道府県・部局・年齢を指定できます。不要ならそのまま「次へ」")

    with st.form("filter_form"):
        filter_input = st.text_input(
            "絞り込み条件",
            placeholder="例）北海道　/　選挙部　/　40代　/　北海道の35〜45歳"
        )
        submitted = st.form_submit_button("次へ →", use_container_width=True, type="primary")

    if submitted:
        if filter_input.strip():
            fp, fd, amin, amax = parse_filters(filter_input)
            parts = []
            if fp:   parts.append(f"都道府県：{fp}")
            if fd:   parts.append(f"部局：{fd}")
            if amin and amax: parts.append(f"年齢：{amin}〜{amax}歳")
            elif amin:        parts.append(f"年齢：{amin}歳以上")
            elif amax:        parts.append(f"年齢：{amax}歳以下")
            if parts:
                st.info("条件：" + "　/　".join(parts))
        else:
            fp = fd = amin = amax = None

        st.session_state.update(
            filter_pref=fp, filter_dept=fd, age_min=amin, age_max=amax
        )

        assignments = do_select()
        if assignments is not None:
            st.session_state.update(assignments=assignments, step="confirm")
            st.rerun()

    col_back, _ = st.columns([1, 3])
    with col_back:
        if st.button("← 戻る"):
            st.session_state.step = "counts"
            st.rerun()


def do_select():
    """人選を実行。成功したら assignments を返す、失敗したら None"""
    unmet  = st.session_state.unmet
    meibo  = st.session_state.meibo
    slots  = st.session_state.slots
    n_pref = st.session_state.n_pref
    n_party= st.session_state.n_party
    fp     = st.session_state.filter_pref
    fd     = st.session_state.filter_dept
    amin   = st.session_state.age_min
    amax   = st.session_state.age_max

    pref_pool, party_pool = [], []
    for person in unmet:
        name = person["name"]
        info = meibo.get(name, {"org": "", "age": None, "org_type": "党本部"})
        if amin is not None and info["age"] is not None and info["age"] < amin: continue
        if amax is not None and info["age"] is not None and info["age"] > amax: continue
        if info["org_type"] == "都道府県本部":
            if fp and fp not in info["org"]: continue
            pref_pool.append(person)
        else:
            if fd and fd not in info["org"]: continue
            party_pool.append(person)

    ok = True
    if n_pref > len(pref_pool):
        st.error(f"都道府県本部の条件を満たす未面談者が{len(pref_pool)}名しかいません（{n_pref}名指定）")
        ok = False
    if n_party > len(party_pool):
        st.error(f"党本部の条件を満たす未面談者が{len(party_pool)}名しかいません（{n_party}名指定）")
        ok = False
    if not ok:
        return None

    selected = random.sample(pref_pool, n_pref) + random.sample(party_pool, n_party)
    random.shuffle(selected)
    return list(zip(slots[:len(selected)], selected))


def show_confirm():
    """④ 確認・確定画面"""
    assignments = st.session_state.assignments
    meibo       = st.session_state.meibo
    date_label  = st.session_state.date_label

    # ── 〇記入済み → 取り消し猶予 ──────────────────────────
    if st.session_state.marks_written:
        st.success("✅ スプレッドシートに〇を記入しました")
        _show_schedule_table(assignments, meibo, date_label)

        if st.session_state.mode == "test":
            st.info("🧪 テストモードのため実際には書き込んでいません")

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            if st.button("↩️ 取り消し（〇を削除）", use_container_width=True):
                if st.session_state.mode != "test":
                    svc = get_service()
                    clear_marks(svc, st.session_state.header, assignments)
                    load_data.clear()
                st.success("取り消しました")
                st.session_state.update(marks_written=False, step="mode")
                st.rerun()
        with col2:
            if st.button("🏠 トップへ戻る", use_container_width=True, type="primary"):
                load_data.clear()
                st.session_state.update(step="mode", marks_written=False)
                st.rerun()
        return

    # ── まだ書き込んでいない → 確認 ────────────────────────
    st.subheader("④ 内容を確認してください")
    _show_schedule_table(assignments, meibo, date_label)

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        label = "✅ 確定（テスト）" if st.session_state.mode == "test" else "✅ 確定（〇記入）"
        if st.button(label, use_container_width=True, type="primary"):
            if st.session_state.mode != "test":
                svc = get_service()
                write_marks(svc, st.session_state.header, assignments)
                load_data.clear()
            st.session_state.marks_written = True
            st.rerun()
    with col2:
        if st.button("🔄 人を入れ替え", use_container_width=True):
            st.session_state.step = "swap"
            st.rerun()
    with col3:
        if st.button("❌ キャンセル", use_container_width=True):
            st.session_state.step = "mode"
            st.rerun()


def _show_schedule_table(assignments, meibo, date_label):
    """スケジュール表を描画（共通）"""
    st.markdown(f"### 📅 {date_label}")
    rows = []
    for slot, person in assignments:
        info = meibo.get(person["name"], {"org": "", "age": None})
        rows.append({
            "時間":  f"{slot['start']}〜{slot['end']}",
            "氏名":  person["name"],
            "所属":  info["org"],
            "年齢":  f"{info['age']}歳" if info["age"] else "－",
        })
    st.dataframe(rows, hide_index=True, use_container_width=True)


def show_swap():
    """人の入れ替え画面"""
    st.subheader("🔄 人の入れ替え")

    assignments = st.session_state.assignments
    meibo       = st.session_state.meibo
    unmet       = st.session_state.unmet
    current     = {p["name"] for _, p in assignments}

    # 外す人
    out_opts = [
        f"{slot['start']}〜{slot['end']}　{p['name']}（{meibo.get(p['name'],{}).get('org','')}）"
        for slot, p in assignments
    ]
    sel_out = st.selectbox("外す人", out_opts)
    out_idx = out_opts.index(sel_out)

    st.divider()

    # 入れる人
    candidates = [p for p in unmet if p["name"] not in current]
    if not candidates:
        st.error("入れ替え可能な未面談者がいません")
    else:
        in_opts = [
            f"{p['name']}（{meibo.get(p['name'],{}).get('org','')}）"
            for p in candidates
        ]
        sel_in = st.selectbox("代わりに入れる人", in_opts)
        in_idx = in_opts.index(sel_in)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ 入れ替え実行", use_container_width=True, type="primary"):
                slot_out, _ = assignments[out_idx]
                new_assign  = list(assignments)
                new_assign[out_idx] = (slot_out, candidates[in_idx])
                st.session_state.update(assignments=new_assign, step="confirm")
                st.rerun()
        with col2:
            if st.button("← 戻る", use_container_width=True):
                st.session_state.step = "confirm"
                st.rerun()


def show_reschedule():
    """リスケ画面"""
    st.subheader("🔄 リスケ：〇を削除して未面談に戻す")

    with st.spinner("データを読み込んでいます..."):
        meibo, header, data = load_data()
        svc = get_service()

    rows_all = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"{SHEET_MEET}!A:ZZ"
    ).execute().get("values", [])

    # 〇が入っている人を収集
    marked = []
    for person in data:
        ri  = person["row_idx"]
        row = rows_all[ri - 1] if ri - 1 < len(rows_all) else []
        for ci in range(1, len(header)):
            if ci < len(row) and row[ci].strip() in ("〇", "○", "o", "O"):
                marked.append({
                    "name":    person["name"],
                    "row_idx": ri,
                    "col_idx": ci,
                    "date":    header[ci],
                })

    if not marked:
        st.info("〇が入っている方が見つかりません")
        if st.button("← トップへ戻る"):
            st.session_state.step = "mode"
            st.rerun()
        return

    opts = [f"{m['name']}（{m['date']}）" for m in marked]
    sel  = st.selectbox("リスケする方を選択", opts)
    idx  = opts.index(sel)
    target = marked[idx]

    st.warning(f"⚠️ **{target['name']}**（{target['date']}）の〇を削除します")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🗑️ 削除する", use_container_width=True, type="primary"):
            cell = f"{SHEET_MEET}!{col_letter(target['col_idx'])}{target['row_idx']}"
            svc.spreadsheets().values().clear(
                spreadsheetId=SHEET_ID, range=cell
            ).execute()
            load_data.clear()
            st.success(f"✅ {target['name']}（{target['date']}）の〇を削除しました")
            st.session_state.step = "mode"
            st.rerun()
    with col2:
        if st.button("← キャンセル", use_container_width=True):
            st.session_state.step = "mode"
            st.rerun()


# ── エントリーポイント ─────────────────────────────────────
def main():
    st.set_page_config(
        page_title="委員長面談スケジューラー",
        page_icon="📅",
        layout="centered",
        initial_sidebar_state="collapsed",
    )
    st.title("📅 委員長面談スケジューラー")

    init_session()

    step = st.session_state.step
    if step == "mode":
        show_mode_select()
    elif step == "time":
        show_time_input()
    elif step == "counts":
        show_counts_input()
    elif step == "filters":
        show_filters_input()
    elif step == "confirm":
        show_confirm()
    elif step == "swap":
        show_swap()
    elif step == "reschedule":
        show_reschedule()


if __name__ == "__main__":
    main()
