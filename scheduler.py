"""
委員長面談スケジューラー
====================================================
使い方:
  python scheduler.py              通常モード
  python scheduler.py --test       テストモード（書き込みなし）
  python scheduler.py --reschedule リスケ（氏名指定で〇を削除）
====================================================
"""

import os, sys, re, random
from datetime import datetime, date, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

sys.stdout.reconfigure(encoding="utf-8")

# ── 設定 ──────────────────────────────────────────────────
SHEET_ID     = "1ZcD0rOlRUsq-SOgsAzVgufZ7OsqITGNtbMnuL0QtJc0"
SHEET_MEET   = "Sheet1"          # 面談管理シート
SHEET_MEIBO  = "OUTPUT (43)"     # 名簿シート
SCOPES       = ["https://www.googleapis.com/auth/spreadsheets"]
SLOT_MIN     = 25
BREAK_MIN    = 5

_DIR         = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE   = os.path.join(_DIR, "token.json")
CREDS_FILE   = os.path.join(_DIR, "credentials.json.json")
TEST_MODE    = "--test" in sys.argv   # main()内でも更新される

# 都道府県名リスト（F列の先頭がこれなら都道府県本部）
PREFECTURES = [
    "北海道","青森","岩手","宮城","秋田","山形","福島",
    "茨城","栃木","群馬","埼玉","千葉","東京","神奈川",
    "新潟","富山","石川","福井","山梨","長野","岐阜",
    "静岡","愛知","三重","滋賀","京都","大阪","兵庫",
    "奈良","和歌山","鳥取","島根","岡山","広島","山口",
    "徳島","香川","愛媛","高知","福岡","佐賀","長崎",
    "熊本","大分","宮崎","鹿児島","沖縄",
]

def is_pref_honbu(org: str) -> bool:
    """F列の値が都道府県本部かどうか判定"""
    org = org.strip()
    return any(org.startswith(p) for p in PREFECTURES)

# ── 認証 ──────────────────────────────────────────────────
def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("トークンを更新中...")
            creds.refresh(Request())
        else:
            print("ブラウザで認証画面が開きます...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return creds

# ── データ読み込み ─────────────────────────────────────────
def load_meibo(svc) -> dict:
    """OUTPUT(43)を読み込み {氏名: {org, age, org_type}} を返す"""
    rows = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"{SHEET_MEIBO}!A:C"
    ).execute().get("values", [])

    meibo = {}
    today = date.today()
    for row in rows[1:]:
        if not row: continue
        name = row[0].strip() if len(row) > 0 else ""
        if not name: continue
        birth_str = row[1].strip() if len(row) > 1 else ""
        org       = row[2].strip() if len(row) > 2 else ""  # C列（index2）

        # 年齢計算（YYYY/MM/DD と YYYY/M 両対応）
        age = None
        if birth_str:
            try:
                parts = birth_str.strip().split('/')
                if len(parts) >= 3:
                    # YYYY/MM/DD
                    bd = date(int(parts[0]), int(parts[1]), int(parts[2]))
                    age = today.year - bd.year - ((today.month, today.day) < (bd.month, bd.day))
                elif len(parts) == 2:
                    # YYYY/M（日なし）→ 月で比較
                    by, bm = int(parts[0]), int(parts[1])
                    age = today.year - by - (today.month < bm)
            except: pass

        org_type = "都道府県本部" if is_pref_honbu(org) else "党本部"
        meibo[name] = {"org": org, "age": age, "org_type": org_type}
    return meibo

def load_meeting_sheet(svc):
    """Sheet1を読み込み (header_row, data_rows) を返す
    header_row = ['氏名', '5/13(水)', ...]
    data_rows  = [{'name':..., 'row_idx':..., 'met':bool}, ...]
    row_idxはSheet上の行番号（1始まり）
    """
    rows = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"{SHEET_MEET}!A:ZZ"
    ).execute().get("values", [])

    if not rows: return [], []
    header = rows[0]  # 行1：氏名, 5/13(水), ...

    # 行2は集計行（スキップ）、行3以降がデータ
    data = []
    for i, row in enumerate(rows[2:], start=3):  # i=シートの行番号（1始まり）
        name = row[0].strip() if row else ""
        if not name: continue
        met = any(
            ci < len(row) and row[ci].strip() in ("〇","○","o","O")
            for ci in range(1, len(header))
        )
        data.append({"name": name, "row_idx": i, "met": met})
    return header, data

# ── 時間割計算 ─────────────────────────────────────────────
def _normalize(text: str) -> str:
    """全角数字・記号を半角に統一"""
    table = str.maketrans(
        "０１２３４５６７８９：／　",
        "0123456789:/ "
    )
    return text.translate(table)

def _extract_date(text: str):
    """テキストから月・日を抽出。見つからなければ None"""
    # 6月2日 / 6月2 / 6/2
    m = re.search(r'(\d{1,2})月(\d{1,2})日?', text)
    if not m:
        m = re.search(r'(\d{1,2})[/](\d{1,2})', text)
    if not m:
        return None
    month, day = int(m.group(1)), int(m.group(2))
    try:
        from datetime import date as _date
        weekdays = "月火水木金土日"
        dow = weekdays[_date(2026, month, day).weekday()]
        label = f"{month}月{day}日（{dow}）"
    except Exception:
        label = f"{month}月{day}日"
    return month, day, label

def _extract_times(text: str) -> list:
    """テキストから時刻を最大2つ抽出して datetime のリストで返す"""
    found = []  # [(position, datetime)]

    # ① HH:MM 形式
    for m in re.finditer(r'(\d{1,2}):(\d{2})', text):
        h, mi = int(m.group(1)), int(m.group(2))
        if 0 <= h <= 23 and 0 <= mi <= 59:
            found.append((m.start(), datetime(2000, 1, 1, h, mi)))

    # ② X時Y分 / 午前X時 / 午後X時 形式
    prev_ampm = None
    for m in re.finditer(r'(午前|午後)?(\d{1,2})時(\d{1,2}分)?', text):
        # ①で既にカバーされた位置はスキップ
        if any(abs(pos - m.start()) < 4 for pos, _ in found):
            continue
        ampm = m.group(1)
        h    = int(m.group(2))
        mi   = int(re.search(r'\d+', m.group(3)).group()) if m.group(3) else 0
        # 午前/午後を前の時刻から引き継ぐ（「午後1時から3時」の3時も午後）
        if ampm:
            prev_ampm = ampm
        elif prev_ampm:
            ampm = prev_ampm
        # 時刻変換
        if ampm == "午後" and h != 12:
            h += 12
        elif ampm == "午前" and h == 12:
            h = 0
        elif not ampm and 1 <= h <= 7:
            h += 12  # 文脈なしの1〜7時は午後とみなす
        if 0 <= h <= 23:
            found.append((m.start(), datetime(2000, 1, 1, h, mi)))

    found.sort(key=lambda x: x[0])
    return [dt for _, dt in found[:2]]

def parse_time_block(line: str):
    """
    多様な形式の空き時間入力をパース
    対応例:
      6月2日 13:00 15:00
      6/2 13:00-15:00
      6月2日月曜 午後1時から3時
      6月2日（月） 13時30分 15時
    Returns (date_label, start_dt, end_dt) or None
    """
    text = _normalize(line)
    result = _extract_date(text)
    if not result:
        return None
    _, _, label = result
    times = _extract_times(text)
    if len(times) < 2:
        return None
    if times[0] >= times[1]:
        return None
    return label, times[0], times[1]

# ── 柔軟入力パーサー群 ────────────────────────────────────

def parse_counts(line, pref_avail, party_avail, n_slots):
    """
    '都道府県2 党本部1' / '2と1' / '2名1名' などから (n_pref, n_party) を返す
    失敗時は (None, エラーメッセージ)
    """
    text = _normalize(line.strip())
    nums = [int(m.group()) for m in re.finditer(r'\d+', text)]

    if not nums:
        return None, "数字が見つかりませんでした"

    if len(nums) == 1:
        n = nums[0]
        pref_kws  = ['都道府県', '県本部', '地方', '県', '道府']
        party_kws = ['党本部', '中央', '本部']
        if any(k in text for k in pref_kws):
            n_pref, n_party = n, 0
        elif any(k in text for k in party_kws):
            n_pref, n_party = 0, n
        else:
            return None, "都道府県本部・党本部どちらの人数か判断できませんでした"
    else:
        # キーワードの登場位置で順序を判定
        pref_pos  = min((text.find(k) for k in ['都道府県','県本部','地方'] if k in text), default=9999)
        party_pos = min((text.find(k) for k in ['党本部','中央'] if k in text), default=9999)
        if party_pos < pref_pos:
            n_pref, n_party = nums[1], nums[0]
        else:
            n_pref, n_party = nums[0], nums[1]  # キーワードなし or 都道府県が先

    if n_pref + n_party > n_slots:
        return None, f"合計{n_pref+n_party}名がコマ数{n_slots}を超えています"
    if n_pref > pref_avail:
        return None, f"都道府県本部の未面談者は{pref_avail}名のみです（{n_pref}名指定）"
    if n_party > party_avail:
        return None, f"党本部の未面談者は{party_avail}名のみです（{n_party}名指定）"
    return (n_pref, n_party), None


def parse_filters(line):
    """
    '北海道の40代' / '選挙部で35歳以上' / '北海道 35-45' などから
    (filter_pref, filter_dept, age_min, age_max) を返す
    """
    text = _normalize(line.strip())
    if not text:
        return None, None, None, None

    filter_pref = None
    age_min = age_max = None

    # 都道府県名
    for pref in PREFECTURES:
        if pref in text:
            filter_pref = pref
            text = text.replace(pref, '')
            break

    # 「40代」→ 40〜49歳
    m = re.search(r'(\d+)代', text)
    if m:
        d = int(m.group(1))
        age_min, age_max = d, d + 9
        text = text[:m.start()] + text[m.end():]

    # 「35〜45歳」「35-45」
    if age_min is None:
        m = re.search(r'(\d+)\s*[〜~\-ー]\s*(\d+)\s*歳?', text)
        if m:
            age_min, age_max = int(m.group(1)), int(m.group(2))
            text = text[:m.start()] + text[m.end():]

    # 「35歳以上」「45歳以下」
    m = re.search(r'(\d+)歳以上', text)
    if m:
        age_min = int(m.group(1)); text = text[:m.start()] + text[m.end():]
    m = re.search(r'(\d+)歳以下', text)
    if m:
        age_max = int(m.group(1)); text = text[:m.start()] + text[m.end():]

    # 残り = 部局キーワード（助詞・空白を除去）
    dept = re.sub(r'[からでのをにはが以上以下　\s]', '', text).strip()
    filter_dept = dept if dept else None

    return filter_pref, filter_dept, age_min, age_max


def parse_confirm(line):
    """
    '確定' 'はい' 'y' 's' '入れ替え' 'キャンセル' など → 'yes'/'swap'/'no'
    完全一致優先・NO > SWAP > YES の順でチェック
    """
    t = _normalize(line.strip().lower())
    YES  = {'y','yes','はい','確定','ok','おk','おけ','決定','いい','もちろん'}
    SWAP = {'s','swap','入れ替え','変更','替え','いれかえ','r','変える','入替'}
    NO   = {'n','no','いいえ','キャンセル','やめ','やめる','cancel','なし','戻る'}
    # 完全一致（NO優先）
    if t in NO:   return 'no'
    if t in SWAP: return 'swap'
    if t in YES:  return 'yes'
    # 部分一致（長い語のみ・NO優先）
    if any(w in t for w in NO   if len(w) >= 3): return 'no'
    if any(w in t for w in SWAP if len(w) >= 3): return 'swap'
    if any(w in t for w in YES  if len(w) >= 2): return 'yes'
    return 'no'  # 不明はキャンセル扱い


def calc_slots(label, ts, te):
    slots, unit, dur = [], timedelta(minutes=SLOT_MIN+BREAK_MIN), timedelta(minutes=SLOT_MIN)
    cur = ts
    while cur + dur <= te:
        slots.append({
            "date":  label,
            "start": cur.strftime("%H:%M"),
            "end":   (cur + dur).strftime("%H:%M"),
        })
        cur += unit
    return slots

# ── 人選 ──────────────────────────────────────────────────
def select_candidates(unmet, meibo, n_pref, n_party,
                      filter_pref=None, filter_dept=None,
                      age_min=None, age_max=None):
    """未面談者から条件に合う人を選出"""
    pref_pool, party_pool = [], []

    for person in unmet:
        name = person["name"]
        info = meibo.get(name, {"org": "", "age": None, "org_type": "党本部"})

        # 年齢フィルタ
        if age_min is not None and info["age"] is not None and info["age"] < age_min: continue
        if age_max is not None and info["age"] is not None and info["age"] > age_max: continue

        if info["org_type"] == "都道府県本部":
            # 都道府県フィルタ
            if filter_pref and filter_pref not in info["org"]: continue
            pref_pool.append(person)
        else:
            # 部局フィルタ
            if filter_dept and filter_dept not in info["org"]: continue
            party_pool.append(person)

    errors = []
    if n_pref > len(pref_pool):
        errors.append(f"都道府県本部の条件を満たす未面談者が{len(pref_pool)}名しかいません（{n_pref}名指定）")
    if n_party > len(party_pool):
        errors.append(f"党本部の条件を満たす未面談者が{len(party_pool)}名しかいません（{n_party}名指定）")
    if errors:
        for e in errors: print(f"  ⚠ {e}")
        return None

    selected = random.sample(pref_pool, n_pref) + random.sample(party_pool, n_party)
    random.shuffle(selected)
    return selected

# ── スケジュール表示 ───────────────────────────────────────
def show_schedule(assignments, meibo):
    date_label = assignments[0][0]["date"] if assignments else ""
    print("\n" + "=" * 50)
    print(f"  【面談スケジュール】{date_label}")
    print("=" * 50)
    for slot, person in assignments:
        name = person["name"]
        info = meibo.get(name, {"org": "", "age": None})
        age_str = f"  {info['age']}歳" if info["age"] else ""
        print(f"  {slot['start']}〜{slot['end']}　{name}（{info['org']}）{age_str}")
    print("=" * 50)

# ── Sheet1に〇を記入 ───────────────────────────────────────
def col_letter(idx):
    r, i = "", idx + 1
    while i > 0:
        i, rem = divmod(i-1, 26)
        r = chr(65+rem) + r
    return r

def find_or_create_date_col(svc, header, date_label):
    """日付ラベルに対応する列インデックスを返す。なければ末尾に追加"""
    nums = re.findall(r'\d+', date_label)
    for i in range(1, len(header)):
        hn = re.findall(r'\d+', header[i])
        if len(nums) >= 2 and len(hn) >= 2 and nums[0]==hn[0] and nums[1]==hn[1]:
            return i

    # 列が存在しない → 末尾に追加
    new_col = len(header)
    cell = f"{SHEET_MEET}!{col_letter(new_col)}1"
    svc.spreadsheets().values().update(
        spreadsheetId=SHEET_ID, range=cell,
        valueInputOption="RAW", body={"values": [[date_label]]}
    ).execute()
    header.append(date_label)
    print(f"  ✓ Sheet1に「{date_label}」列を追加しました（{col_letter(new_col)}列）")
    return new_col

def write_marks(svc, header, assignments, meibo):
    date_label = assignments[0][0]["date"]
    col_idx = find_or_create_date_col(svc, header, date_label)
    updates = []
    for slot, person in assignments:
        sheet_row = person["row_idx"]
        cell = f"{SHEET_MEET}!{col_letter(col_idx)}{sheet_row}"
        updates.append({"range": cell, "values": [["〇"]]})
    if updates:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"valueInputOption": "RAW", "data": updates}
        ).execute()
        print(f"  ✓ {len(updates)}名分の〇を記入しました")

# ── 〇の取り消し ──────────────────────────────────────────
def clear_marks(svc, header, assignments):
    """write_marks直後に呼んで、記入した〇をすべて削除する"""
    date_label = assignments[0][0]["date"]
    nums = re.findall(r'\d+', date_label)
    col_idx = None
    for i in range(1, len(header)):
        hn = re.findall(r'\d+', header[i])
        if len(nums) >= 2 and len(hn) >= 2 and nums[0] == hn[0] and nums[1] == hn[1]:
            col_idx = i
            break
    if col_idx is None:
        print("  ⚠ 列が見つからないため取り消しできません"); return

    ranges = [f"{SHEET_MEET}!{col_letter(col_idx)}{p['row_idx']}"
              for _, p in assignments]
    svc.spreadsheets().values().batchClear(
        spreadsheetId=SHEET_ID, body={"ranges": ranges}
    ).execute()
    print(f"  ✓ {len(assignments)}名分の〇を削除しました")


# ── ③ 人の入れ替え ────────────────────────────────────────
def swap_person(assignments, unmet, meibo):
    """現在の候補からAを外し、未面談者からBを入れる"""

    # 現在の候補一覧
    current_names = [p["name"] for _, p in assignments]
    print("\n  現在の候補:")
    for i, (slot, person) in enumerate(assignments, 1):
        info = meibo.get(person["name"], {"org": ""})
        print(f"    {i}. {slot['start']}〜{slot['end']}　{person['name']}（{info['org']}）")

    # 外す人を検索
    kw_out = input("\n  外す人の氏名（一部でも可）: ").strip()
    if not kw_out:
        print("  キャンセルしました。"); return

    hits_out = [(i, slot, p) for i, (slot, p) in enumerate(assignments) if kw_out in p["name"]]
    if not hits_out:
        print(f"  「{kw_out}」が候補に見つかりません。"); return
    if len(hits_out) > 1:
        print("  複数ヒットしました:")
        for idx, (_, slot, p) in enumerate(hits_out, 1):
            print(f"    {idx}. {p['name']}（{slot['start']}〜{slot['end']}）")
        sel = input("  番号を選択: ").strip()
        if not sel.isdigit() or not (1 <= int(sel) <= len(hits_out)):
            print("  無効な番号です。"); return
        idx_in_assign, slot_out, person_out = hits_out[int(sel)-1]
    else:
        idx_in_assign, slot_out, person_out = hits_out[0]

    # 入れる人を検索（現在の候補以外の未面談者から）
    kw_in = input(f"\n  {person_out['name']} の代わりに入れる人の氏名（一部でも可）: ").strip()
    if not kw_in:
        print("  キャンセルしました。"); return

    # 既に選ばれている人は除外
    excluded = set(current_names)
    hits_in = [p for p in unmet if kw_in in p["name"] and p["name"] not in excluded]
    if not hits_in:
        print(f"  「{kw_in}」に一致する未面談者が見つかりません（既に選ばれている場合も含む）。"); return
    if len(hits_in) > 1:
        print("  複数ヒットしました:")
        for idx, p in enumerate(hits_in, 1):
            info = meibo.get(p["name"], {"org": ""})
            print(f"    {idx}. {p['name']}（{info['org']}）")
        sel = input("  番号を選択: ").strip()
        if not sel.isdigit() or not (1 <= int(sel) <= len(hits_in)):
            print("  無効な番号です。"); return
        person_in = hits_in[int(sel)-1]
    else:
        person_in = hits_in[0]

    # 入れ替え実行
    assignments[idx_in_assign] = (slot_out, person_in)
    info_out = meibo.get(person_out["name"], {"org": ""})
    info_in  = meibo.get(person_in["name"],  {"org": ""})
    print(f"\n  ✓ 入れ替えました:")
    print(f"     {person_out['name']}（{info_out['org']}）→ {person_in['name']}（{info_in['org']}）")


# ── リスケモード ───────────────────────────────────────────
def reschedule_mode(svc):
    print("\n=== リスケ：〇を削除して未面談に戻す ===")
    header, data = load_meeting_sheet(svc)
    rows_all = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"{SHEET_MEET}!A:ZZ"
    ).execute().get("values", [])

    # 〇が入っている人を {name: [(row_idx, col_idx, date_label), ...]} で管理
    marked_map = {}
    for person in data:
        ri = person["row_idx"]
        row = rows_all[ri-1] if ri-1 < len(rows_all) else []
        for ci in range(1, len(header)):
            if ci < len(row) and row[ci].strip() in ("〇","○","o","O"):
                marked_map.setdefault(person["name"], []).append((ri, ci, header[ci]))

    if not marked_map:
        print("〇が入っている方が見つかりません。"); return

    # 氏名で検索
    keyword = input("\nリスケする方の氏名（一部でも可）: ").strip()
    if not keyword:
        print("キャンセルしました。"); return

    # 部分一致で候補を絞る
    hits = [(name, entries) for name, entries in marked_map.items() if keyword in name]
    if not hits:
        print(f"「{keyword}」に一致する方が見つかりません。")
        print("現在〇がある方:")
        for n in sorted(marked_map): print(f"  {n}")
        return

    # 1名だけヒット → そのまま
    # 複数ヒット → 選択
    if len(hits) > 1:
        print(f"\n{len(hits)}名ヒットしました:")
        for idx, (name, entries) in enumerate(hits, 1):
            dlabels = "、".join(e[2] for e in entries)
            print(f"  {idx}. {name}（{dlabels}）")
        sel = input("番号を選択: ").strip()
        if not sel.isdigit() or not (1 <= int(sel) <= len(hits)):
            print("無効な番号です。"); return
        name, entries = hits[int(sel)-1]
    else:
        name, entries = hits[0]

    # 〇の列が複数ある場合はどれを消すか選ぶ
    if len(entries) > 1:
        print(f"\n{name} の〇:")
        for idx, (_, _, dlabel) in enumerate(entries, 1):
            print(f"  {idx}. {dlabel}")
        sel = input("削除する日付の番号: ").strip()
        if not sel.isdigit() or not (1 <= int(sel) <= len(entries)):
            print("無効な番号です。"); return
        ri, ci, dlabel = entries[int(sel)-1]
    else:
        ri, ci, dlabel = entries[0]

    # 確認
    print(f"\n  対象: {name}（{dlabel}）の〇を削除します")
    ans = input("  よろしいですか？（y=実行、それ以外=キャンセル）: ").strip().lower()
    if ans != "y":
        print("キャンセルしました。"); return

    cell = f"{SHEET_MEET}!{col_letter(ci)}{ri}"
    svc.spreadsheets().values().clear(spreadsheetId=SHEET_ID, range=cell).execute()
    print(f"\n✓ {name}（{dlabel}）の〇を削除 → 未面談に戻りました")

# ── メイン ────────────────────────────────────────────────
def main():
    global TEST_MODE

    # ── モード選択メニュー ─────────────────────────────────
    # コマンドライン引数がない場合はインタラクティブ選択
    if "--reschedule" not in sys.argv and "--test" not in sys.argv:
        print("╔" + "═" * 42 + "╗")
        print("║   委員長面談スケジューラー               ║")
        print("╠" + "═" * 42 + "╣")
        print("║  1. 新規スケジュール（人選・〇記入）     ║")
        print("║  2. リスケ（〇を削除して未面談に戻す）   ║")
        print("║  3. テストモード（書き込みなし）         ║")
        print("╚" + "═" * 42 + "╝")
        sel = input("  モードを選択してください（1〜3）: ").strip()
        if sel == "2":
            sys.argv.append("--reschedule")
        elif sel == "3":
            TEST_MODE = True
        # 1 またはそれ以外 → 通常モード

    creds = authenticate()
    svc   = build("sheets", "v4", credentials=creds)

    if "--reschedule" in sys.argv:
        reschedule_mode(svc)
        return

    # データ読み込み
    print("データを読み込んでいます...")
    meibo = load_meibo(svc)
    header, data = load_meeting_sheet(svc)
    unmet = [p for p in data if not p["met"]]

    total       = len(data)
    met_count   = total - len(unmet)
    pref_count  = sum(1 for p in unmet if meibo.get(p["name"],{}).get("org_type")=="都道府県本部")
    party_count = sum(1 for p in unmet if meibo.get(p["name"],{}).get("org_type")=="党本部")
    print(f"\n読み込み完了: 全{total}名（面談済み {met_count}名 ／ 未面談 {len(unmet)}名）")
    print(f"未面談内訳: 都道府県本部 {pref_count}名 ／ 党本部 {party_count}名")

    # 時間枠入力
    print("\n委員長の空き時間を入力してください。")
    print("  例）6月2日 13:00 15:00")
    print("      6/2 13:00-15:00")
    print("      6月2日月曜 午後1時から3時\n")
    while True:
        line = input("空き時間: ").strip()
        result = parse_time_block(line)
        if result:
            date_label, ts, te = result
            print(f"  → {date_label}  {ts.strftime('%H:%M')}〜{te.strftime('%H:%M')}")
            break
        print("  ※ 日付と開始・終了時刻が読み取れませんでした。再入力してください。")

    slots = calc_slots(date_label, ts, te)
    print(f"  → {len(slots)}コマ（各{SLOT_MIN}分）: ", end="")
    print("  ".join(f"{s['start']}〜{s['end']}" for s in slots))

    # 人数指定（柔軟入力）
    print(f"\n何名選びますか？（最大{len(slots)}コマ）")
    print(f"  未面談: 都道府県本部 {pref_count}名 ／ 党本部 {party_count}名")
    print("  例）都道府県2 党本部1　/　2と1　/　都道府県から2名、党本部1名")
    while True:
        line = input("  人数: ").strip()
        counts, err = parse_counts(line, pref_count, party_count, len(slots))
        if counts:
            n_pref, n_party = counts
            print(f"  → 都道府県本部 {n_pref}名 ／ 党本部 {n_party}名")
            break
        print(f"  ※ {err}")
        print(f"  例）都道府県2 党本部1　/　2と1　/　都道府県から2名党1")

    # 絞り込み条件（1行で複合指定・任意）
    print("\n絞り込み条件（不要ならそのままEnter）")
    print("  例）北海道　/　選挙部　/　40代　/　北海道の35〜45歳　/　選挙部で40歳以上")
    while True:
        line = input("  条件: ").strip()
        if not line:
            filter_pref = filter_dept = age_min = age_max = None
            break
        filter_pref, filter_dept, age_min, age_max = parse_filters(line)
        # 読み取り結果を表示して確認
        parts = []
        if filter_pref: parts.append(f"都道府県={filter_pref}")
        if filter_dept: parts.append(f"部局キーワード={filter_dept}")
        if age_min is not None and age_max is not None:
            parts.append(f"年齢={age_min}〜{age_max}歳")
        elif age_min is not None:
            parts.append(f"年齢={age_min}歳以上")
        elif age_max is not None:
            parts.append(f"年齢={age_max}歳以下")
        if parts:
            print(f"  → {' ／ '.join(parts)}")
            break
        else:
            print("  ※ 条件を読み取れませんでした。")
            print("  例）北海道　/　選挙部　/　40代　/　35〜45歳　/　北海道40代")
            retry = input("  再入力しますか？（Enter=条件なしで続ける）: ").strip()
            if not retry:
                filter_pref = filter_dept = age_min = age_max = None
                break

    # 人選
    selected = select_candidates(unmet, meibo, n_pref, n_party,
                                  filter_pref, filter_dept, age_min, age_max)
    if selected is None:
        return

    assignments = list(zip(slots, selected))

    # ── ③ 確認画面（入れ替えループ）─────────────────────────
    while True:
        show_schedule(assignments, meibo)

        if TEST_MODE:
            print("\n【テストモード】スプレッドシートへの書き込みは行いません。")
            print("本番実行: python scheduler.py")
            return

        print("\n" + "┌" + "─" * 48 + "┐")
        print("│  ⚠  この内容で確定しますか？                  │")
        print("│  確定するとSheet1に〇が記入されます。          │")
        print("└" + "─" * 48 + "┘")
        print("  確定 → 「はい」「y」「OK」など")
        print("  入替 → 「入れ替え」「s」など")
        print("  中止 → 「キャンセル」「n」など")
        ans = parse_confirm(input("  選択: "))

        if ans == "yes":
            break
        elif ans == "swap":
            swap_person(assignments, unmet, meibo)
        else:
            print("キャンセルしました。スプレッドシートは変更されていません。")
            return

    # 〇記入
    print("\n--- Sheet1に〇を記入中 ---")
    write_marks(svc, header, assignments, meibo)

    # ── 取り消し猶予 ──────────────────────────────────────
    print()
    print("  ┌" + "─" * 36 + "┐")
    print("  │  〇を記入しました。               │")
    print("  │  取り消す場合は「取り消し」と入力 │")
    print("  │  確定する場合はそのままEnter      │")
    print("  └" + "─" * 36 + "┘")
    undo = input("  > ").strip()
    if undo and any(k in _normalize(undo) for k in ['取り消', 'undo', 'cancel', 'キャンセル', 'やめ']):
        print("\n--- 〇を取り消し中 ---")
        clear_marks(svc, header, assignments)
        print("\n取り消しました。スプレッドシートの〇を削除して未面談に戻しました。")
        return

    # 最終テキスト出力
    print("\n" + "=" * 50)
    print(f"【委員長面談スケジュール】{date_label}")
    print("=" * 50)
    for slot, person in assignments:
        name = person["name"]
        info = meibo.get(name, {"org": ""})
        print(f"{slot['start']}〜{slot['end']}　{name}（{info['org']}）")
    print("=" * 50)
    print("\n完了。リスケ時は: python scheduler.py --reschedule")

if __name__ == "__main__":
    main()
