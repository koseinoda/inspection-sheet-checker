"""
変状図・点検表 整合性チェックシステムの判定ロジック本体。

このモジュールはStreamlit（UI）やpdfplumber固有のオブジェクトに依存しない
「純粋関数」として実装している。これにより：

- pytestで実際のExcel/PDFファイルを用意せずにロジックだけをテストできる
- 判定ロジックの変更がUI（app.py）に影響しない
- 将来、別の路線・換気塔にも適用しやすくなる（設定ファイルを変えるだけで対応できる）

設定（点検項目・行番号・PDF検索キーワード等）は items_config.json から読み込む。
設定ファイルが見つからない場合は、このファイル内のDEFAULT_CONFIGにフォールバックする。
"""

import json
import os
import re
import unicodedata

# ============================================================
# 設定の読み込み
# ============================================================

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG_PATH = os.path.join(_THIS_DIR, "items_config.json")

# items_config.json が見つからない場合に使うフォールバック設定。
# 設定ファイルの内容と同じものを保持している。
DEFAULT_CONFIG = {
    "item_rows": {
        "ひび割れ": 8,
        "はく離": 9,
        "鉄筋露出": 10,
        "付属安全設備": 11,
        "昇降設備": 12,
        "排水": 13,
        "漏水": 14,
        "覆板・止め金具": 16,
        "縁石コンクリート破損": 17,
        "ガードレール": 18,
        "地番及び目標物": 19,
        "歩道幅員": 20,
        "その他": 21,
    },
    "item_pdf_keywords": {
        "ひび割れ": ["ひび割れ"],
        "はく離": ["はく離", "剥離"],
        "鉄筋露出": ["鉄筋露出"],
        "付属安全設備": ["付属安全設備"],
        "昇降設備": ["昇降設備"],
        "排水": ["排水", "滞水"],
        "漏水": ["漏水"],
        "覆板・止め金具": ["覆板", "止め金具"],
        "ガードレール": ["ガードレール"],
    },
    "levels": ["AA", "A1", "A2", "A", "B", "C", "S"],
    "healthy_levels": ["S"],
    "address_item_name": "地番及び目標物",
    "no_pdf_items": ["歩道幅員", "その他"],
    "edge_stone_item_name": "縁石コンクリート破損",
    "edge_stone_primary_keywords": ["縁石"],
    "edge_stone_fallback_keywords": ["ひび割れ"],
}


def load_config(config_path=None):
    """
    設定ファイル（JSON）を読み込む。見つからない・壊れている場合は
    DEFAULT_CONFIGにフォールバックする（アプリが完全に止まらないようにするため）。
    "_comment"のようなアンダースコア始まりのキーはコメント用なので無視する。
    """
    path = config_path or DEFAULT_CONFIG_PATH
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except Exception:
        raw = DEFAULT_CONFIG

    config = {k: v for k, v in raw.items() if not k.startswith("_")}

    # 必須キーが欠けていた場合はデフォルト値で補う
    for key, default_value in DEFAULT_CONFIG.items():
        config.setdefault(key, default_value)

    return config


def build_constants(config):
    """
    設定の辞書から、判定ロジックで使う定数群を組み立てる。
    """
    item_rows = config["item_rows"]
    address_item_name = config["address_item_name"]

    return {
        "ITEM_ROWS": item_rows,
        "ITEM_PDF_KEYWORDS": config["item_pdf_keywords"],
        "LEVELS": config["levels"],
        "LEVEL_SET": set(config["levels"]),
        "HEALTHY_LEVELS": set(config["healthy_levels"]),
        "ADDRESS_ITEM_NAME": address_item_name,
        "NO_PDF_ITEMS": set(config["no_pdf_items"]),
        "EDGE_STONE_ITEM_NAME": config["edge_stone_item_name"],
        "EDGE_STONE_PRIMARY_KEYWORDS": config["edge_stone_primary_keywords"],
        "EDGE_STONE_FALLBACK_KEYWORDS": config["edge_stone_fallback_keywords"],
        "ITEMS": [name for name in item_rows.keys() if name != address_item_name],
    }


# モジュール読み込み時にデフォルトの設定を読み込んでおく（app.py側で別の設定を
# 使いたい場合は reload_config() を呼び直せる）
CONFIG = load_config()
CONSTANTS = build_constants(CONFIG)

ITEM_ROWS = CONSTANTS["ITEM_ROWS"]
ITEM_PDF_KEYWORDS = CONSTANTS["ITEM_PDF_KEYWORDS"]
LEVELS = CONSTANTS["LEVELS"]
LEVEL_SET = CONSTANTS["LEVEL_SET"]
HEALTHY_LEVELS = CONSTANTS["HEALTHY_LEVELS"]
ADDRESS_ITEM_NAME = CONSTANTS["ADDRESS_ITEM_NAME"]
NO_PDF_ITEMS = CONSTANTS["NO_PDF_ITEMS"]
EDGE_STONE_ITEM_NAME = CONSTANTS["EDGE_STONE_ITEM_NAME"]
EDGE_STONE_PRIMARY_KEYWORDS = CONSTANTS["EDGE_STONE_PRIMARY_KEYWORDS"]
EDGE_STONE_FALLBACK_KEYWORDS = CONSTANTS["EDGE_STONE_FALLBACK_KEYWORDS"]
ITEMS = CONSTANTS["ITEMS"]


def reload_config(config_path=None):
    """
    別の設定ファイルを読み込んで、モジュールレベルの定数を入れ替える。
    （例：別路線・別換気塔用の設定ファイルに切り替える場合に使う）
    """
    global CONFIG, CONSTANTS
    global ITEM_ROWS, ITEM_PDF_KEYWORDS, LEVELS, LEVEL_SET, HEALTHY_LEVELS
    global ADDRESS_ITEM_NAME, NO_PDF_ITEMS, EDGE_STONE_ITEM_NAME
    global EDGE_STONE_PRIMARY_KEYWORDS, EDGE_STONE_FALLBACK_KEYWORDS, ITEMS

    CONFIG = load_config(config_path)
    CONSTANTS = build_constants(CONFIG)

    ITEM_ROWS = CONSTANTS["ITEM_ROWS"]
    ITEM_PDF_KEYWORDS = CONSTANTS["ITEM_PDF_KEYWORDS"]
    LEVELS = CONSTANTS["LEVELS"]
    LEVEL_SET = CONSTANTS["LEVEL_SET"]
    HEALTHY_LEVELS = CONSTANTS["HEALTHY_LEVELS"]
    ADDRESS_ITEM_NAME = CONSTANTS["ADDRESS_ITEM_NAME"]
    NO_PDF_ITEMS = CONSTANTS["NO_PDF_ITEMS"]
    EDGE_STONE_ITEM_NAME = CONSTANTS["EDGE_STONE_ITEM_NAME"]
    EDGE_STONE_PRIMARY_KEYWORDS = CONSTANTS["EDGE_STONE_PRIMARY_KEYWORDS"]
    EDGE_STONE_FALLBACK_KEYWORDS = CONSTANTS["EDGE_STONE_FALLBACK_KEYWORDS"]
    ITEMS = CONSTANTS["ITEMS"]

    return CONSTANTS


# ============================================================
# 基本ユーティリティ
# ============================================================

def compact_normalize(text):
    """
    全角半角を統一し、空白・改行を完全に除去した文字列を返す。
    大きな見出し文字では文字同士の間隔が広く、PDFのテキスト抽出時に
    余分なスペースが入ってしまうことがあるため、項目名の有無チェックには
    このcompact版も併用する。
    """
    if text is None:
        return ""
    normalized = unicodedata.normalize("NFKC", str(text))
    return normalized.replace(" ", "").replace("　", "").replace("\n", "")


def normalize_for_compare(value):
    """
    住所など自由記述テキストの全角半角・空白の差を無視して比較するための正規化。
    """
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = re.sub(r"\s+", "", text)
    return text.strip()


def _word_center(word):
    cx = (word["x0"] + word["x1"]) / 2
    cy = (word["top"] + word["bottom"]) / 2
    return cx, cy


# ============================================================
# PDFテキスト・単語からの抽出ロジック
# ============================================================

def extract_kanki_no_from_tables(tables):
    """
    pdfplumberのpage.extract_tables()が返す表データ（リストのリストのリスト）から、
    「線名／換気No．」という見出し行の直下のセルに書かれている実際の換気口Noを
    抽出する。テストではpage.extract_tables()の戻り値と同じ形のリストを直接渡せる。
    """
    for table in tables or []:
        for r_idx, row in enumerate(table):
            for c_idx, cell in enumerate(row):
                if not cell:
                    continue
                normalized = unicodedata.normalize("NFKC", cell)
                compact = normalized.replace(" ", "").replace("\n", "")

                if "換気" in compact and "No" in compact:
                    if r_idx + 1 < len(table) and c_idx < len(table[r_idx + 1]):
                        value_cell = table[r_idx + 1][c_idx]
                        if value_cell:
                            value = unicodedata.normalize("NFKC", value_cell)
                            value = value.replace(" ", "").replace("\n", "").strip()
                            if value:
                                return value
    return None


def is_page_text_duplicated(page_text, markers=None):
    """
    ページ全体のテキストに、文字が1文字ずつ二重に描画されている
    （例：「断面図」が「断断面面図図」になっている）現象が起きているかを
    既知の固定ラベルで検出する。
    CADの出力不具合等で、ページ単位でこの現象が起きることがあり、
    その場合は換気口Noなどの値も同様に二重になってしまう。
    """
    if markers is None:
        markers = ["変状箇所詳細図", "平面図", "断面図", "断面展開図"]

    normalized = unicodedata.normalize("NFKC", page_text or "")
    compact = normalized.replace(" ", "").replace("　", "").replace("\n", "")

    for marker in markers:
        doubled = "".join(ch * 2 for ch in marker)
        if doubled in compact:
            return True
    return False


def dedupe_doubled_value(value):
    """
    「2299」のように、1文字ずつ二重に描画されてしまった値を「29」のように戻す。
    全ての文字が連続するペアで同じになっている場合のみ補正する
    （「11」のような正当な2桁の値を誤って「1」にしてしまわないよう、
    呼び出し側でis_page_text_duplicated()による確認が取れた場合だけ使うこと）。
    """
    if not value or len(value) % 2 != 0:
        return value
    if all(value[i] == value[i + 1] for i in range(0, len(value), 2)):
        return value[0::2]
    return value


def extract_kanki_no_with_text(tables, page_text):
    """
    extract_kanki_no_from_tables()の結果に対し、ページ全体に文字の二重描画が
    確認できる場合のみ、二重化補正（dedupe_doubled_value）を適用する。
    テスト時はpage.extract_text()の戻り値を直接渡せる。
    """
    raw_value = extract_kanki_no_from_tables(tables)
    if raw_value is None:
        return None

    if is_page_text_duplicated(page_text):
        return dedupe_doubled_value(raw_value)

    return raw_value


def extract_kanki_no(page):
    """
    pdfplumberのPageオブジェクトから換気口Noを抽出する（実運用用）。
    実際の抽出ロジックはextract_kanki_no_with_text()に委譲している。
    """
    try:
        tables = page.extract_tables()
    except Exception:
        tables = []

    try:
        page_text = page.extract_text() or ""
    except Exception:
        page_text = ""

    return extract_kanki_no_with_text(tables, page_text)


def extract_level_from_header(text, keywords, levels=None):
    """
    各ページの見出しボックスなどにある「項目名＋レベル文字」のパターン
    （例：「鉄筋露出　C」）から、その項目のレベルを直接読み取る。
    図面上の細かい位置関係に頼るより確実なので、こちらを優先的に使う。
    """
    levels = levels if levels is not None else LEVELS
    compact = compact_normalize(text)

    for keyword in keywords:
        start = 0
        while True:
            idx = compact.find(keyword, start)
            if idx == -1:
                break
            after = compact[idx + len(keyword): idx + len(keyword) + 5]
            for level in levels:
                if after.startswith(level):
                    return level
            start = idx + len(keyword)

    return None


def find_nearest_level(item_words, level_words, max_distance=220):
    """
    項目名の単語（item_words）に最も位置が近いレベル文字（level_words）を探す。
    PDFは表ではなく図面上に項目名とレベル文字が個別に配置されているため、
    ページ全体で最初に見つかったレベルを使うと別項目のレベルが混ざってしまう。
    そのため、座標（位置）が一番近いものだけをその項目のレベルとして採用する。
    （換気口Noが複数ページにまたがる場合があるため、異なるページの単語同士は比較しない）
    """
    best_level = None
    best_dist = float("inf")

    for iw in item_words:
        icx, icy = _word_center(iw)
        iw_page = iw.get("_page_index")
        for lw in level_words:
            if iw_page is not None and lw.get("_page_index") != iw_page:
                continue
            lcx, lcy = _word_center(lw)
            dist = ((icx - lcx) ** 2 + (icy - lcy) ** 2) ** 0.5
            if dist < best_dist:
                best_dist = dist
                best_level = lw["text"]

    if best_level is not None and best_dist <= max_distance:
        return best_level
    return None


def collapse_adjacent_duplicate_runs(s):
    """
    文字列の中で「直後の文字と同じ」になっている箇所だけをペアとみなし、
    1文字にまとめる。dedupe_doubled_value()と違い、文字列全体が均一に
    二重化されていなくても対応できる（二重描画されている区間と、
    されていない区間が同じ文字列内に混在しているケースに対応するため）。
    例：「番台台東東区区東東上上野野3 3 - - 1 1 8 8(常陽銀行)」
        → 「番台東区東上野3-18(常陽銀行)」
    """
    if not s:
        return s

    result = []
    i = 0
    n = len(s)
    while i < n:
        if i + 1 < n and s[i] == s[i + 1]:
            result.append(s[i])
            i += 2
        else:
            result.append(s[i])
            i += 1
    return "".join(result)


def extract_chiban_address(text):
    """
    ページのテキストから「地番」ラベルの後に続く住所・目標物のテキストを抽出する。
    （例：「地番　台東区三ノ輪1-28(鮨酒肴　杉玉)」→「台東区三ノ輪1-28(鮨酒肴　杉玉)」）

    ページ全体に文字の二重描画（is_page_text_duplicated参照）が確認できる場合は、
    ラベル自体も「地地番番」のように二重化されているため、そちらを優先して探し、
    抽出した値にcollapse_adjacent_duplicate_runs()で補正をかける。
    """
    normalized = unicodedata.normalize("NFKC", text)
    duplicated = is_page_text_duplicated(text)

    label = "地地番番" if duplicated else "地番"
    label_len = len(label)

    for line in normalized.splitlines():
        compact = line.replace(" ", "").replace("　", "")
        if compact.startswith(label):
            value = compact[label_len:].strip("：: 　")
            if value:
                return collapse_adjacent_duplicate_runs(value) if duplicated else value

    # 行頭に無い場合に備えて、文字列中のどこかにあるラベルも探す
    idx = normalized.find(label)
    if idx != -1:
        remainder = normalized[idx + label_len:]
        value = remainder.splitlines()[0].strip() if remainder else ""
        value = value.strip("：: 　").replace(" ", "").replace("　", "")
        if value:
            return collapse_adjacent_duplicate_runs(value) if duplicated else value

    return None


# ============================================================
# PDFの複数ページの統合（同じ換気口Noが複数ページに分かれる場合への対応）
# ============================================================

def build_no_to_pages(page_kanki_nos):
    """
    {ページ番号: 検出された換気口No（Noneなら未検出）} の辞書から、
    以下の3つを組み立てる：
      - pdf_no_to_pages: {換気口No: [ページ番号, ...]}（同じNoが複数ページに
        分かれている場合は全て集める。上書きはしない）
      - pdf_page_to_no: {ページ番号: 換気口No}
      - no_detected_pages: 換気口Noを検出できなかったページ番号のリスト

    page_kanki_nosはページ番号順（1始まり）であることを前提とする。
    """
    pdf_no_to_pages = {}
    pdf_page_to_no = {}
    no_detected_pages = []

    for page_index in sorted(page_kanki_nos.keys()):
        kanki_no = page_kanki_nos[page_index]
        if kanki_no is None:
            no_detected_pages.append(page_index)
            continue
        pdf_no_to_pages.setdefault(kanki_no, []).append(page_index)
        pdf_page_to_no[page_index] = kanki_no

    return pdf_no_to_pages, pdf_page_to_no, no_detected_pages


# ============================================================
# 項目ごとの判定（このモジュールの中心となる関数）
# ============================================================

def evaluate_item(item, text, words, level_words):
    """
    1つの点検項目について、PDFのテキスト・単語情報から
    {"pdf_status": ..., "pdf_shape": ..., "address_detected": ...} を判定する。

    - item: 点検項目名（ITEM_ROWSのキーのいずれか）
    - text: その換気口Noに対応するPDFページ（複数ページの場合は連結済み）のテキスト
    - words: page.extract_words()相当の単語リスト
             （各要素は {"text","x0","x1","top","bottom","_page_index"} を持つdict）
    - level_words: wordsのうちレベル文字（A/B/C/S等）に該当するものだけのリスト

    戻り値のaddress_detectedは「地番及び目標物」項目のときだけ意味を持つ
    （Trueなら住所を検出できた、Falseなら検出できなかった）。それ以外の項目では
    Noneを返す。
    """
    compact_text = compact_normalize(text)

    # ---- 地番及び目標物：レベル判定ではなく住所テキストの抽出 ----
    if item == ADDRESS_ITEM_NAME:
        extracted_address = extract_chiban_address(text)
        return {
            "pdf_status": extracted_address if extracted_address else "",
            "pdf_shape": "",
            "address_detected": bool(extracted_address),
        }

    # ---- そもそもPDFに記載が無い項目 ----
    if item in NO_PDF_ITEMS:
        return {
            "pdf_status": "",
            "pdf_shape": "（PDFに記載なし）",
            "address_detected": None,
        }

    # ---- 「ひび割れ」：縁石コンクリート破損のフォールバックに奪われていないか確認 ----
    if item == "ひび割れ":
        edge_primary_found = (
            any(k in text for k in EDGE_STONE_PRIMARY_KEYWORDS)
            or any(k in compact_text for k in EDGE_STONE_PRIMARY_KEYWORDS)
            or any(
                any(k in w.get("text", "") for k in EDGE_STONE_PRIMARY_KEYWORDS)
                for w in words
            )
        )
        edge_fallback_found = (
            any(k in text for k in EDGE_STONE_FALLBACK_KEYWORDS)
            or any(k in compact_text for k in EDGE_STONE_FALLBACK_KEYWORDS)
        )

        if edge_fallback_found and not edge_primary_found:
            return {
                "pdf_status": "無",
                "pdf_shape": "なし（「ひび割れ」表記は縁石コンクリート破損として処理）",
                "address_detected": None,
            }

    # ---- 縁石コンクリート破損：専用ロジック ----
    if item == EDGE_STONE_ITEM_NAME:
        primary_words = [
            w for w in words
            if any(k in w.get("text", "") for k in EDGE_STONE_PRIMARY_KEYWORDS)
        ]
        primary_found = (
            bool(primary_words)
            or any(k in text for k in EDGE_STONE_PRIMARY_KEYWORDS)
            or any(k in compact_text for k in EDGE_STONE_PRIMARY_KEYWORDS)
        )

        if primary_found:
            found_level = extract_level_from_header(text, EDGE_STONE_PRIMARY_KEYWORDS)
            if not found_level and primary_words:
                found_level = find_nearest_level(primary_words, level_words)
            found_level = found_level or ""

            pdf_status = "無" if found_level in HEALTHY_LEVELS else "有"

            return {
                "pdf_status": pdf_status,
                "pdf_shape": found_level if found_level else "判定不明",
                "address_detected": None,
            }

        fallback_found = (
            any(k in text for k in EDGE_STONE_FALLBACK_KEYWORDS)
            or any(k in compact_text for k in EDGE_STONE_FALLBACK_KEYWORDS)
        )
        if fallback_found:
            # 「縁石」という明確な記載が無く、「ひび割れ」項目と共有する
            # 言葉でしか検出できないため、自動判定はせず要確認とする
            return {
                "pdf_status": "要確認",
                "pdf_shape": "（「ひび割れ」表記のみで縁石の記載なし。目視確認が必要）",
                "address_detected": None,
            }

        return {"pdf_status": "無", "pdf_shape": "なし", "address_detected": None}

    # ---- それ以外の通常項目 ----
    keywords = ITEM_PDF_KEYWORDS.get(item, [item])

    item_words = [
        w for w in words
        if any(keyword in w.get("text", "") for keyword in keywords)
    ]
    item_found = (
        bool(item_words)
        or any(keyword in text for keyword in keywords)
        or any(keyword in compact_text for keyword in keywords)
    )

    if not item_found:
        return {"pdf_status": "無", "pdf_shape": "なし", "address_detected": None}

    # 1. まず「項目名＋レベル文字」の見出しパターンを探す（最も確実）
    found_level = extract_level_from_header(text, keywords)

    # 2. 見つからない場合のみ、座標が一番近いレベル文字を使う
    if not found_level and item_words:
        found_level = find_nearest_level(item_words, level_words)

    found_level = found_level or ""

    if found_level in HEALTHY_LEVELS:
        # S（健全）の場合は「無」と判定する
        pdf_status = "無"
    else:
        # found_levelが空欄（判定不明）の場合も含め、項目名が見つかった以上は
        # ひとまず「有」として扱う（要確認）。A/AA/A1/A2/B/Cは度合いの違いのみ。
        pdf_status = "有"

    return {
        "pdf_status": pdf_status,
        "pdf_shape": found_level if found_level else "判定不明",
        "address_detected": None,
    }


# ============================================================
# OK/NG判定・対象外判定
# ============================================================

def judge(item, excel_value, pdf_value):
    """
    1件（換気口No × 点検項目）について、ExcelとPDFの記載が一致しているかを判定する。
    戻り値は "OK" または "NG"。
    """
    if item == ADDRESS_ITEM_NAME:
        excel_addr = normalize_for_compare(excel_value)
        pdf_addr = normalize_for_compare(pdf_value)
        return "OK" if excel_addr and excel_addr == pdf_addr else "NG"

    excel_u = str(excel_value).strip() if excel_value is not None else ""
    pdf_u = str(pdf_value).strip() if pdf_value is not None else ""

    return "OK" if excel_u == pdf_u else "NG"


def is_excluded(item, excel_value, pdf_status):
    """
    その行を「対象外」（比較対象から除外）として扱うべきかを判定する。
    - Excel側の「変状有無」が空欄
    - そもそもPDFに記載が無い項目（歩道幅員・その他 等）
    - 「縁石コンクリート破損」が「要確認」（自動判定できない）状態
    のいずれかに当たる場合はTrue。
    """
    excel_str = str(excel_value).strip() if excel_value is not None else ""
    pdf_str = str(pdf_status).strip() if pdf_status is not None else ""

    if excel_str == "":
        return True
    if item in NO_PDF_ITEMS:
        return True
    if pdf_str == "要確認":
        return True
    return False
