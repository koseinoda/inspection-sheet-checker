"""
matching_logic.py のテストケース。

実際のExcel/PDFファイルを使わずに、判定ロジックだけを検証する。
カバーしているシナリオ：
    1. 正常データ（ExcelとPDFが一致するケース）
    2. 不一致データ（ExcelとPDFが食い違うケース）
    3. 対象外データ（Excel空欄／PDFに記載が無い項目／要確認項目）
    4. 換気口No未検出データ（PDFのページから換気口Noが読み取れないケース）

実行方法：
    cd このファイルがあるディレクトリの一つ上
    pip install pytest --break-system-packages
    pytest tests/test_matching_logic.py -v
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import matching_logic as ml


def make_word(text, x0, x1, top, bottom, page=1):
    """テスト用に、page.extract_words()が返す形式に合わせた単語dictを作る。"""
    return {"text": text, "x0": x0, "x1": x1, "top": top, "bottom": bottom, "_page_index": page}


def level_words_from(words):
    """単語リストから、レベル文字（A/B/C/S等）だけを抜き出す。"""
    return [w for w in words if w["text"] in ml.LEVEL_SET]


# ============================================================
# 1. 正常データ（ExcelとPDFが一致するケース）
# ============================================================

class TestNormalMatchingCases:
    def test_keyword_with_header_pattern_is_detected_as_present(self):
        """「鉄筋露出　C」のような見出しパターンから、有＋レベルCを検出できる。"""
        text = "変状箇所詳細図 鉄筋露出Ｃ 平面図 断面図"
        words = [
            make_word("鉄筋露出Ｃ", 100, 200, 10, 30),
        ]
        result = ml.evaluate_item("鉄筋露出", text, words, level_words_from(words))
        assert result["pdf_status"] == "有"
        assert result["pdf_shape"] == "C"

    def test_healthy_level_s_is_treated_as_none(self):
        """レベルSは健全＝「無」と判定される。"""
        text = "はく離Ｓ"
        words = [make_word("はく離Ｓ", 100, 200, 10, 30)]
        result = ml.evaluate_item("はく離", text, words, level_words_from(words))
        assert result["pdf_status"] == "無"
        assert result["pdf_shape"] == "S"

    def test_item_not_mentioned_is_none(self):
        """項目名がページに無ければ「無」と判定される。"""
        text = "鉄筋露出Ｃ"
        words = [make_word("鉄筋露出Ｃ", 100, 200, 10, 30)]
        result = ml.evaluate_item("漏水", text, words, level_words_from(words))
        assert result["pdf_status"] == "無"
        assert result["pdf_shape"] == "なし"

    def test_synonym_keyword_taisui_is_detected_as_haisui(self):
        """「排水」はPDF上で「滞水」と表記されることがあり、これも検出できる。"""
        text = "滞水Ｃ"
        words = [make_word("滞水Ｃ", 100, 200, 10, 30)]
        result = ml.evaluate_item("排水", text, words, level_words_from(words))
        assert result["pdf_status"] == "有"
        assert result["pdf_shape"] == "C"

    def test_address_item_matches_excel_exactly(self):
        """地番及び目標物：PDFとExcelの住所が一致すればOK。"""
        text = "地番　台東区三ノ輪1-28(鮨酒肴　杉玉)"
        result = ml.evaluate_item(ml.ADDRESS_ITEM_NAME, text, [], [])
        assert result["address_detected"] is True
        judge_result = ml.judge(
            ml.ADDRESS_ITEM_NAME,
            "台東区三ノ輪1-28(鮨酒肴 杉玉)",  # Excel側はスペースの入れ方が違う
            result["pdf_status"],
        )
        assert judge_result == "OK"

    def test_judge_ok_when_excel_and_pdf_match(self):
        assert ml.judge("ひび割れ", "有", "有") == "OK"
        assert ml.judge("ひび割れ", "無", "無") == "OK"

    def test_edge_stone_with_primary_keyword_is_confidently_detected(self):
        """「縁石」という明確な記載があれば、要確認にせず確定的に判定する。"""
        text = "縁石Ｃ"
        words = [make_word("縁石Ｃ", 100, 200, 10, 30)]
        result = ml.evaluate_item(
            ml.EDGE_STONE_ITEM_NAME, text, words, level_words_from(words)
        )
        assert result["pdf_status"] == "有"
        assert result["pdf_shape"] == "C"


# ============================================================
# 2. 不一致データ（ExcelとPDFが食い違うケース）
# ============================================================

class TestMismatchCases:
    def test_judge_ng_when_excel_and_pdf_differ(self):
        assert ml.judge("鉄筋露出", "無", "有") == "NG"
        assert ml.judge("鉄筋露出", "有", "無") == "NG"

    def test_address_mismatch_is_ng(self):
        result = ml.judge(ml.ADDRESS_ITEM_NAME, "台東区三ノ輪1-28", "台東区根岸5-19")
        assert result == "NG"

    def test_address_ng_when_pdf_could_not_detect_anything(self):
        # PDFから住所を検出できなかった場合、空文字とExcelを比べてNGになる
        result = ml.judge(ml.ADDRESS_ITEM_NAME, "台東区三ノ輪1-28", "")
        assert result == "NG"


# ============================================================
# 3. 対象外データ
# ============================================================

class TestExcludedCases:
    def test_no_pdf_item_is_always_excluded(self):
        """「歩道幅員」「その他」はそもそもPDFに記載が無いため対象外。"""
        text = "何か別の内容"
        result = ml.evaluate_item("歩道幅員", text, [], [])
        assert result["pdf_shape"] == "（PDFに記載なし）"
        assert ml.is_excluded("歩道幅員", "10m", result["pdf_status"]) is True

    def test_empty_excel_value_is_excluded(self):
        assert ml.is_excluded("ひび割れ", "", "有") is True
        assert ml.is_excluded("ひび割れ", None, "有") is True

    def test_edge_stone_fallback_only_is_needs_review_and_excluded(self):
        """
        「縁石」の記載が無く「ひび割れ」表記のみの場合は、自動で有/無を決めず
        「要確認」として対象外に回す。
        """
        text = "ひび割れＣ"  # 「縁石」という文字は無い
        words = [make_word("ひび割れＣ", 100, 200, 10, 30)]
        result = ml.evaluate_item(
            ml.EDGE_STONE_ITEM_NAME, text, words, level_words_from(words)
        )
        assert result["pdf_status"] == "要確認"
        assert ml.is_excluded(ml.EDGE_STONE_ITEM_NAME, "有", result["pdf_status"]) is True

    def test_hibiware_item_is_suppressed_when_edge_stone_claims_it(self):
        """
        「縁石」の記載が無く「ひび割れ」表記のみの場合、それは縁石コンクリート
        破損のことを指している可能性が高いため、内部側の「ひび割れ」項目では
        この記載を使わず「無」のままにする。
        """
        text = "ひび割れＣ"  # 「縁石」という文字は無い
        words = [make_word("ひび割れＣ", 100, 200, 10, 30)]
        result = ml.evaluate_item("ひび割れ", text, words, level_words_from(words))
        assert result["pdf_status"] == "無"

    def test_hibiware_item_detected_normally_when_no_edge_stone_conflict(self):
        """縁石コンクリート破損に関係する記載が無いページでは、通常通り検出する。"""
        text = "ひび割れＣ 縁石"  # 縁石という文字もある＝縁石側で確定的に処理される
        words = [
            make_word("ひび割れＣ", 100, 200, 10, 30),
            make_word("縁石", 300, 340, 10, 30),
        ]
        result = ml.evaluate_item("ひび割れ", text, words, level_words_from(words))
        # 「縁石」が明確にあるので、縁石コンクリート破損側のフォールバックは
        # 発生しない＝「ひび割れ」項目は通常通り検出される
        assert result["pdf_status"] == "有"


# ============================================================
# 4. 換気口No未検出データ
# ============================================================

class TestUndetectedVentNumberCases:
    def test_extract_kanki_no_returns_none_when_no_table_matches(self):
        tables = [
            [["何か無関係な表", "値"], ["別の行", "値2"]],
        ]
        assert ml.extract_kanki_no_from_tables(tables) is None

    def test_extract_kanki_no_returns_none_for_empty_tables(self):
        assert ml.extract_kanki_no_from_tables([]) is None
        assert ml.extract_kanki_no_from_tables(None) is None

    def test_extract_kanki_no_found_normally(self):
        tables = [
            [["線　名", "換気　No."], ["日比谷", "26"]],
        ]
        assert ml.extract_kanki_no_from_tables(tables) == "26"

    def test_build_no_to_pages_collects_undetected_pages(self):
        page_kanki_nos = {
            1: "1",
            2: None,  # 表紙・見出しページなど、換気口Noを検出できないページ
            3: "2",
            4: None,
        }
        pdf_no_to_pages, pdf_page_to_no, no_detected_pages = ml.build_no_to_pages(
            page_kanki_nos
        )
        assert pdf_no_to_pages == {"1": [1], "2": [3]}
        assert pdf_page_to_no == {1: "1", 3: "2"}
        assert no_detected_pages == [2, 4]

    def test_build_no_to_pages_merges_multiple_pages_for_same_no(self):
        """同じ換気口Noが複数ページに分かれている場合、上書きせず両方集める。"""
        page_kanki_nos = {1: "13", 2: "13", 3: "14"}
        pdf_no_to_pages, _, _ = ml.build_no_to_pages(page_kanki_nos)
        assert pdf_no_to_pages["13"] == [1, 2]
        assert pdf_no_to_pages["14"] == [3]

    def test_missing_vent_number_in_pdf_is_distinguishable(self):
        """
        Excel側にある換気口Noが、build_no_to_pagesの結果に無い場合、
        「PDFに対応ページが見つからない」ケースとして検出できる。
        """
        page_kanki_nos = {1: "1", 2: "2"}
        pdf_no_to_pages, _, _ = ml.build_no_to_pages(page_kanki_nos)
        excel_sheet_names = ["1", "2", "3"]  # Excelには"3"もある
        missing_in_pdf = [n for n in excel_sheet_names if n not in pdf_no_to_pages]
        assert missing_in_pdf == ["3"]


# ============================================================
# 文字の二重描画（CAD出力不具合）への対応
# ============================================================

class TestDuplicatedCharacterCorrection:
    def test_detects_duplicated_page_from_known_marker(self):
        """「断面図」が「断断面面図図」になっていれば、二重描画ページと判定する。"""
        page_text = "変状箇所詳細図 鉄筋露出Ｃ 断断 面面 図図 断断"
        assert ml.is_page_text_duplicated(page_text) is True

    def test_normal_page_is_not_flagged_as_duplicated(self):
        page_text = "変状箇所詳細図 鉄筋露出Ｃ 断 面 図 平 面 図"
        assert ml.is_page_text_duplicated(page_text) is False

    def test_dedupe_doubled_value_halves_repeated_characters(self):
        assert ml.dedupe_doubled_value("2299") == "29"
        assert ml.dedupe_doubled_value("11") == "1"  # 単純に「1文字ペア」なら常に半分にする

    def test_dedupe_doubled_value_leaves_non_doubled_value_unchanged(self):
        # ペアが揺れている（全部同じ文字ペアではない）場合は補正しない
        assert ml.dedupe_doubled_value("21") == "21"
        assert ml.dedupe_doubled_value("123") == "123"  # 奇数桁は対象外

    def test_extract_kanki_no_with_text_corrects_only_when_page_is_duplicated(self):
        """
        実際の不具合再現：7ページ目で「29」が「2299」として抽出され、
        ページ全体が二重描画されていることが分かっている場合のみ補正する。
        """
        duplicated_page_text = (
            "変状箇所詳細図 鉄筋露出Ｃ 断面展開図 漏水Ｃ 平面図 漏水跡Ｃ "
            "北千住方 中目黒方（起点側）（終点側）断断 面面 図図 断断"
        )
        tables = [[["線　名", "換　気　No."], ["日比谷", "2299"]]]

        result = ml.extract_kanki_no_with_text(tables, duplicated_page_text)
        assert result == "29"

    def test_extract_kanki_no_with_text_does_not_corrupt_legitimate_double_digit_no(self):
        """
        二重描画が確認できない正常なページでは、「11」のような正当な2桁の
        換気口Noを誤って書き換えてはいけない。
        """
        normal_page_text = "線　名　換気　No. 日比谷 11 断 面 図 平 面 図"
        tables = [[["線　名", "換　気　No."], ["日比谷", "11"]]]

        result = ml.extract_kanki_no_with_text(tables, normal_page_text)
        assert result == "11"

    def test_collapse_adjacent_duplicate_runs_handles_mixed_doubled_and_normal_parts(self):
        """
        文字列の一部だけが二重化されている場合（住所部分は二重・括弧内は正常等）
        にも対応できる。
        """
        s = "番台台東東区区東東上上野野33--1188(常陽銀行)"
        assert ml.collapse_adjacent_duplicate_runs(s) == "番台東区東上野3-18(常陽銀行)"

    def test_extract_chiban_address_corrects_duplicated_label_and_value(self):
        """
        実際の不具合再現：7ページ目で「地番」ラベル自体が「地地番番」に、
        住所も1文字ずつ二重描画されてしまうケースを正しく補正する。
        """
        duplicated_text = (
            "変状箇所詳細図 鉄筋露出Ｃ 断面展開図 漏水Ｃ 平面図 漏水跡Ｃ "
            "北千住方 中目黒方（起点側）（終点側）断断 面面 図図 断断 "
            "線　名　換　気　No. 日比谷 29 "
            "地地番番　台台東東区区東東上上野野33－－1188（常陽銀行）"
        )
        result = ml.extract_chiban_address(duplicated_text)
        assert result == "台東区東上野3-18(常陽銀行)"

    def test_extract_chiban_address_unaffected_on_normal_page(self):
        """二重描画が無い通常のページでは、従来通りの動作のままになる。"""
        normal_text = "地番　台東区三ノ輪1-28(鮨酒肴　杉玉)"
        result = ml.extract_chiban_address(normal_text)
        assert result == "台東区三ノ輪1-28(鮨酒肴杉玉)"


# ============================================================
# find_nearest_level / extract_level_from_header の細かい挙動
# ============================================================

class TestPositionBasedLevelMatching:
    def test_picks_closest_level_word_on_same_page(self):
        item_words = [make_word("鉄筋露出", 100, 200, 100, 120, page=1)]
        level_words = [
            make_word("C", 210, 220, 100, 120, page=1),   # 近い
            make_word("S", 800, 820, 900, 920, page=1),   # 遠い
        ]
        assert ml.find_nearest_level(item_words, level_words) == "C"

    def test_ignores_level_words_on_different_page(self):
        """換気口Noが複数ページにまたがる場合、異なるページの単語同士は比較しない。"""
        item_words = [make_word("鉄筋露出", 100, 200, 100, 120, page=1)]
        level_words = [
            make_word("S", 210, 220, 100, 120, page=2),  # 同じ座標だが別ページ
        ]
        assert ml.find_nearest_level(item_words, level_words) is None

    def test_returns_none_when_too_far(self):
        item_words = [make_word("鉄筋露出", 100, 200, 100, 120, page=1)]
        level_words = [make_word("C", 5000, 5010, 5000, 5020, page=1)]
        assert ml.find_nearest_level(item_words, level_words, max_distance=220) is None

    def test_extract_level_from_header_finds_level_immediately_after_keyword(self):
        assert ml.extract_level_from_header("鉄筋露出　Ｃ", ["鉄筋露出"]) == "C"
        assert ml.extract_level_from_header("ひび割れ S", ["ひび割れ"]) == "S"

    def test_extract_level_from_header_returns_none_when_not_adjacent(self):
        text = "鉄筋露出　無関係な文字列がここに続く　C"
        # "C"がキーワードの直後5文字以内に無いのでNoneになる
        assert ml.extract_level_from_header(text, ["鉄筋露出"]) is None


# ============================================================
# extract_chiban_address / normalize_for_compare の細かい挙動
# ============================================================

class TestAddressExtraction:
    def test_extracts_address_after_label(self):
        text = "地番　台東区三ノ輪1-28(鮨酒肴　杉玉)"
        assert ml.extract_chiban_address(text) == "台東区三ノ輪1-28(鮨酒肴杉玉)"

    def test_returns_none_when_label_not_present(self):
        assert ml.extract_chiban_address("関係ない内容のテキスト") is None

    def test_normalize_for_compare_ignores_fullwidth_and_spaces(self):
        a = ml.normalize_for_compare("台東区三ノ輪1－28　(鮨酒肴 杉玉)")
        b = ml.normalize_for_compare("台東区三ノ輪1-28(鮨酒肴杉玉)")
        assert a == b


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-v"]))
