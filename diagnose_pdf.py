"""
pdfplumberがPDFを正しく読み取れているかを直接確認するための診断スクリプト。
Streamlitを経由せず、エラーが起きた場合はそのまま画面に表示する
（app.py内のtry/exceptで握りつぶされてしまう問題を避けるため）。

使い方（PowerShell）：
    python diagnose_pdf.py "PDFファイルへのパス"

例：
    python diagnose_pdf.py "C:\\Users\\kosei\\OneDrive\\地下構造用フォルダ\\変状図（反映済み）.pdf"
"""

import sys
import unicodedata

import pdfplumber


def main():
    if len(sys.argv) < 2:
        print("使い方: python diagnose_pdf.py \"PDFファイルへのパス\"")
        sys.exit(1)

    pdf_path = sys.argv[1]

    print(f"pdfplumberのバージョン: {pdfplumber.__version__}")
    print(f"対象ファイル: {pdf_path}")
    print("=" * 60)

    with pdfplumber.open(pdf_path) as pdf:
        print(f"総ページ数: {len(pdf.pages)}")
        print("=" * 60)

        for page_index, page in enumerate(pdf.pages, start=1):
            print(f"\n--- {page_index}ページ目 ---")

            # 1. extract_text() の確認
            try:
                text = page.extract_text() or ""
                print(f"extract_text(): 成功（文字数={len(text)}）")
                preview = text.replace("\n", " ")[:80]
                print(f"  先頭80文字: {preview}")
            except Exception as e:
                print(f"extract_text(): ★失敗★ {type(e).__name__}: {e}")

            # 2. extract_tables() の確認（換気口No検出に使用）
            try:
                tables = page.extract_tables()
                print(f"extract_tables(): 成功（表の数={len(tables)}）")

                found_kanki_no = None
                for table in tables:
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
                                        found_kanki_no = (
                                            unicodedata.normalize("NFKC", value_cell)
                                            .replace(" ", "")
                                            .replace("\n", "")
                                            .strip()
                                        )
                if found_kanki_no:
                    print(f"  → 換気口No検出: 「{found_kanki_no}」")
                else:
                    print("  → 換気口Noは検出できませんでした（表は読めているが、該当パターンが無い）")
            except Exception as e:
                print(f"extract_tables(): ★失敗★ {type(e).__name__}: {e}")

            # 3. extract_words() の確認（レベル判定に使用）
            try:
                words = page.extract_words(x_tolerance=15, y_tolerance=5)
                print(f"extract_words(): 成功（単語数={len(words)}）")
            except Exception as e:
                print(f"extract_words(): ★失敗★ {type(e).__name__}: {e}")

    print("\n" + "=" * 60)
    print("診断完了。上記で「★失敗★」と出ている箇所があれば、そのエラー内容を教えてください。")


if __name__ == "__main__":
    main()
