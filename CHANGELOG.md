# Changelog

## 0.1.0 — 2026-09-11

初版。

- `.jtd` / `.jtdc` (JustCompressedDocument) の読み取り
- 本文、レイアウト枠 (`LayoutBoxText`)、脚注 (`Footnote` + `FootnoteLink`)
- 表: 物理行 → 論理セルの再構成、右罫線の空断片の除去、`Table.rows / columns / cell(r, c)`
- ルビ、タブ、改ページ、段落インデント、脚注番号などの自動生成テキスト
- python-docx 風の読み取り専用 API と CLI (`jtdlib FILE --ruby --all --json --dump`)
- 正解コーパスとの比較テスト (`JTDLIB_CORPUS`)
- RFC 0010 草稿 (`docs/`)

## 0.1.1 — 2026-10-09

外部レビューで指摘されたバグの修正と、生成ツールの追加。

修正

- サロゲートペア (U+10000 以上の文字、𠮷・絵文字など) が 2 文字に分かれ `UnicodeEncodeError` になっていた
- 300 語を超えるインライン文字列が黙って切れ、非表示インラインの末尾が本文に混入していた。終端は領域内で探し、終端が無い場合は `Run.truncated` で印を付ける
- 破損ファイルで `struct.error` / `BadLhafile` がそのまま出ていた。`NotJtdError` / `CompressedDocumentError` に統一
- README の手順どおりの `pytest` が `ModuleNotFoundError` で止まっていた
- 列数の多い表で罫線の対応付けが列数の 4 乗で遅くなっていた (80 列の 1 行: 約 3 秒 → 約 40 ms)
- README: `doc.text` のセル並び順が一太郎のテキスト保存と一致しない場合があることを明記

追加

- `tools/jtdwriter.py`: python-docx 風の API で組み立てた文書を一太郎マクロに変換し、COM 経由で `.jtd` を保存する (Windows + 一太郎が必要)
- `tests/data/サロゲート.jtd`: 一太郎2023 で書き出したサロゲートペア入りの正解データ
