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
