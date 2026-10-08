# jtdlib

一太郎の `.jtd` / `.jtdc` からテキストを取り出すPythonライブラリ。
本ツールは一太郎がインストールされていない環境でも利用可能である。
読み取り専用で、API の形は python-docx に近づけた。

本プログラムは[OpenJTD](https://github.com/KimEJ/OpenJTD) (Apache-2.0, clean-room) の RFC と、
一太郎2023 で生成した合成文書および行政機関がインターネットで公開している一太郎文書を使用した。

## 背景

一太郎文書からテキストを取り出す公開ツールとして、デジタル庁の
[一太郎→Word 変換ツール](https://laws.e-gov.go.jp/tools/analog-inspection-tool/)
（Excel マクロ `ConvertingIchitaro2Word.xlsm`）がある。マクロの動作は次のとおり。

1. `JXW.Application` を COM で起動し、`TaroLibrary.SaveDocument(..., 10)` でテキスト形式に保存する。
   書式はこの時点で失われる。
2. 保存したテキストを 1 行ずつ読み、改ページ（`Chr(12)`）を含む行は捨てる。
3. 罫線文字（`│ ┌ ├ ┼ └ ─ ┤ ┬ ┐`）を含む行を表とみなし、`│` で分割して列ごとに文字列を連結する。
   列数が変わったところで 1 列を 1 段落として書き出し、その際に列内の半角スペースを削除する。
4. `Word.Application` を起動し、新規文書の `Content.Text` に結果を代入して docx で保存する。

従って、一太郎・Word・Excel を入れた Windows 機が必要で、出力は書式のない docx になる。
表は罫線文字からの推測で列ごとにまとめられるため、複数行にわたるセルや結合セルは元のセル単位に
戻らず、英数字の語間などの半角スペースも失われる。

jtdlib は一太郎のファイルを直接解析するので、一太郎も Word も Windows も要らず、Linuxでもテキストを取り出せる。表は罫線文字ではなくファイル内のセルレコードから再構成するので、複数行セルや結合セルもセル単位で取れる。一方で出力は文字列と表構造だけで、docx を作る機能はない。docx が要る場合は python-docx などで組み立てること。

## インストール

```
git clone https://github.com/omote-masahito/jtdlib jtdlib
pip install ./jtdlib              # .jtd
pip install "./jtdlib[jtdc]"      # .jtdc (圧縮形式) も読む。lhafile が入る
```

clone せずに入れるなら `pip install "git+https://github.com/omote-masahito/jtdlib"`、extras 付きなら `pip install "jtdlib[jtdc] @ git+https://github.com/omote-masahito/jtdlib"`。
開発用は `pip install -e "./jtdlib[test]"`。

依存は olefile のみ(`.jtdc` を読む場合は lhafile も必要)。

## 使い方

```python
from jtdlib import Document

doc = Document("a.jtd")            # パス / bytes / バイナリファイルオブジェクト

doc.text                           # 全文。表のセルも含むが、並び順は一太郎のテキスト保存と一致しない場合がある (後述)
for p in doc.paragraphs:           # 本文の段落 (表の外)
    print(p.text, p.indent_left, p.page_break_before)
for t in doc.tables:
    for row in t.rows:
        print([c.text for c in row.cells])
    t.columns                      # x 座標で切った列
    t.cell(0, 1)                   # rows[0].cells[1]
for block in doc.iter_inner_content():   # Paragraph | Table を本文順に
    ...
for run in doc.paragraphs[0].runs:
    run.text, run.ruby             # ルビは Run に付く
doc.paragraphs[0].text_with_ruby() # "親文字(ルビ)" 表記
doc.footnotes                      # Footnote(marker="*1", paragraphs=[...], page=, number_in_page=)
doc.layout_boxes                   # レイアウト枠 (テキストボックス)。本文と同じ API を持つ
```

CLI:

```
jtdlib a.jtd                 # 本文テキスト
jtdlib a.jtd --ruby --all    # ルビ付き、レイアウト枠と脚注も
jtdlib a.jtd --json          # 構造を JSON で
jtdlib a.jtd --dump          # 解析用 (セル座標など)
```

## python-docx との違い

- 一太郎の表は「物理行 + セル断片」でしか格納されない。`Row` は「同じ物理行から始まる論理セルの集まり」という近似で、結合セルは 1 つの `Cell` として 1 回だけ現れる (python-docx のように結合範囲に同じセルが繰り返されない)。`Table.cell(r, c)` も補完しない。
- セル座標 `Cell.x0 / x1` の単位は 1 文字幅の 1/4 (`COORD_UNITS_PER_CHAR`)。`Cell.width_chars` で文字数換算。
- 書式 (フォント・文字スタイル・画像) は持たない。段落インデント (カラム単位) と改ページだけ。
- 表のセル文字列は抽出できるが、`doc.text` / `iter_strings()` でのセルの並び順は一太郎の「テキスト保存 (罫線内文字列=セル単位)」と一致しない場合がある。一太郎のセル列挙順は未解読で、検証した行政文書 9 本のうち順序まで一致したのは 3 本 (残りは空でない文字列の多重集合として一致。RFC 0010 の Known Gaps)。
- 書き込みはできない。

## 対応状況

- 本文 (`DocumentText`)、レイアウト枠 (`LayoutBoxText`)、脚注 (`Footnote` + `FootnoteLink`)
- ルビ、タブ、改ページ、段落インデント、脚注番号などの自動生成テキスト
- 表: 物理行→論理セルの縦結合、右罫線の空断片の除去
- `.jtdc` (JustCompressedDocument = LHA -lh5-)
- 未対応: 図・画像、文字スタイル、`LayoutBox` 座標と本文アンカーの対応、fmt 1 の行スペック詳細

## 開発

```
pip install -e ".[test]"
pytest                          # 同梱の小さな fixture (tests/data) で回る
JTDLIB_CORPUS=dir1:dir2 pytest  # 正解付きコーパスとの比較も回す (Windows は ; 区切り)
```

コーパス比較テスト (`tests/test_corpus.py`) は、一太郎2023 で書き出した正解 (`<stem>_cells.txt`:
テキスト保存・セル単位・ルビ=親(ルビ)・セル内改行=U+E000) と `.jtd` / `.jtdc` の組を
`JTDLIB_CORPUS` のディレクトリから読む。検証には行政機関がインターネットで公開している一太郎文書と`tools/jtd_synth.py` の合成文書を使っているが、文書本体はリポジトリに含めない。

検査項目:
- multiset: 空でない文字列の多重集合が一致 (全文書で必須)
- exact: 並び順まで一致。コーパス直下の `expectations.json` で `exact: true` の文書は必須、
  `false` は strict xfail (セル列挙順を解読して一致するようになったら `true` に上げる)
- `.jtd` と `.jtdc` の出力が同じ、ファイルの sha256 が `manifest.json` と一致

`tools/` は正解データ生成用のスクリプト (Windows + 一太郎 + 32bit Python + pywin32 が必要):
`taro.py` (COM ラッパー)、`jtd_synth.py` (合成コーパス)、`jtd_corpus.py` (7 形式書き出し)、
`verify_corpus.py` (対話用の比較 CLI)。マクロ名の索引、使用方法は一太郎のドキュメントを参照すること。

## 謝辞 (Acknowledgements)

- [OpenJTD](https://github.com/KimEJ/OpenJTD) と作者の KimEJ (kimeojin) 氏。一太郎文書形式の clean-room な解析を
  公開仕様 (RFC) と Rust 実装 `rjtd` として公開しており、jtdlib のコンテナ・`.jtdc` 展開・レコード構造の理解は
  RFC 0001/0003/0005/0009 を出発点にしている。`docs/0010-*.md` はその続きとして作成したものである。
- [olefile](https://github.com/decalage2/olefile) と [lhafile](https://github.com/FrodeSolheim/lhafile)。

## 本コードについて

このライブラリの設計・実装・テスト・ドキュメント、および RFC 0010 の草稿は、Claudeの手を借りて作成した。

## ライセンス

Apache-2.0
