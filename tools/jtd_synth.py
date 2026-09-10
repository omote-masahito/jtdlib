"""jtd_synth.py — 狙った構造だけを持つ最小 .jtd を量産する (taro.py 経由)

  py -3.13-32 jtd_synth.py sources/

レシピはマクロ文字列 (一太郎マクロ言語)。1 要素ずつ変えた文書を並べ、
/DocumentText のレコード差分から未解読フィールドを特定するための素材にする。
各レシピは新規文書に対して実行され、sources/<name>.jtd に形式 0 で保存される。
GetString(2) の結果 (期待テキスト) を sources/<name>.expected.txt に併せて書く。
"""
from __future__ import annotations

import sys
from pathlib import Path

from taro import Taro, FMT_NORMAL, macro_str

RECIPES: dict[str, str] = {
    "s01_plain": '''Insert("これは一段落目です。", 1)
Insert("これは二段落目です。", 1)
Insert("", 1)''',

    "s02_ruby": '''Insert("銀河鉄道の夜を読む。", 1)
SelectString("銀河", 3, 0)
Ruby(1, "ぎんが")
JumpEnd()
SelectString("鉄道", 3, 0)
Ruby(1, "てつどう")
JumpEnd()''',

    "s03_indent": '''Insert("インデントなしの段落。", 1)
Insert("行頭2カラムの段落。", 1)
CursorUp()
Indent(1, 2, 0, 0, 0, 101)
JumpEnd()
Insert("先頭行だけ1カラム下げた段落。", 1)
CursorUp()
Indent(1, 0, 0, 1, 0, 101)
JumpEnd()''',

    "s04_table_2x2": '''Insert("表の前の段落。", 1)
@@TABLE 見出しA,見出しB;a1,b1@@
Insert("表の後の段落。", 1)''',

    "s05_table_3x3": '''@@TABLE h1,h2,h3;r1c1,r1c2,r1c3;r2c1,r2c2,r2c3@@''',

    "s06_table_cell_newline": '''@@TABLE 項目,内容;x,一行目@@
JumpSearchString("一行目", 3, 0)
CursorRight(3)
InsertReturn()
Insert("二行目")
JumpEnd()''',

    "s07_pagebreak": '''Insert("1ページ目。", 1)
InsertPage()
Insert("2ページ目。", 1)''',

    # Footnote は脚注文字を挿入するだけ。脚注エリアへは JumpFootnote(1) で移動して本文を入れ、
    # JumpFootnote(2) で戻る (方向の意味はヘルプ JumpFootnote 参照。失敗したらこのレシピだけ落ちる)。
    "s08_footnote": '''Insert("脚注つきの文。", 1)
JumpSearchString("。", 3, 0)
Footnote("*%1")
JumpFootnote(1)
Insert("これは脚注本文です。")
JumpFootnote(2)
JumpEnd()''',
}


def expand(src: str, t: Taro) -> None:
    """@@TABLE r1c1,r1c2;r2c1,r2c2@@ を insert_table 呼び出しに展開しながら実行する。"""
    for chunk in src.split("@@"):
        if not chunk.strip():
            continue
        if chunk.startswith("TABLE "):
            rows = [r.split(",") for r in chunk[6:].strip().split(";")]
            t.insert_table(rows)
        else:
            t.run(chunk.strip("\n"))


def main(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    with Taro(visible=True) as t:
        print("environment:", t.version())
        for name, src in RECIPES.items():
            print("==", name)
            t.add()
            try:
                expand(src, t)
                p = t.save_document(out_dir / (name + ".x"), FMT_NORMAL)
                lines = t.get_strings(2)
                (out_dir / (name + ".expected.txt")).write_text("\n".join(lines[1:]) + "\n", encoding="utf-8")
                print(f"   -> {p.name if p else 'NOT SAVED'}  {lines[0]}")
            except Exception as e:
                print("   FAILED:", e)
            finally:
                t.close()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(Path(sys.argv[1]))
