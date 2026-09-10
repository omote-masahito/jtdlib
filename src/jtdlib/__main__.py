"""CLI: python -m jtdlib FILE... / jtdlib FILE...

  既定       本文テキスト (一太郎のテキスト保存と同じ順序) を標準出力へ
  --ruby     ルビを「親文字(ルビ)」で出す
  --all      レイアウト枠と脚注も出す
  --json     構造を JSON で出す
  --dump     解析用の詳細表示 (セル座標、インデント、改ページ)
"""
from __future__ import annotations

import argparse
import json
import sys

from . import __version__
from .document import Document
from .model import Paragraph


def _to_json(doc: Document) -> dict:
    def para(p: Paragraph) -> dict:
        d: dict = {"text": p.text}
        if any(r.ruby for r in p.runs):
            d["runs"] = [{"text": r.text, **({"ruby": r.ruby} if r.ruby else {})} for r in p.runs if not r.hidden]
        if p.indent_left is not None:
            d["indent"] = [p.indent_left, p.indent_first]
        if p.page_break_before:
            d["page_break_before"] = True
        return d

    def blocks(bs) -> list:
        out = []
        for b in bs:
            if isinstance(b, Paragraph):
                out.append({"type": "paragraph", **para(b)})
            else:
                out.append({"type": "table", "width": b.width, "rows": [
                    [{"x0": c.x0, "x1": c.x1, "lines": [c.line_start, c.line_end],
                      "paragraphs": [para(p) for p in c.paragraphs]} for c in r.cells]
                    for r in b.rows]})
        return out

    return {
        "body": blocks(doc.blocks),
        "layout_boxes": [blocks(b.blocks) for b in doc.layout_boxes],
        "footnotes": [{"marker": f.marker, "page": f.page, "number_in_page": f.number_in_page,
                       "paragraphs": [para(p) for p in f.paragraphs]} for f in doc.footnotes],
    }


def _dump(doc: Document) -> None:
    for k, b in enumerate(doc.layout_boxes):
        print(f"  LayoutBox[{k}]:", repr(b.text[:80]))
    for k, t in enumerate(doc.footnote_texts):
        print(f"  Footnote[{k}]:", repr(t[:80]))
    for b in doc.blocks:
        if isinstance(b, Paragraph):
            extra = ""
            if b.indent_left is not None:
                extra += f" [indent {b.indent_left}/{b.indent_first}]"
            if b.page_break_before:
                extra += " [page-break]"
            print("  P:", repr(b.text_with_ruby()) + extra)
        else:
            print(f"  T: width={b.width} lines={b.n_lines} cells={len(b.cells)}")
            for r in b.raw_rows:
                print("     ", [(c.x0, c.x1, c.text) for c in r.cells], "(border)" if r.is_border else "")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="jtdlib", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--ruby", action="store_true")
    ap.add_argument("--all", action="store_true", help="レイアウト枠と脚注も出力")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--dump", action="store_true")
    ap.add_argument("--version", action="version", version=__version__)
    a = ap.parse_args(argv)
    rc = 0
    for f in a.files:
        try:
            doc = Document(f)
        except Exception as e:  # noqa: BLE001
            print(f"{f}: {e}", file=sys.stderr)
            rc = 1
            continue
        if len(a.files) > 1 and not a.json:
            print(f"== {f}")
        if a.json:
            print(json.dumps(_to_json(doc), ensure_ascii=False, indent=1))
        elif a.dump:
            _dump(doc)
        else:
            print("\n".join(doc.iter_strings(ruby=a.ruby)))
            if a.all:
                for k, b in enumerate(doc.layout_boxes):
                    print(f"\n--- LayoutBox {k} ---")
                    print("\n".join(b.iter_strings(ruby=a.ruby)))
                if doc.footnotes:
                    print("\n--- Footnotes ---")
                    for fn in doc.footnotes:
                        print(f"{fn.marker} {fn.text}".strip())
        doc.close()
    return rc


if __name__ == "__main__":
    sys.exit(main())
