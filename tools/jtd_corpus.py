"""jtd_corpus.py — sources/*.jtd を一太郎で開き、同一文書を複数形式で書き出す (taro.py 経由)

  py -3.13-32 jtd_corpus.py sources/ corpus/

出力 corpus/<stem>/:
  <stem>.jtd        形式 0   一太郎2023 が書き直したもの
  <stem>.jtdc       形式 5   圧縮 (JSCompDocument / lh5)
  <stem>.txt        形式 10  UTF-8 テキスト
  <stem>_cells.txt  形式 10  UTF-8, 罫線内文字列=セル単位, 改行置換 CELL_NL
  <stem>.odt        形式 22  OpenDocument (表・ルビの構造化正解)
  <stem>.rtf        形式 15  RTF (形式 20 の .doc も中身は RTF なので 15 で代表)
  <stem>.pdf        形式 23
  <stem>.strings.json   GetString(2) の配列 (セル単位の本文、ルビなし)
  manifest.json     各出力の sha256/サイズ、一太郎とOSのバージョン

原本 (sources/) は Documents.Open で読むだけで書き換えない。
"""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

from taro import Taro, FMT_NORMAL, FMT_NORMAL_COMPRESSED, FMT_TEXT, FMT_ODF, FMT_RTF, FMT_PDF

CELL_NL = "\uE000"

VARIANTS = (  # (suffix hint, fmt, codeset, cell_mode)
    ("", FMT_NORMAL, "ShiftJIS", 0),
    ("", FMT_NORMAL_COMPRESSED, "ShiftJIS", 0),
    ("", FMT_TEXT, "UTF8", 0),
    ("_cells", FMT_TEXT, "UTF8", 1),
    ("", FMT_ODF, "ShiftJIS", 0),
    ("", FMT_RTF, "ShiftJIS", 0),
    ("", FMT_PDF, "ShiftJIS", 0),
)


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def build(src_dir: Path, out_dir: Path) -> None:
    sources = sorted(list(src_dir.glob("*.jtd")) + list(src_dir.glob("*.jtdc")))
    if not sources:
        sys.exit(f"{src_dir} に .jtd/.jtdc が無い")
    manifest: dict = {}
    with Taro(visible=True) as t:
        manifest["_environment"] = t.version()
        for src in sources:
            dst = out_dir / src.stem
            dst.mkdir(parents=True, exist_ok=True)
            entry: dict = {"source": {"path": str(src), "sha256": sha256(src)}}
            print("==", src.name)
            t.open(src)
            try:
                for suffix, fmt, codeset, cell_mode in VARIANTS:
                    try:
                        p = t.save_document(dst / f"{src.stem}{suffix}.x", fmt, codeset, cell_mode,
                                            CELL_NL if cell_mode else "")
                    except Exception as e:
                        entry[f"fmt{fmt}{suffix}"] = {"error": str(e).splitlines()[0]}
                        print(f"   fmt{fmt:<3}{suffix:<7} ERROR {str(e).splitlines()[0]}")
                        continue
                    if p is None:
                        entry[f"fmt{fmt}{suffix}"] = {"error": "not written"}
                        print(f"   fmt{fmt:<3}{suffix:<7} not written")
                        continue
                    entry[p.name] = {"fmt": fmt, "size": p.stat().st_size, "sha256": sha256(p)}
                    print(f"   fmt{fmt:<3}{suffix:<7} -> {p.name:<28}{p.stat().st_size:>8} B")
                try:
                    lines = t.get_strings(2)
                    (dst / f"{src.stem}.strings.json").write_text(
                        json.dumps(lines[1:], ensure_ascii=False, indent=1), encoding="utf-8")
                    entry["strings"] = {"count": len(lines) - 1}
                except Exception as e:
                    entry["strings"] = {"error": str(e).splitlines()[0]}
            finally:
                t.close()
            manifest[src.stem] = entry
    (out_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print("manifest ->", out_dir / "manifest.json")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    build(Path(sys.argv[1]), Path(sys.argv[2]))
