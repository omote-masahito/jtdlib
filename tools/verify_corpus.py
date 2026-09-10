"""verify_corpus.py — jtdlib の出力を一太郎の正解データと突き合わせる (対話用 CLI)

  python verify_corpus.py corpus/

同じ比較は pytest でも回る: JTDLIB_CORPUS=corpus pytest tests/test_corpus.py

corpus/<stem>/<stem>.jtd (と .jtdc) を jtdlib で読み、<stem>_cells.txt
(テキスト保存, 罫線内文字列=セル単位, ルビ=親(ルビ), セル内改行=U+E000) と比較する。
  exact : 段落・セルの並び順まで完全一致
  cells : 空でない文字列の多重集合が一致 (一太郎のセル列挙順は未解読なので順序は問わない)
"""
import collections
import sys
from pathlib import Path

from jtdlib import Document  # pip install -e .

CELL = "\ue000"


def check(root: Path) -> int:
    fails = 0
    for d in sorted(p for p in root.iterdir() if p.is_dir()):
        stem = d.name
        cells = d / f"{stem}_cells.txt"
        if not cells.exists():
            continue
        exp = cells.read_text(encoding="utf-8-sig").splitlines()
        for ext in ("jtd", "jtdc"):
            f = d / f"{stem}.{ext}"
            if not f.exists():
                continue
            try:
                doc = Document(f)
            except Exception as e:
                print(f"FAIL {f.name}: {e}")
                fails += 1
                continue
            got = [s.replace("\n", CELL) for s in doc.iter_strings(ruby=True)]
            exact = exp == got
            ee = collections.Counter(x for x in exp if x.strip(CELL))
            ge = collections.Counter(x for x in got if x.strip(CELL))
            missing, extra = ee - ge, ge - ee
            ok = exact or (not missing and not extra)
            tag = "OK  " if ok else "FAIL"
            print(f"{tag} {f.name:<36} exact={str(exact):<5} cells: exp={sum(ee.values()):3d} "
                  f"got={sum(ge.values()):3d} missing={sum(missing.values()):2d} extra={sum(extra.values()):2d}")
            if not ok:
                fails += 1
                for x in list(missing)[:3]:
                    print("      missing:", repr(x[:80]))
                for x in list(extra)[:3]:
                    print("      extra  :", repr(x[:80]))
    return fails


if __name__ == "__main__":
    n = check(Path(sys.argv[1]))
    print("failures:", n)
    sys.exit(1 if n else 0)
