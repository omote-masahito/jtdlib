"""コーパス比較テスト (旧 tools/verify_corpus.py)。JTDLIB_CORPUS が無ければ skip。

コーパスは一太郎2023 で書き出した正解付きの文書集合 (tools/jtd_corpus.py)。開発では、行政機関が
インターネットで公開している一太郎文書と、tools/jtd_synth.py の合成文書を使っている。
文書本体はリポジトリに含めない。

正解は一太郎の「テキスト保存 (罫線内文字列=セル単位, ルビ=親(ルビ), セル内改行=U+E000)」。
  * multiset: 空でない文字列の多重集合が一致する (全文書で必須)
  * exact   : 並び順まで完全一致する。expectations.json で exact=true の文書は必須、
              false の文書は strict xfail (一致するようになったら expectations を上げる)
"""
from __future__ import annotations

import collections
import hashlib
import json

import pytest

from jtdlib import Document
from jtdlib.constants import CELL_NEWLINE as CELL


def _got(case) -> list[str]:
    with Document(case.path) as doc:
        return [s.replace("\n", CELL) for s in doc.iter_strings(ruby=True)]


def _nonempty(xs) -> collections.Counter:
    return collections.Counter(x for x in xs if x.strip(CELL))


def test_cells_multiset(case):
    exp, got = _nonempty(case.expected()), _nonempty(_got(case))
    missing, extra = exp - got, got - exp
    assert not missing and not extra, {
        "missing": [x[:80] for x in list(missing)[:5]],
        "extra": [x[:80] for x in list(extra)[:5]],
    }


def test_exact_order(case, request):
    if case.exact is None:
        pytest.skip("exact-order expectation unknown for external corpus")
    if not case.exact:
        request.applymarker(pytest.mark.xfail(strict=True, reason="cell enumeration order not yet decoded"))
    assert case.expected() == _got(case)


def test_jtd_and_jtdc_agree(stem_dir):
    """圧縮形式は展開後に同じ DocumentText を持つはず。"""
    jtd, jtdc = stem_dir / f"{stem_dir.name}.jtd", stem_dir / f"{stem_dir.name}.jtdc"
    if not jtdc.exists():
        pytest.skip("no .jtdc")
    with Document(jtd) as a, Document(jtdc) as b:
        assert list(a.iter_all_text()) == list(b.iter_all_text())
        assert [len(t.cells) for t in a.tables] == [len(t.cells) for t in b.tables]
        assert len(a.footnotes) == len(b.footnotes)


def test_fixture_integrity(corpus_root):
    """コーパスが一太郎の書き出し時から改変されていないこと (manifest.json の sha256)。"""
    man = json.loads((corpus_root / "manifest.json").read_text(encoding="utf-8"))
    checked = 0
    for stem, entry in man.items():
        if stem.startswith("_"):
            continue
        for name, info in entry.items():
            if name == "source" or not isinstance(info, dict):
                continue
            p = corpus_root / stem / name
            if not p.exists():
                continue
            assert hashlib.sha256(p.read_bytes()).hexdigest() == info["sha256"], name
            checked += 1
    assert checked > 0
