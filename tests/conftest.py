"""fixtures:
  tests/data/  — 手作りの小さな .jtd (脚注、罫線)。リポジトリに同梱。
  環境変数 JTDLIB_CORPUS=dir1<os.pathsep>dir2 — 正解付きコーパス (jtd_corpus.py / jtd_synth.py の出力:
    <stem>/<stem>.jtd, .jtdc, <stem>_cells.txt、任意で manifest.json と expectations.json)。
    リポジトリには含めない。未設定ならコーパス比較テストは skip。
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path

import pytest

HERE = Path(__file__).parent
DATA = HERE / "data"
CELL = "\ue000"          # テキスト保存 (セル単位) のセル内改行


def data_file(name: str) -> Path:
    return DATA / name


@dataclass(frozen=True)
class CorpusCase:
    root: Path
    stem: str
    ext: str                       # "jtd" | "jtdc"
    exact: bool | None             # None = 不明 (外部コーパス)

    @property
    def path(self) -> Path:
        return self.root / self.stem / f"{self.stem}.{self.ext}"

    @property
    def cells_path(self) -> Path:
        return self.root / self.stem / f"{self.stem}_cells.txt"

    def expected(self) -> list[str]:
        return self.cells_path.read_text(encoding="utf-8-sig").splitlines()

    @property
    def id(self) -> str:
        return f"{self.stem}.{self.ext}"


def _corpus_roots() -> list[Path]:
    roots: list[Path] = []
    for p in os.environ.get("JTDLIB_CORPUS", "").split(os.pathsep):
        if p and Path(p).is_dir():
            roots.append(Path(p))
    return roots


def corpus_cases() -> list[CorpusCase]:
    cases: list[CorpusCase] = []
    for root in _corpus_roots():
        exp = {}
        f = root / "expectations.json"
        if f.exists():
            exp = {k: v for k, v in json.loads(f.read_text(encoding="utf-8")).items() if not k.startswith("_")}
        for d in sorted(p for p in root.iterdir() if p.is_dir()):
            if not (d / f"{d.name}_cells.txt").exists():
                continue
            for ext in ("jtd", "jtdc"):
                if (d / f"{d.name}.{ext}").exists():
                    cases.append(CorpusCase(root, d.name, ext, exp.get(d.name, {}).get("exact")))
    return cases


def corpus_roots() -> list[Path]:
    return _corpus_roots()


def stems() -> list[tuple[Path, str]]:
    return sorted({(c.root, c.stem) for c in corpus_cases()})


def pytest_generate_tests(metafunc):
    skip = pytest.mark.skip(reason="set JTDLIB_CORPUS to run corpus comparison tests")
    if "case" in metafunc.fixturenames:
        cases = corpus_cases()
        metafunc.parametrize("case", cases or [pytest.param(None, marks=skip)],
                             ids=[c.id for c in cases] or ["no-corpus"])
    if "stem_dir" in metafunc.fixturenames:
        s = stems()
        metafunc.parametrize("stem_dir", [r / n for r, n in s] or [pytest.param(None, marks=skip)],
                             ids=[n for _, n in s] or ["no-corpus"])
    if "corpus_root" in metafunc.fixturenames:
        roots = [r for r in _corpus_roots() if (r / "manifest.json").exists()]
        metafunc.parametrize("corpus_root", roots or [pytest.param(None, marks=skip)],
                             ids=[r.name for r in roots] or ["no-corpus"])
