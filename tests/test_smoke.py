"""同梱の 3 文書での動作確認。コーパス比較テスト (verify_corpus 相当) は次段で追加する。"""
import io

import pytest

from jtdlib import Document, NotJtdError, Paragraph, Table
from tests.conftest import data_file


def test_footnote_single():
    doc = Document(data_file("テスト脚注.jtd"))
    assert len(doc.footnotes) == 1
    assert doc.footnotes[0].text.startswith("用語")
    assert doc.paragraphs[0].text == "概要"
    assert not doc.tables


def test_footnote_multi_and_links():
    doc = Document(data_file("テスト脚注2.jtd"))
    assert len(doc.footnotes) == 5
    assert all(fn.page is not None for fn in doc.footnotes)
    assert doc.footnotes[2].text.startswith("RFC 0009 の w10")


def test_tables_present():
    doc = Document(data_file("罫線テスト.jtd"))
    assert len(doc.tables) == 2
    kinds = [type(b) for b in doc.iter_inner_content()]
    assert Table in kinds and Paragraph in kinds
    assert "標準で作成" in doc.text


def test_sources_and_context_manager():
    p = data_file("テスト脚注.jtd")
    raw = p.read_bytes()
    with Document(raw) as d1, Document(io.BytesIO(raw)) as d2, Document(str(p)) as d3:
        assert d1.text == d2.text == d3.text


def test_not_jtd():
    with pytest.raises(NotJtdError):
        Document(b"not an ole file at all")


def test_cli(capsys):
    from jtdlib.__main__ import main
    assert main([str(data_file("テスト脚注.jtd")), "--all"]) == 0
    out = capsys.readouterr().out
    assert "概要" in out and "--- Footnotes ---" in out
    assert main([str(data_file("テスト脚注.jtd")), "--json"]) == 0


def test_keisen_2x2():
    """罫線テスト.jtd: 1 行 30 字の用紙に 2×2 表 (座標単位 = 1/4 文字、全幅 120)。
    右端の (118,120) 空断片は右罫線であってセルではない。"""
    doc = Document(data_file("罫線テスト.jtd"))
    for t in doc.tables:
        assert t.width == 120
        assert [len(r) for r in t.rows] == [2, 2]
        assert [(c.x0, c.x1) for c in t.columns] == [(2, 60), (60, 116)]
        assert t.cell(1, 1).width_chars == 14.0
