"""jtdwriter のテスト。

compile() のテストは一太郎なしで動く。往復テストは Windows + 一太郎 + jtdlib が
揃っている環境でのみ実行され、それ以外では skip される。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))  # tools/ を import 可能に
import jtdwriter  # noqa: E402
from jtdwriter import Document, _expr  # noqa: E402


# --- マクロ文の生成 (一太郎不要) -------------------------------------------

def test_expr_escapes_quotes_and_control_chars():
    assert _expr('a"b') == '"a""b"'
    assert _expr("a\tb") == '"a" & Char(9) & "b"'
    assert _expr("\n") == "Char(10)"
    assert _expr("") == '""'


def test_plain_paragraph():
    doc = Document()
    doc.add_paragraph("こんにちは")
    assert doc.compile().splitlines() == ['Insert("こんにちは",0)', "InsertReturn(1,1)"]


def test_run_attributes_are_set_then_reset_in_reverse():
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("x", bold=True, size=12, underline=True, color="R255G0B0")
    assert doc.compile().splitlines() == [
        "InputSizePoint(12)",
        "InputBold(1)",
        'InputDecoration(.ｱﾝﾀﾞｰﾗｲﾝ=1,.文字色="R255G0B0")',
        'Insert("x",0)',
        'InputDecoration(.ｱﾝﾀﾞｰﾗｲﾝ=0,.文字色="R0G0B0")',
        "InputBold(0)",
        "InputSizePoint()",
        "InsertReturn(1,1)",
    ]


def test_alignment_needs_range_and_is_cancelled_on_next_paragraph():
    doc = Document()
    doc.add_paragraph("c", alignment="center")
    assert doc.compile().splitlines() == [
        'Insert("c",0)',
        "SelectRangeStart(3)",
        "FormatCenter()",
        "JumpEnd()",
        "InsertReturn(1,1)",
        "SelectRangeStart(3)",
        "CancelParagraphAttribute(1,0,0,0,0,0,0)",
        "JumpEnd()",
    ]


def test_table_uses_paragraph_unit_selection():
    doc = Document()
    doc.add_table(0, 0, [["a", "b"], ["1", "2"]])
    assert doc.compile().splitlines() == [
        'Insert("a" & Char(9) & "b",1)',
        'Insert("1" & Char(9) & "2",0)',
        "SelectRangeStart(3)",
        "CursorUp(1)",
        "ConvertStringToKeisenTable(,2)",
        "JumpEnd()",
        "InsertReturn(1,1)",
    ]


def test_page_break_uses_insert_page():
    doc = Document()
    doc.add_page_break()
    assert doc.compile() == "InsertPage()"


def test_table_rejects_ragged_rows():
    doc = Document()
    doc.add_table(0, 0, [["a", "b"], ["1"]])
    with pytest.raises(ValueError):
        doc.compile()


def test_heading_falls_back_to_bold_and_size():
    doc = Document()
    doc.add_heading("H", 2)
    assert doc.compile().splitlines()[:2] == ["InputSizePoint(14)", "InputBold(1)"]


def test_heading_style_is_resolved_at_compile_time():
    doc = Document()
    doc.add_heading("H", 1)
    doc.heading_styles = {1: "大見出し1"}   # add_heading の後で代入しても効く
    lines = doc.compile().splitlines()
    assert lines[0] == 'CallParagraphStyle("大見出し1")'   # 文書への取り込みは一度だけ
    assert lines[1] == 'SetParagraphStyle(1,"大見出し1")'
    assert lines[-1] == 'SetParagraphStyle(0,"大見出し1")'
    assert lines.count('CallParagraphStyle("大見出し1")') == 1
    assert "InputBold(1)" not in lines


# --- 往復 (Windows + 一太郎 + jtdlib) --------------------------------------

def _taro_available() -> bool:
    if sys.platform != "win32" or jtdwriter.Taro is None:
        return False
    try:
        import win32com.client as w32
        w32.Dispatch("JXW.Application")
        return True
    except Exception:
        return False


@pytest.mark.skipif(not _taro_available(), reason="一太郎 (JXW.Application) が使える Windows 環境でのみ実行")
def test_roundtrip_through_ichitaro(tmp_path: Path):
    jtdlib = pytest.importorskip("jtdlib")
    doc = jtdwriter.demo()
    out = doc.save(tmp_path / "rt.jtd")
    assert out is not None and out.exists()

    rd = jtdlib.Document(str(out))
    texts = [p.text for p in rd.paragraphs]
    assert texts[:4] == [
        "jtdwriter サンプル",
        "python-docx 風の API で組み立てた文書を、一太郎マクロに変換して保存する。",
        "中央揃え",
        '下線と赤字と"引用符"、タブ区切りの表 (100%)。',
    ]
    assert texts[-3:] == ["表の後", "2ページ目、右揃え", "3ページ目、左揃え"]
    assert [p.page_break_before for p in rd.paragraphs[-3:]] == [False, True, True]
    assert len(rd.tables) == 1
    cells = [[c.text for c in r.cells] for r in rd.tables[0].rows]
    assert cells == doc.tables[0].rows
