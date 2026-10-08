"""破損入力が公開例外 (JtdError 系) に収まること。"""
import pytest

from jtdlib import Document, JtdError, NotJtdError, CompressedDocumentError
from jtdlib.container import OLE_MAGIC, JSC_MAGIC, decompress_jsc


def test_not_ole():
    with pytest.raises(NotJtdError):
        Document(b"hello")


def test_truncated_ole_header():
    with pytest.raises(JtdError):
        Document(OLE_MAGIC + b"\x00" * 100)


def test_short_jsc_header():
    with pytest.raises(CompressedDocumentError):
        decompress_jsc(JSC_MAGIC + b"\x00" * 4 + b"-lh5-")


def test_truncated_jsc_payload():
    pytest.importorskip("lhafile")
    blob = JSC_MAGIC + b"\x00" * 12 + b"\x20\x00-lh5-" + b"\x00\x00\xff\xff\x00\x00" + b"\x00" * 40
    with pytest.raises(CompressedDocumentError):
        decompress_jsc(blob)


def test_bad_lha_header():
    pytest.importorskip("lhafile")
    # 長さは足りるがヘッダのチェックサム・内容が不正
    hdr = b"\x16\x00-lh5-" + b"\x05\x00\x00\x00" + b"\x05\x00\x00\x00" + b"\x00" * 16
    blob = JSC_MAGIC + b"\x00" * 12 + hdr + b"\x00" * 16
    with pytest.raises(CompressedDocumentError):
        decompress_jsc(blob)
