"""OLE2 コンテナと .jtdc (JustCompressedDocument) の展開。"""
from __future__ import annotations

import io
import struct

import olefile

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
JSC_MAGIC = b"&\x00JustCompressedDocument"


class JtdError(Exception):
    """jtdlib の基底例外。"""


class NotJtdError(JtdError):
    """OLE2 ファイルではない、または一太郎文書のストリームを持たない。"""


class CompressedDocumentError(JtdError):
    """.jtdc の展開に失敗した。"""


def open_ole(data: bytes) -> olefile.OleFileIO:
    """bytes から OleFileIO を開く。.jtdc なら JSCompDocument を展開して内側の OLE を返す。"""
    if not data.startswith(OLE_MAGIC):
        raise NotJtdError("not an OLE2 compound file")
    ole = olefile.OleFileIO(data)
    if ole.exists("JSCompDocument") and not ole.exists("DocumentText"):
        inner = decompress_jsc(ole.openstream("JSCompDocument").read())
        ole.close()
        if not inner.startswith(OLE_MAGIC):
            raise CompressedDocumentError("decompressed JSCompDocument is not an OLE2 file")
        return olefile.OleFileIO(inner)
    return ole


def decompress_jsc(blob: bytes) -> bytes:
    """JustCompressedDocument: 38 バイトのヘッダの後に LHA level-1 ヘッダ + -lh5- ペイロードが 1 本。
    lhafile に [start : start+2+hsize+packed] + b"\\x00" を渡せば展開できる。"""
    if not blob.startswith(JSC_MAGIC):
        raise CompressedDocumentError("not a JustCompressedDocument stream")
    i = blob.find(b"-lh5-")
    if i < 2:
        raise CompressedDocumentError("lh5 payload not found")
    start = i - 2
    hsize = blob[start]
    packed = struct.unpack("<I", blob[start + 7:start + 11])[0]
    try:
        import lhafile
    except ImportError as e:   # pragma: no cover
        raise ImportError("reading .jtdc requires the 'lhafile' package (pip install jtdlib[jtdc])") from e
    buf = blob[start:start + 2 + hsize + packed] + b"\x00"
    lf = lhafile.LhaFile(io.BytesIO(buf))
    info = lf.infolist()[0]
    return lf.read(info.filename)


def read_words(raw: bytes) -> tuple[int, ...]:
    """UTF-16BE ストリームをワード列にする。"""
    n = len(raw) // 2
    return struct.unpack(">%dH" % n, raw[: n * 2])
