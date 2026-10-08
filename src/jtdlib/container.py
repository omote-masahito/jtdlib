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
    try:
        ole = olefile.OleFileIO(data)
    except (OSError, ValueError, struct.error) as e:   # olefile は壊れた FAT/ヘッダで OSError 系を投げる
        raise NotJtdError(f"broken OLE2 compound file: {e}") from e
    if ole.exists("JSCompDocument") and not ole.exists("DocumentText"):
        try:
            inner = decompress_jsc(ole.openstream("JSCompDocument").read())
        finally:
            ole.close()
        if not inner.startswith(OLE_MAGIC):
            raise CompressedDocumentError("decompressed JSCompDocument is not an OLE2 file")
        try:
            return olefile.OleFileIO(inner)
        except (OSError, ValueError, struct.error) as e:
            raise CompressedDocumentError(f"decompressed JSCompDocument is a broken OLE2 file: {e}") from e
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
    if len(blob) < start + 11:
        raise CompressedDocumentError("truncated LHA header")
    hsize = blob[start]
    packed = struct.unpack("<I", blob[start + 7:start + 11])[0]
    end = start + 2 + hsize + packed
    if end > len(blob):
        raise CompressedDocumentError(f"truncated LHA payload: need {end} bytes, have {len(blob)}")
    try:
        import lhafile
    except ImportError as e:   # pragma: no cover
        raise ImportError("reading .jtdc requires the 'lhafile' package (pip install jtdlib[jtdc])") from e
    buf = blob[start:end] + b"\x00"
    # lhafile は壊れたヘッダで BadLhafile、壊れた圧縮データで各種例外を投げる。
    # 呼び出し側が except JtdError で括れるよう CompressedDocumentError に包む。
    try:
        lf = lhafile.LhaFile(io.BytesIO(buf))
        infos = lf.infolist()
        if not infos:
            raise CompressedDocumentError("LHA archive has no entries")
        return lf.read(infos[0].filename)
    except CompressedDocumentError:
        raise
    except Exception as e:
        raise CompressedDocumentError(f"LHA decompression failed: {type(e).__name__}: {e}") from e


def read_words(raw: bytes) -> tuple[int, ...]:
    """UTF-16BE ストリームをワード列にする。"""
    n = len(raw) // 2
    return struct.unpack(">%dH" % n, raw[: n * 2])
