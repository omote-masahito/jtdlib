"""jtdlib — 一太郎 .jtd / .jtdc の純 Python リーダー (読み取り専用、python-docx 風 API)。

    from jtdlib import Document
    doc = Document("a.jtd")
    print(doc.text)

依存: olefile (必須)、lhafile (.jtdc を読むときのみ: pip install jtdlib[jtdc])。
構造の根拠は OpenJTD (Apache-2.0, clean-room) の RFC と、一太郎2023 で生成したコーパスの観測。
"""
from .constants import COORD_UNITS_PER_CHAR
from .container import CompressedDocumentError, JtdError, NotJtdError
from .document import Document
from .model import Cell, Column, Footnote, LayoutBox, Paragraph, Row, Run, Table

__version__ = "0.1.0"

__all__ = [
    "Document", "Paragraph", "Run", "Table", "Row", "Column", "Cell", "Footnote", "LayoutBox",
    "JtdError", "NotJtdError", "CompressedDocumentError", "COORD_UNITS_PER_CHAR", "__version__",
]
