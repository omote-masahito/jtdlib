"""Document: 一太郎文書を開いて本文・レイアウト枠・脚注を組み立てる。"""
from __future__ import annotations

import os
from pathlib import Path
from typing import IO, Iterator

from .constants import SEL_AUTO_TEXT
from .container import NotJtdError, open_ole, read_words
from .model import Footnote, LayoutBox, Paragraph, Table, _BlockContainer
from .parser import TextParser

STREAM_TEXT = "DocumentText"
STREAM_LAYOUT_BOX = "LayoutBoxText"
STREAM_FOOTNOTE = "Footnote"
STREAM_FOOTNOTE_LINK = "FootnoteLink"

# 一太郎文書と判定するために、どれか 1 つはあるべきストリーム
_SIGNATURE_STREAMS = (STREAM_TEXT, "DocumentInfo", "DocumentViewStyles", "Font", "LayoutBox")


class Document(_BlockContainer):
    """一太郎 .jtd / .jtdc の読み取り専用ドキュメント。

        doc = Document("a.jtd")
        for p in doc.paragraphs: print(p.text)
        for t in doc.tables:
            for row in t.rows: print([c.text for c in row.cells])
        for b in doc.iter_inner_content(): ...   # Paragraph | Table を本文順に
        doc.text                                  # 全文 (セルの並び順は一太郎のテキスト保存と一致しないことがある)

    source: パス、bytes、または open() したバイナリファイルオブジェクト。
    """

    def __init__(self, source: str | os.PathLike | bytes | IO[bytes]):
        if isinstance(source, (bytes, bytearray)):
            data = bytes(source)
        elif hasattr(source, "read"):
            data = source.read()
        else:
            data = Path(source).read_bytes()
        self._ole = open_ole(data)
        self.stream_names: list[str] = ["/".join(e) for e in self._ole.listdir(streams=True, storages=True)]
        if not any(self._ole.exists(s) for s in _SIGNATURE_STREAMS):
            raise NotJtdError("OLE2 file without 一太郎 document streams")

        self.blocks: list[Paragraph | Table] = []
        if self._ole.exists(STREAM_TEXT):        # 空文書には DocumentText ストリームが無い
            self.blocks = self._parse_stream(STREAM_TEXT)[0]
        self.layout_boxes: list[LayoutBox] = [
            LayoutBox(index=k, blocks=b) for k, b in enumerate(self._parse_stream(STREAM_LAYOUT_BOX))
        ]
        self.footnotes: list[Footnote] = self._extract_footnotes(self._parse_stream(STREAM_FOOTNOTE))
        self._attach_footnote_links()

    # --- リソース管理 ---
    def close(self) -> None:
        self._ole.close()

    def __enter__(self) -> "Document":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def stream(self, name: str) -> bytes:
        """生ストリームを返す (解析用)。"""
        return self._ole.openstream(name).read()

    # --- ストリーム解析 ---
    def _parse_stream(self, name: str) -> list[list[Paragraph | Table]]:
        if not self._ole.exists(name):
            return []
        raw = self.stream(name)
        segs = TextParser.segments(read_words(raw))
        if not segs:
            return [TextParser(raw).parse()] if name == STREAM_TEXT else []
        return [TextParser(raw, seg).parse() for seg in segs]

    def _attach_footnote_links(self) -> None:
        """FootnoteLink: ヘッダ 6 ワード (w4 = 件数)、以降 15 ワード/件:
        [001a, next, prev, 0, position, ffff, ffff, 0100, next, prev, num_in_page, page, index, ffff, next]
        (next/prev は 0 始まりの索引、端は ffff)"""
        if not self._ole.exists(STREAM_FOOTNOTE_LINK):
            return
        u = read_words(self.stream(STREAM_FOOTNOTE_LINK))
        if len(u) < 6:
            return
        entries = []
        i = 6
        while i + 15 <= len(u) and u[i] == 0x001A:
            entries.append(u[i:i + 15])
            i += 15
        for fn, e in zip(self.footnotes, entries):
            fn.position, fn.number_in_page, fn.page = e[4], e[10], e[11]

    @staticmethod
    def _extract_footnotes(segments: list[list[Paragraph | Table]]) -> list[Footnote]:
        """Footnote ストリーム: エントリ = [0x0000 マーカー][0x0001 インライン 番号] 本文… 0x000D。
        末尾には雛形由来の "Note" エントリ (0x000D なし) が残るので除く。"""
        out: list[Footnote] = []
        for blocks in segments:
            cur: Footnote | None = None
            for b in blocks:
                if not isinstance(b, Paragraph):
                    continue
                runs = list(b.runs)
                if cur is None:
                    marker = ""
                    if runs and runs[0].selector == SEL_AUTO_TEXT:
                        marker = runs[0].text
                        runs = runs[1:]
                    cur = Footnote(marker=marker)
                cur.paragraphs.append(Paragraph(runs=runs, indent_left=b.indent_left, indent_first=b.indent_first))
                if b.entry_end:
                    out.append(cur)
                    cur = None
            if cur is not None and not (cur.marker == "Note" and not cur.text.strip()):
                out.append(cur)
        return out

    # --- 便利プロパティ ---
    @property
    def layout_box_texts(self) -> list[str]:
        return [b.text for b in self.layout_boxes]

    @property
    def footnote_texts(self) -> list[str]:
        return [f"{fn.marker} {fn.text}".strip() for fn in self.footnotes]

    def iter_all_text(self) -> Iterator[str]:
        """本文、レイアウト枠、脚注の順に全テキストを返す (全文検索・索引用)。"""
        yield from self.iter_strings()
        for b in self.layout_boxes:
            yield from b.iter_strings()
        for fn in self.footnotes:
            yield fn.text

    def __repr__(self) -> str:
        return (f"Document(paragraphs={len(self.paragraphs)}, tables={len(self.tables)}, "
                f"layout_boxes={len(self.layout_boxes)}, footnotes={len(self.footnotes)})")
