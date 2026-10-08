"""TextV.01 領域のパーサー。DocumentText / LayoutBoxText / Footnote に共通。

処理の骨格 (RFC 0009/0010):
  * 0x1F でラン開始、0x0A 段落終端、0x0C 改ページ、0x0E 物理行終端 (表)、0x0D 脚注エントリ終端
  * 0x1C レコード: 0x0010 (段落属性 / 表の行スペック)、0x0030 (セル断片)、0x0001 (インライン)
  * 表は物理行と断片で格納されている。断片を (x0, x1) が同じで下罫線のない行をまたいで縦結合し、
    論理セルにする (_logical_cells)。
"""
from __future__ import annotations

import bisect

from .constants import (
    ENTRY_END, INLINE_END, INLINE_START, PAGE_BREAK, PARA_END, REC_CELL, REC_INLINE, REC_LINE,
    RECORD, ROW_END, SEL_AUTO_TEXT, SEL_RUBY_BASE, SEL_RUBY_TEXT, SUB_INDENT, SUB_ROWSPEC, TAB,
    TEXT_RUN,
)
from .container import read_words

BORDER_GAP = 2   # fmt 2 の行スペックで観測される罫線スパン間のギャップ (座標単位)
from .model import Cell, Paragraph, Run, Table, _Chunk, _Line


def decode_words(words) -> str:
    """UTF-16 コード単位の列を文字列にする。
    サロゲートペア (U+10000 以上、「𠮷」など) を 1 文字に復元する。chr() を語ごとに
    適用すると孤立サロゲートになり、print / json.dumps で UnicodeEncodeError になる。
    対にならない孤立サロゲートは U+FFFD に置き換える。"""
    return b"".join(x.to_bytes(2, "big") for x in words).decode("utf-16-be", "replace")


class TextParser:
    def __init__(self, data: bytes, segment: tuple[int, int] | None = None):
        self.u = read_words(data)
        self.segment = segment
        self.blocks: list[Paragraph | Table] = []
        self.para = Paragraph()
        self.pending_attrs: dict = {}
        self.table: Table | None = None
        self.line: _Line | None = None
        self.chunk: _Chunk | None = None
        self.page_break_pending = False
        self._end = len(self.u)
        self.truncated = 0            # 終端 0x1E が見つからなかったインラインの数

    @staticmethod
    def segments(u: tuple[int, ...]) -> list[tuple[int, int]]:
        """SsmgV.01 ストリーム内の全 "TextV.01" セグメント (start, end)。
        ヘッダ w[5] がセグメント数 (DocumentText=1, LayoutBoxText=枠の数, Footnote=1)。
        "TextV.01" の直後 2 ワードが領域長 (UTF-16 ユニット数、レコード含む)。"""
        magic = tuple(b"TextV.01"[i] << 8 | b"TextV.01"[i + 1] for i in range(0, 8, 2))
        out = []
        i = 0
        while i < len(u) - 6:
            if u[i:i + 4] == magic:
                length = (u[i + 4] << 16) | u[i + 5]
                start = i + 6
                end = min(start + length, len(u))
                out.append((start, end))
                i = end
            else:
                i += 1
        return out

    def region(self) -> tuple[int, int]:
        if self.segment is not None:
            return self.segment
        segs = self.segments(self.u)
        if segs:
            return segs[0]
        # フォールバック: 最初の 0x1F から末尾まで
        try:
            return self.u.index(TEXT_RUN) + 1, len(self.u)
        except ValueError:
            return 0, len(self.u)

    # --- 段落・セルの終端処理 ---
    def _flush_para(self):
        """0x0A (段落終端) に到達したとき。"""
        p = self.para
        if self.pending_attrs:
            for k, v in self.pending_attrs.items():
                setattr(p, k, v)
            self.pending_attrs = {}
        if self.page_break_pending:
            p.page_break_before = True
            self.page_break_pending = False
        if self.chunk is not None:
            # セル断片の中の 0x0A: この断片は段落終端で終わる。同一物理行に続きが来ることはない前提
            self.chunk.para = p
            self.chunk.hard = True
        else:
            self.blocks.append(p)
        self.para = Paragraph()

    def _end_cell(self):
        if self.chunk is not None:
            if not self.chunk.hard:
                self.chunk.para = self.para     # 0x0A なしで行末: ソフト改行 (次の物理行へ続く)
            self.chunk = None
        self.para = Paragraph()

    def _end_row(self):
        self._end_cell()
        self.line = None

    def _end_table(self):
        self._end_row()
        if self.table is not None:
            cells = self._logical_cells(self.table._lines)
            # 表の右端に幅 <= 罫線ギャップ (2) の空断片が現れることがある (罫線テスト.jtd:
            # 幅 120 の表で (118, 120))。右罫線の位置を占めるだけでセルではないので除く。
            w = self.table.width
            self.table.cells = [c for c in cells
                                if not (w and c.x1 >= w and c.x1 - c.x0 <= BORDER_GAP and not c.text)]
            self.blocks.append(self.table)
        self.table = None

    @staticmethod
    def _find_subentry(body: tuple[int, ...], tag: int, fixed_len: int | None = None) -> list[int] | None:
        """0x0010 レコード本体から (tag, len, payload...) 形のサブエントリを探す。"""
        for k in range(len(body) - 1):
            if body[k] == tag and (fixed_len is None or body[k + 1] == fixed_len):
                ln = body[k + 1]
                if 0 < ln <= len(body) - k - 2:
                    return list(body[k + 2:k + 2 + ln])
        return None

    @staticmethod
    def _header_entries(header: tuple[int, ...] | None) -> list[tuple[int, int, int, int]]:
        """行スペック payload [width, 0, fmt, entries..., (ffff)] の罫線スパン列を
        (tag, ?, style, width) に正規化して返す。
        fmt 0/1: 4 ワード固定。fmt 2 (古い文書): タグ 0x0008/0x0000 は (tag, width) の 2 ワード。
        style はスパン下の罫線の有無 (0 = なし)。"""
        if not header or len(header) < 3:
            return []
        fmt = header[2]
        out = []
        k = 3
        while k < len(header) and header[k] != 0xFFFF:
            tag = header[k]
            if fmt == 2 and tag in (0x0008, 0x0000):
                if k + 1 >= len(header):
                    break
                out.append((tag, 0, 8 if tag == 0x0008 else 0, header[k + 1]))
                k += 2
            else:
                if k + 3 >= len(header):
                    break
                out.append((tag, header[k + 1], header[k + 2], header[k + 3]))
                k += 4
        return out

    @staticmethod
    def _map_styles(chunks: list[_Chunk], ents: list[tuple[int, int, int, int]]) -> list[int | None]:
        """各断片に、それを覆う罫線スパンの style を対応づける。
        スパン列は表の左端からの累積位置で並ぶ (fmt 2: スパン間に幅 2 の罫線ギャップ、
        fmt 0/1: ギャップなし)。原点は「ある断片の左端がいずれかのスパンの左端に一致する」
        候補 (重複除去済み) を総当たりで選び、候補ごとの照合は累積位置が単調なので二分探索で行う。
        計算量は O(原点候補数 × 断片数 × log スパン数)。"""
        if not ents or not chunks:
            return [None] * len(chunks)
        best: tuple[int, list[int | None]] = (-1, [None] * len(chunks))
        full = 3 * len(chunks)                 # 全断片が幅まで一致したときの上限スコア
        for gap in (2, 0):
            prefix = [0]
            for e in ents:
                prefix.append(prefix[-1] + e[3] + gap)
            # 幅 0 のスパンは位置を占めないので照合対象から外す (添字は元の ents に戻せるよう保持)
            idx = [j for j, e in enumerate(ents) if e[3] != 0]
            if not idx:
                continue
            starts = [prefix[j] for j in idx]              # スパン左端 (原点基準)
            ends = [prefix[j] + ents[j][3] for j in idx]   # スパン右端
            # 同点のときは旧実装 (総当たり) と同じ候補を選ぶよう、列挙順を保った重複除去にする
            origins = dict.fromkeys(c.x0 - prefix[k] for c in chunks for k in range(len(ents)))
            for origin in origins:
                styles: list[int | None] = []
                score = 0
                for ch in chunks:
                    x = ch.x0 - origin
                    # 元の条件: 最小の j で starts[j]-1 <= x <= ends[j]。ends は単調増加なので
                    # x <= ends[j] を満たす最小の j を二分探索し、左端条件を確かめる。
                    m = bisect.bisect_left(ends, x)
                    if m < len(ends) and starts[m] - 1 <= x:
                        j = idx[m]
                        w = ents[j][3]
                        styles.append(ents[j][2])
                        score += 2 + (w in (ch.x1 - ch.x0, ch.x1 - ch.x0 + 2))
                    else:
                        styles.append(None)
                if score > best[0]:
                    best = (score, styles)
                    if score == full:
                        return best[1]
        return best[1]

    @classmethod
    def _logical_cells(cls, lines: list[_Line]) -> list[Cell]:
        """物理行の断片を論理セルに縦結合する。
        行スペックのスパン style は「その行の下の罫線」(下罫線ありは 0x08、表上端の境界行も 0x08)。
        判定: 直前の物理行に同じ (x0, x1) の断片があり、その行のスパン style が 0 (下罫線なし)
        なら同じ論理セルの続き。それ以外は新しいセル。ヘッダ省略行 (罫線配置が前行と同じ) は
        前行の style を引き継ぐ。"""
        cells: list[Cell] = []
        prev: dict[tuple[int, int], tuple[Cell, int | None]] = {}
        for li, line in enumerate(lines):
            ents = cls._header_entries(line.header)
            styles = cls._map_styles(line.chunks, ents)
            cur: dict[tuple[int, int], tuple[Cell, int | None]] = {}
            for ch, style in zip(line.chunks, styles):
                key = (ch.x0, ch.x1)
                if key in prev and not prev[key][1]:
                    cell = prev[key][0]
                    cell.line_end = li
                else:
                    cell = Cell(x0=ch.x0, x1=ch.x1, line_start=li, line_end=li, continues=bool(ch.flag))
                    cells.append(cell)
                cell._chunks.append((ch.para, ch.hard))
                if line.header is None and key in prev:
                    style = prev[key][1]
                cur[key] = (cell, style)
            prev = cur
        cells.sort(key=lambda c: (c.line_start, c.x0))
        return cells

    # --- レコード ---
    def _record(self, i: int) -> int:
        """u[i] == 0x1C。レコードを処理し、次の読み出し位置を返す。"""
        u = self.u
        if i + 2 >= len(u):
            return i + 1
        cls, ln = u[i + 1], u[i + 2]
        if ln < 3 or i + ln > len(u):
            return i + 1
        body = u[i + 3:i + ln]          # w3..w(len-1)
        w = lambda k: u[i + k] if i + k < i + ln else None

        if cls == REC_LINE:
            # w4 はフラグ (0x0026 段落属性 / 0x008f 表行 / 0x0002 / 0x0020 ...)。本体は可変長の
            # サブエントリ列で、行スペック (0x008f) と段落属性 (0x0026) をタグで探す。
            kind = w(4)
            rowspec = self._find_subentry(body, SUB_ROWSPEC)
            indent = self._find_subentry(body, SUB_INDENT, fixed_len=5)
            if rowspec is not None:      # 表の物理行
                if self.table is None:
                    if self.para.runs:
                        self._flush_para()
                    self.table = Table(width=rowspec[0] if rowspec else 0)
                else:
                    self._end_row()
                self.line = _Line(header=tuple(rowspec))
                self.table._lines.append(self.line)
            elif kind == 0xFFFF:         # 終端
                self._end_table()
            else:                        # 段落属性
                attrs: dict = {"raw_attrs": tuple(body)}
                if indent is not None:
                    attrs["indent_left"] = indent[1]
                    attrs["indent_first"] = indent[3]
                self.pending_attrs = attrs
        elif cls == REC_CELL:              # 表セル断片ヘッダ (12 ワード固定)
            if self.table is None:
                self.table = Table()
            if self.line is None:        # 罫線配置が前行と同じ物理行では 0x0010 が省略される
                self.line = _Line(header=None)
                self.table._lines.append(self.line)
            self._end_cell()
            # 001c 0030 000c 0000 [b0] [b1] 00ff [flag] 000c 0000 0030 001f
            self.chunk = _Chunk(x0=w(4) or 0, x1=w(5) or 0, flag=w(7) or 0, para=Paragraph())
            self.line.chunks.append(self.chunk)
        elif cls == REC_INLINE:              # インラインセグメント: 001c 0001 0007 0000 000x [sel] 001d <text> 001e
            selector = u[i + ln - 2] if ln >= 2 else None
            j = i + ln - 1               # レコード末尾の語が 0x1D
            if j < len(u) and u[j] == INLINE_START:
                end = self._end
                # 終端 0x1E は領域内で探す。固定上限を置くと長いインラインが黙って切れ、
                # 残りが本文に混入する (レビュー指摘)。領域内に終端がなければ壊れたレコード。
                k = j + 1
                while k < end and u[k] != INLINE_END:
                    k += 1
                truncated = k >= end
                raw = u[j + 1:k]
                text = decode_words(raw)
                if selector == SEL_RUBY_TEXT and self.para.runs and not truncated:
                    self.para.runs[-1].ruby = text
                elif raw and all(x < 0x20 for x in raw) and not truncated:
                    # 制御コードだけのインライン (0x02, 0x08 など) = 図・枠などのオブジェクトアンカー
                    self.para.runs.append(Run("", hidden=True, selector=selector))
                elif selector in (SEL_RUBY_BASE, SEL_AUTO_TEXT):
                    # 0x0001 は自動生成テキスト (脚注番号 "*1" など)。一太郎のテキスト保存にも現れる
                    self.para.runs.append(Run(text, selector=selector, truncated=truncated))
                else:
                    self.para.runs.append(Run(text, hidden=True, selector=selector, truncated=truncated))
                if truncated:
                    self.truncated += 1
                    return end
                # 0x1E の後、次の 0x1F までは小さなトレーラ
                k += 1
                while k < end and u[k] != TEXT_RUN:
                    k += 1
                return min(k + 1, end)
            return i + ln
        # cls 0x0000 (コンテキストマーカー), 0x0020 (表セクション遷移) 等は読み飛ばす
        return i + ln

    def parse(self) -> list[Paragraph | Table]:
        u = self.u
        i, end = self.region()
        self._end = end
        buf: list[int] = []           # UTF-16 コード単位。flush 時にまとめて復号する

        def flush_buf():
            if buf:
                self.para.runs.append(Run(decode_words(buf)))
                buf.clear()

        while i < end:
            x = u[i]
            if x == RECORD:
                flush_buf()
                i = self._record(i)
                continue
            # 行終端 (0x0E) の後にレコードなしで本文が続けば、表はそこで終わっている
            if self.table is not None and self.line is None and x not in (TEXT_RUN, INLINE_END):
                self._end_table()
            if x == PARA_END:
                flush_buf()
                self._flush_para()
            elif x == ROW_END:
                flush_buf()
                self._end_row()
            elif x == PAGE_BREAK:
                flush_buf()
                if self.para.runs:
                    self._flush_para()
                self.page_break_pending = True
            elif x == ENTRY_END:
                flush_buf()
                self.para.entry_end = True
                self._flush_para()
            elif x == TEXT_RUN or x == INLINE_END:
                pass
            elif x == TAB:
                buf.append(0x09)
            elif x < 0x20 or 0x80 <= x < 0xA0:
                pass                     # その他の制御コードは無視 (preservation は今後)
            else:
                buf.append(x)
            i += 1
        flush_buf()
        if self.table is not None:
            self._end_table()
        if self.para.runs:
            self._flush_para()
        return self.blocks
