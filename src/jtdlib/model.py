"""読み取り専用のオブジェクトモデル (python-docx 風)。

    Document
      .paragraphs        -> list[Paragraph]
      .tables            -> list[Table]
      .iter_inner_content() -> Paragraph | Table を本文順に
      .footnotes         -> list[Footnote]
      .layout_boxes      -> list[LayoutBox]
    Paragraph.runs -> list[Run];  Run.text / Run.ruby
    Table.rows -> list[Row];  Row.cells -> list[Cell];  Cell.paragraphs / Cell.text
    Table.cell(r, c), Table.columns

python-docx との対応で意図的に違うところ:
  * 一太郎の表は「物理行 + セル断片」でしか格納されないので、Row は「同じ物理行から始まる
    論理セルの集まり」という近似。結合セルは 1 つの Cell として現れる (python-docx のように
    結合範囲に同じ Cell が繰り返されることはない)。
  * 書式 (フォント・スタイル) は持たない。段落インデントと改ページだけ。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterator

from .constants import COORD_UNITS_PER_CHAR


@dataclass
class Run:
    text: str
    ruby: str | None = None          # ルビ (ふりがな)。text が親文字
    hidden: bool = False             # 表示されないインライン (テンプレート指示文、オブジェクトアンカー)
    selector: int | None = None      # 0x0001 インラインレコードのセレクタ
    truncated: bool = False          # インラインの終端 0x1E が領域内に無く、領域末尾まで読んだ

    def __repr__(self) -> str:
        extra = f" ruby={self.ruby!r}" if self.ruby else ""
        extra += " hidden" if self.hidden else ""
        extra += " truncated" if self.truncated else ""
        return f"Run({self.text!r}{extra})"


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    indent_left: int | None = None   # 段落左端 (単位: カラム = 文字)
    indent_first: int | None = None  # 先頭行の絶対位置 (単位: カラム)
    page_break_before: bool = False
    entry_end: bool = False          # 直後に 0x000D (脚注エントリ終端) があった (Footnote ストリームのみ)
    raw_attrs: tuple[int, ...] | None = None  # 0x0010 レコード本体の生ワード (解析用)

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs if not r.hidden)

    def text_with_ruby(self, fmt: str = "{base}({ruby})") -> str:
        """一太郎のテキスト保存と同じ「親文字(ルビ)」表記。"""
        return "".join(fmt.format(base=r.text, ruby=r.ruby) if r.ruby else r.text
                       for r in self.runs if not r.hidden)

    def __repr__(self) -> str:
        return f"Paragraph({self.text!r})"


@dataclass
class _Chunk:
    """物理行に現れたセル断片 (パース中のみ使う)。"""
    x0: int
    x1: int
    flag: int
    para: Paragraph
    hard: bool = False               # 断片末尾が 0x0A (段落境界) か


@dataclass
class _Line:
    """表の物理行。header は 0x0010/0x008f 行スペック (罫線配置が前行と同じなら None)。"""
    header: tuple[int, ...] | None
    chunks: list[_Chunk] = field(default_factory=list)


@dataclass
class Cell:
    """論理セル。同じ (x0, x1) のセル断片を、下罫線のない物理行をまたいで縦結合したもの。

    x0/x1 は表の左端からの座標で、単位は 1 文字幅の 1/4 (COORD_UNITS_PER_CHAR)。
    """
    x0: int = 0
    x1: int = 0
    line_start: int = 0              # 表内の物理行番号 (0 始まり)
    line_end: int = 0
    continues: bool = False          # 0x0030 レコードのフラグ語 (セル内改行時に非 0 を観測)
    _chunks: list[tuple[Paragraph, bool]] = field(default_factory=list, repr=False)

    @property
    def paragraphs(self) -> list[Paragraph]:
        """断片を結合した段落列。0x0A で終わる断片が段落境界、終わらない断片は次の断片へ
        ソフト改行で続く。最後の断片が 0x0A で終わっていなければ空段落を 1 つ立てる
        (一太郎のテキスト保存 (セル単位) と同じ)。"""
        paras: list[Paragraph] = []
        acc: list[Run] = []
        hard_last = True
        for p, hard in self._chunks:
            acc.extend(p.runs)
            hard_last = hard
            if hard:
                paras.append(Paragraph(runs=acc))
                acc = []
        if not hard_last or not self._chunks:
            paras.append(Paragraph(runs=acc))
        return paras

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)

    def text_with_ruby(self, fmt: str = "{base}({ruby})") -> str:
        return "\n".join(p.text_with_ruby(fmt) for p in self.paragraphs)

    @property
    def width(self) -> int:
        """幅 (座標単位)。文字数換算は width / COORD_UNITS_PER_CHAR。"""
        return self.x1 - self.x0

    @property
    def width_chars(self) -> float:
        return self.width / COORD_UNITS_PER_CHAR

    @property
    def n_lines(self) -> int:
        return self.line_end - self.line_start + 1

    def __repr__(self) -> str:
        return f"Cell(x={self.x0}..{self.x1}, lines={self.line_start}..{self.line_end}, {self.text!r})"


@dataclass
class Row:
    """同じ物理行から始まる論理セルの集まり (近似的な「表の行」)。"""
    cells: list[Cell] = field(default_factory=list)
    line: int = 0                    # 開始物理行

    @property
    def is_border(self) -> bool:
        """本文のない全幅セル 1 個だけの行 (表の上端境界として現れる)。"""
        return len(self.cells) == 1 and self.cells[0].x0 == 0 and not self.cells[0].text.strip("\n")

    @property
    def text(self) -> str:
        return "\t".join(c.text for c in self.cells)

    def __iter__(self) -> Iterator[Cell]:
        return iter(self.cells)

    def __len__(self) -> int:
        return len(self.cells)


@dataclass
class Column:
    """x 座標の区間で定義した列。cells はその区間に左端が含まれるセル (上から順)。"""
    x0: int
    x1: int
    cells: list[Cell] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(c.text for c in self.cells)


@dataclass
class Table:
    width: int = 0                   # 表全幅 (座標単位)
    cells: list[Cell] = field(default_factory=list)      # 論理セル (開始行, x0 順)
    _lines: list[_Line] = field(default_factory=list, repr=False)  # 物理行 (生データ)

    @property
    def raw_rows(self) -> list[Row]:
        """境界行を含む全行。"""
        rows: dict[int, Row] = {}
        for c in self.cells:
            rows.setdefault(c.line_start, Row(line=c.line_start)).cells.append(c)
        return [rows[k] for k in sorted(rows)]

    @property
    def rows(self) -> list[Row]:
        return [r for r in self.raw_rows if not r.is_border]

    @property
    def columns(self) -> list[Column]:
        """セル左端 x0 の集合で列を切る。結合セルは左端の列にだけ入る。"""
        cells = [c for r in self.rows for c in r.cells]
        xs = sorted({c.x0 for c in cells})
        if not xs:
            return []
        bounds = xs + [max(c.x1 for c in cells)]
        cols = [Column(x0=bounds[i], x1=bounds[i + 1]) for i in range(len(xs))]
        for c in sorted(cells, key=lambda c: (c.line_start, c.x0)):
            cols[xs.index(c.x0)].cells.append(c)
        return cols

    def cell(self, row_idx: int, col_idx: int) -> Cell:
        """rows[row_idx].cells[col_idx]。python-docx と違い結合セルの補完はしない。"""
        return self.rows[row_idx].cells[col_idx]

    @property
    def n_lines(self) -> int:
        return len(self._lines)

    @property
    def text(self) -> str:
        return "\n".join(r.text for r in self.rows)

    def __iter__(self) -> Iterator[Row]:
        return iter(self.rows)


@dataclass
class Footnote:
    marker: str                      # 本文中の脚注番号と同じ文字列 (例 "*1")
    paragraphs: list[Paragraph] = field(default_factory=list)
    # FootnoteLink ストリーム由来
    position: int | None = None      # 本文領域内の脚注マーカー (0x0000 レコード) の位置 (ワード)
    page: int | None = None          # ページ番号 (0 始まり)
    number_in_page: int | None = None  # ページ内通し番号 (1 始まり)

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)

    def __repr__(self) -> str:
        return f"Footnote({self.marker!r}, {self.text[:40]!r})"


class _BlockContainer:
    """Paragraph と Table の列を持つもの (Document 本文、LayoutBox) の共通部分。"""
    blocks: list[Paragraph | Table]

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        return iter(self.blocks)

    @property
    def paragraphs(self) -> list[Paragraph]:
        return [b for b in self.blocks if isinstance(b, Paragraph)]

    @property
    def tables(self) -> list[Table]:
        return [b for b in self.blocks if isinstance(b, Table)]

    def iter_strings(self, ruby: bool = False) -> Iterator[str]:
        """一太郎の GetString(2) / テキスト保存 (セル単位) と同じ順序で文字列を返す。
        ruby=True なら「親文字(ルビ)」表記。"""
        for b in self.blocks:
            if isinstance(b, Paragraph):
                if b.page_break_before:
                    yield ""             # テキスト保存は改ページを空行として出す
                yield b.text_with_ruby() if ruby else b.text
            else:
                for c in b.cells:
                    paras = c.paragraphs
                    while len(paras) > 1 and not paras[0].text:   # 先頭の空段落は書き出しに現れない
                        paras = paras[1:]
                    yield "\n".join(p.text_with_ruby() if ruby else p.text for p in paras)

    @property
    def text(self) -> str:
        return "\n".join(self.iter_strings())


@dataclass
class LayoutBox(_BlockContainer):
    """レイアウト枠 (テキストボックス)。ODT の draw:text-box に対応。
    一太郎のテキスト保存には含まれない。"""
    index: int
    blocks: list[Paragraph | Table] = field(default_factory=list)

    def __repr__(self) -> str:
        return f"LayoutBox({self.index}, {self.text[:40]!r})"
