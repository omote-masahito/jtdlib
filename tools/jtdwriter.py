r"""jtdwriter.py — python-docx 風の追記専用 API で一太郎マクロ文を生成する

一太郎 (JXW.Application) への受け渡しは taro.py の Taro.run() に委ねる。
このモジュール自体は「命令列 (テープ) を組み立てて文字列にする」だけで、
compile() は一太郎がない環境でも動く。save() だけが Windows + 一太郎を要求する。

使い方:
    from jtdwriter import Document
    doc = Document()
    doc.add_heading("報告書", 1)
    p = doc.add_paragraph("この文書は")
    p.add_run("自動生成", bold=True, underline=True)
    p.add_run("です。")
    t = doc.add_table(rows=2, cols=3)
    t.cell(0, 0).text = "項目"
    doc.save(r"C:\tmp\out.jtd")        # 一太郎を起動して保存
    print(doc.compile())               # 生成されるマクロ文を確認

実機確認済み (一太郎2023, 2026-09-12〜14):
  * Run は複数行一括。Insert + 入力属性系 (InputBold / InputSizePoint /
    InputDecoration) で Run 単位の書式が付く。範囲選択は不要
  * InputCharacterSize(1) では自由指定したポイント数が標準に戻らない。
    InputSizePoint() (引数省略) で戻す
  * FormatCenter は COM (Application.Run) 経由では範囲必須。SelectRangeStart(3) を
    直前に張る。文字揃えは次の段落に引き継がれるので改行後に
    CancelParagraphAttribute で解除する
  * 表: 各行をタブ区切りで Insert(…, 1) → 最終行は改行なし → SelectRangeStart(3)
    → CursorUp(n-1) → ConvertStringToKeisenTable (区切り文字省略=自動判定)。
    選択単位は 3 (段落) でないと拒否される。表の前後に空段落が 1 本ずつ残る。
    行末の空セルがあっても列数は保たれる
  * 文字列リテラルの " は "" に二重化。タブ・改行はリテラルに埋めず Char(n) で連結
  * 名前付き引数のキーワードは日本語 (半角カナ) のみ (.Underline は不可)

  * 段落書式: システム書式 (大見出し1, H1 …) は CallParagraphStyle で文書に取り込んで
    から SetParagraphStyle(1, 名前)。改行後に SetParagraphStyle(0, 名前) で解除
    (モード 0 は名前に関わらずカーソル段落の書式を外す)

  * 改ページは InsertPage (改ページ)
  * FormatCenter / FormatRight / FormatLeft は同じ手順 (範囲指定→実行→JumpEnd) で動く
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

try:
    from taro import FMT_NORMAL, Taro, macro_str
except ImportError:  # pywin32 がない環境。compile() だけ使える
    Taro = None  # type: ignore[assignment,misc]
    FMT_NORMAL = 0

    def macro_str(s: str) -> str:
        return '"' + s.replace('"', '""') + '"'


# ---------------------------------------------------------------------------
# マクロ文の組み立て
# ---------------------------------------------------------------------------

def _expr(text: str) -> str:
    """マクロの文字列式にする。タブ・改行・改ページは Char() で連結する。"""
    parts: list[str] = []
    buf = ""
    for ch in text:
        if ch in "\t\n\f":
            if buf:
                parts.append(macro_str(buf))
                buf = ""
            parts.append(f"Char({ord(ch)})")
        else:
            buf += ch
    if buf or not parts:
        parts.append(macro_str(buf))
    return " & ".join(parts)


class _Tape:
    """命令列。t("Cmd", a, b) → Cmd(a,b) / t("Cmd", kw="…") → Cmd(.kw=…)"""

    def __init__(self) -> None:
        self.lines: list[str] = []

    @staticmethod
    def _fmt(x) -> str:
        if x is None:
            return ""
        if isinstance(x, bool):
            return "1" if x else "0"
        if isinstance(x, (int, float)):
            return str(x)
        return x  # 文字列はすでにマクロ式 (macro_str / _expr 済み) とみなす

    def __call__(self, cmd: str, *args, **kw) -> None:
        pos = [self._fmt(a) for a in args]
        while pos and pos[-1] == "":   # 末尾の省略引数は落とす
            pos.pop()
        named = [f".{k}={self._fmt(v)}" for k, v in kw.items()]
        self.lines.append(f"{cmd}({','.join(pos + named)})")

    def source(self) -> str:
        return "\n".join(self.lines)


# InputDecoration / Style の名前付き引数。Python 側は英語、マクロ側は日本語固定。
_DECO_KW = {
    "color": "文字色",
    "deco_color": "飾り色",
    "underline": "ｱﾝﾀﾞｰﾗｲﾝ",
    "strike": "取消ﾗｲﾝ",
    "overline": "ｱｯﾊﾟｰﾗｲﾝ",
    "border": "文字囲い",
    "highlight": "塗りつぶし",
}

_ALIGN_CMD = {"left": "FormatLeft", "center": "FormatCenter", "right": "FormatRight"}



# ---------------------------------------------------------------------------
# python-docx 風の公開 API
# ---------------------------------------------------------------------------

@dataclass
class Font:
    name: str | None = None        # 和文フォント名
    size: float | None = None      # pt
    bold: bool = False
    italic: bool = False
    underline: bool | int = False  # True=線種1, int=線種 1〜15
    strike: bool | int = False
    color: str | None = None       # "R255G0B0" 形式

    def decoration(self) -> dict[str, int | str]:
        d: dict[str, int | str] = {}
        if self.underline:
            d["underline"] = 1 if self.underline is True else int(self.underline)
        if self.strike:
            d["strike"] = 1 if self.strike is True else int(self.strike)
        if self.color:
            d["color"] = macro_str(self.color)
        return d


@dataclass
class Run:
    text: str
    font: Font = field(default_factory=Font)

    # python-docx の run.bold 等と同じ場所に置くための転送プロパティ
    @property
    def bold(self) -> bool:
        return self.font.bold

    @bold.setter
    def bold(self, v: bool) -> None:
        self.font.bold = v

    @property
    def italic(self) -> bool:
        return self.font.italic

    @italic.setter
    def italic(self, v: bool) -> None:
        self.font.italic = v

    @property
    def underline(self):
        return self.font.underline

    @underline.setter
    def underline(self, v) -> None:
        self.font.underline = v

    def emit(self, t: _Tape) -> None:
        f = self.font
        deco = f.decoration()
        # 「入力〜」系はこれから入力する文字に属性を付ける。範囲選択は不要。
        if f.name:
            t("InputFont", macro_str(f.name))
        if f.size:
            t("InputSizePoint", f.size)
        if f.bold:
            t("InputBold", 1)
        if f.italic:
            t("InputItalic", 1)
        if deco:
            t("InputDecoration", **{_DECO_KW[k]: v for k, v in deco.items()})
        t("Insert", _expr(self.text), 0)
        # 設定した項目だけ戻す (逆順)
        if deco:
            reset = {_DECO_KW[k]: (macro_str("R0G0B0") if k == "color" else 0) for k in deco}
            t("InputDecoration", **reset)
        if f.italic:
            t("InputItalic", 0)
        if f.bold:
            t("InputBold", 0)
        if f.size:
            t("InputSizePoint")          # 引数省略で標準ポイントに戻る
        if f.name:
            t("InputFont")


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    style: str | None = None           # 一太郎側に登録済みの段落書式名
    alignment: str | None = None       # "left" | "center" | "right"
    heading_level: int | None = None   # add_heading が設定。compile 時に書式/サイズを解決

    def add_run(self, text: str, *, bold: bool = False, italic: bool = False,
                underline: bool | int = False, strike: bool | int = False,
                size: float | None = None, name: str | None = None,
                color: str | None = None) -> Run:
        r = Run(text, Font(name=name, size=size, bold=bold, italic=italic,
                           underline=underline, strike=strike, color=color))
        self.runs.append(r)
        return r

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def emit(self, t: _Tape, doc: "Document") -> None:
        style = self.style
        if self.heading_level is not None and style is None:
            style = doc.heading_styles.get(self.heading_level)
        runs = self.runs
        if self.heading_level is not None and style is None:
            # 段落書式がない環境では太字＋サイズで代替
            size = doc.heading_sizes.get(self.heading_level, 12)
            runs = [Run(r.text, Font(name=r.font.name, size=size, bold=True)) for r in self.runs]
        if style:
            t("SetParagraphStyle", 1, macro_str(style))
        for r in runs:
            r.emit(t)
        if self.alignment:
            t("SelectRangeStart", 3)          # COM 実行では範囲必須
            t(_ALIGN_CMD[self.alignment])
            t("JumpEnd")                      # 選択の後始末
        t("InsertReturn", 1, 1)
        if style:
            t("SetParagraphStyle", 0, macro_str(style))   # 次の段落への引き継ぎを解除 (モード0はカーソル段落の書式を名前に関わらず外す)
        if self.alignment:
            # 文字揃えは次の段落に引き継がれるので、新しい段落側で解除する
            t("SelectRangeStart", 3)
            t("CancelParagraphAttribute", True, False, False, False, False, False, False)
            t("JumpEnd")


class _Cell:
    def __init__(self, table: "Table", r: int, c: int) -> None:
        self._t, self._r, self._c = table, r, c

    @property
    def text(self) -> str:
        return self._t.rows[self._r][self._c]

    @text.setter
    def text(self, v: str) -> None:
        self._t.rows[self._r][self._c] = str(v)


@dataclass
class Table:
    rows: list[list[str]]

    def cell(self, r: int, c: int) -> _Cell:
        return _Cell(self, r, c)

    def add_row(self) -> list[str]:
        row = [""] * len(self.rows[0])
        self.rows.append(row)
        return row

    def emit(self, t: _Tape, doc: "Document") -> None:
        n = len(self.rows)
        cols = len(self.rows[0])
        for i, row in enumerate(self.rows):
            if len(row) != cols:
                raise ValueError(f"row {i}: 列数が揃っていない ({len(row)} != {cols})")
            for c in row:
                if "\t" in c or "\n" in c:
                    raise ValueError("セル内のタブ・改行は未対応")
        # 1行=1段落で挿入。区切りはタブ (Char(9))。最終行は改行せず末尾にカーソルを残す。
        for i, row in enumerate(self.rows):
            t("Insert", _expr("\t".join(row)), 0 if i == n - 1 else 1)
        # 段落単位で先頭行まで選択を伸ばし、罫線表に変換する (変換が選択を消費する)。
        # 区切り文字は省略 → 自動判定 (タブ→カンマ→改行の順)。
        t("SelectRangeStart", 3)
        if n > 1:
            t("CursorUp", n - 1)
        t("ConvertStringToKeisenTable", None, cols)
        # 変換後のカーソル位置は不定なので文末へ。追記専用なので表の直後と等価。
        t("JumpEnd")
        t("InsertReturn", 1, 1)


class _PageBreak:
    def emit(self, t: _Tape, doc: "Document") -> None:
        t("InsertPage")   # 改ページ (回数省略=1)。文字列中の \f は _expr が Char(12) にする


class Document:
    """追記専用。add_* で積んだ順にマクロ文を生成する。"""

    def __init__(self) -> None:
        self._body: list[Paragraph | Table | _PageBreak] = []
        #: add_heading(level) → 一太郎の段落書式名。GetParagraphStyleList で得た名前を入れる。
        #: 未設定の level は heading_sizes の太字＋サイズで代替する。compile 時に解決する。
        self.heading_styles: dict[int, str] = {}
        #: 段落書式を使わないときの見出しサイズ (pt)
        self.heading_sizes: dict[int, float] = {1: 16, 2: 14, 3: 12}

    # --- 構築 ---
    def add_paragraph(self, text: str = "", style: str | None = None,
                      alignment: str | None = None) -> Paragraph:
        p = Paragraph(style=style, alignment=alignment)
        if text:
            p.add_run(text)
        self._body.append(p)
        return p

    def add_heading(self, text: str, level: int = 1) -> Paragraph:
        p = Paragraph(heading_level=level)
        p.add_run(text)
        self._body.append(p)
        return p

    def add_table(self, rows: int, cols: int, data: list[list[str]] | None = None) -> Table:
        if data is not None:
            tb = Table([[str(c) for c in r] for r in data])
        else:
            tb = Table([[""] * cols for _ in range(rows)])
        self._body.append(tb)
        return tb

    def add_page_break(self) -> None:
        self._body.append(_PageBreak())

    @property
    def paragraphs(self) -> list[Paragraph]:
        return [x for x in self._body if isinstance(x, Paragraph)]

    @property
    def tables(self) -> list[Table]:
        return [x for x in self._body if isinstance(x, Table)]

    # --- 出力 ---
    def compile(self) -> str:
        t = _Tape()
        # 段落書式はシステムから文書に取り込む (CallParagraphStyle) まで
        # SetParagraphStyle が「登録されていません」になる。取り込んだ書式は
        # 文書内のユーザースタイルとして保存される (docx の styles.xml 相当)。
        for name in self._used_styles():
            t("CallParagraphStyle", macro_str(name))
        for item in self._body:
            item.emit(t, self)
        return t.source()

    def _used_styles(self) -> list[str]:
        seen: list[str] = []
        for p in self.paragraphs:
            name = p.style
            if name is None and p.heading_level is not None:
                name = self.heading_styles.get(p.heading_level)
            if name and name not in seen:
                seen.append(name)
        return seen

    def save(self, path: str | Path, fmt: int = FMT_NORMAL, *,
             visible: bool = False, batch: bool = True, taro: Taro | None = None) -> Path | None:
        """一太郎を起動し、新規文書に流し込んで保存する。
        batch=False で 1 文ずつ run する (エラー箇所の切り分け用、遅い)。
        taro を渡すと既存の Taro セッションを使い、終了させない。"""
        if Taro is None:
            raise RuntimeError("taro.py (pywin32) が読み込めない環境では save() は使えない")
        source = self.compile()

        def _do(tr: Taro):
            tr.add()
            if batch:
                tr.run(source)
            else:
                for line in source.splitlines():
                    tr.run(line)
            return tr.save_document(path, fmt)

        if taro is not None:
            return _do(taro)
        with Taro(visible=visible) as tr:
            return _do(tr)


# ---------------------------------------------------------------------------
def demo() -> Document:
    doc = Document()
    doc.add_heading("jtdwriter サンプル", 1)
    p = doc.add_paragraph("python-docx 風の API で組み立てた文書を、")
    p.add_run("一太郎マクロ", bold=True)
    p.add_run("に変換して保存する。")
    doc.add_paragraph("中央揃え", alignment="center")
    p = doc.add_paragraph("下線と")
    p.add_run("赤字", color="R255G0B0", underline=True)
    p.add_run('と"引用符"、タブ区切りの表 (100%)。')
    doc.add_table(0, 0, [
        ["機能", "python-docx", "一太郎マクロ", "備考"],
        ["段落", "add_paragraph", "Insert / InsertReturn", "追記専用"],
        ["文字書式", "Run.font", "InputBold ほか", "範囲選択不要"],
        ["文字揃え", "alignment", "FormatCenter ほか", "範囲選択が必要"],
        ["段落書式", "heading_styles", "SetParagraphStyle", "CallParagraphStyle で取込"],
        ["罫線表", "add_table", "ConvertStringToKeisenTable", "列幅は均等"],
        ["改ページ", "add_page_break", "InsertPage", ""],
    ])
    doc.add_paragraph("表の後")
    doc.add_page_break()
    doc.add_paragraph("2ページ目、右揃え", alignment="right")
    doc.add_page_break()
    doc.add_paragraph("3ページ目、左揃え", alignment="left")
    return doc


if __name__ == "__main__":
    import sys
    doc = demo()
    if len(sys.argv) > 1:
        print(doc.save(sys.argv[1], visible=True))
    else:
        print(doc.compile())
