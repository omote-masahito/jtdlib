"""taro.py — 一太郎 (JXW.Application) の COM ラッパー

pywin32 で判明した規約 (probe1〜8):
  * プロパティの get/set は dynamic Dispatch のままでよい (Visible, Documents, Count, ActiveDocument, Name ...)
  * メソッドは dynamic 層を通すと壊れる (属性参照だけで引数ゼロ Invoke が走る)。
    必ず GetIDsOfNames -> IDispatch::Invoke(DISPATCH_METHOD) を直接叩く。
  * 引数は VARIANT で型を明示する (str->VT_BSTR, int->VT_I4, float->VT_R8)。
  * Taro33.TLB の DISPID は実行時と一致しないので、makepy 生成クラスは使わない。名前解決のみ信用する。
  * マクロは Application.Run(ソース文字列, trStatementMacro=3) で実行できる。複数行可。
    返り値のある関数は、マクロ内で Open/LinePrint/Close を使ってファイルに書き、Python で読む。
  * TaroLibrary の命令は生 Invoke でも叩けるが、Run 経由なら型の心配がないので基本は Run。

使い方:
  from taro import Taro
  with Taro(visible=True) as t:
      t.add()                                   # 新規文書
      t.run('Insert("こんにちは", 1)')          # マクロ文
      lines = t.get_strings()                   # GetString(2): 罫線セルも要素にした本文配列
      t.save_as("out.jtd", 0)
      t.save_as("out.docx", 20)
      t.close()                                 # 保存せず閉じる

  py -3.13-32 taro.py selftest        # 動作確認
  py -3.13-32 taro.py selftest scan   # + 保存形式コードの総当たり
"""
from __future__ import annotations

import sys
import tempfile
import time
from pathlib import Path

import pythoncom
import win32com.client as w32
from win32com.client import VARIANT

# Constants (Taro33.TLB)
TR_STATEMENT_MACRO = 3
TR_FILE_MACRO = 9
TR_SAVE, TR_NOTSAVE, TR_CONFIRM = 2, 3, 1

# SaveDocument / SaveAs の保存形式
FMT_NORMAL, FMT_TEMPLATE, FMT_NORMAL_COMPRESSED, FMT_TEMPLATE_COMPRESSED = 0, 1, 5, 6
FMT_TEXT, FMT_RTF, FMT_HTML, FMT_WORD, FMT_ODF, FMT_PDF = 10, 15, 16, 20, 22, 23

# SaveDocument が形式ごとに強制する拡張子 (一太郎2023 で実測)。20 の .doc は中身が RTF。
FMT_EXT = {0: "jtd", 1: "jtt", 5: "jtdc", 6: "jttc", 10: "txt", 11: "txt", 12: "txt", 13: "txt",
           15: "rtf", 16: "htm", 17: "htm", 20: "doc", 21: "ppt", 22: "odt", 23: "pdf",
           40: "jsw", 50: "jaw", 51: "jtw", 60: "jbw", 61: "juw", 70: "jfw", 71: "jvw"}


def _v(x):
    if isinstance(x, VARIANT):
        return x
    if isinstance(x, bool):
        return VARIANT(pythoncom.VT_BOOL, x)
    if isinstance(x, int):
        return VARIANT(pythoncom.VT_I4, x)
    if isinstance(x, float):
        return VARIANT(pythoncom.VT_R8, x)
    if isinstance(x, str):
        return VARIANT(pythoncom.VT_BSTR, x)
    if isinstance(x, Path):
        return VARIANT(pythoncom.VT_BSTR, str(x))
    raise TypeError(f"unsupported COM arg: {type(x).__name__}")


def call(obj, name: str, *args):
    """dynamic 層を迂回して IDispatch::Invoke を直接叩く。"""
    o = obj._oleobj_ if hasattr(obj, "_oleobj_") else obj
    r = o.Invoke(o.GetIDsOfNames(0, name), 0, pythoncom.DISPATCH_METHOD, 1, *[_v(a) for a in args])
    if isinstance(r, pythoncom.TypeIIDs[pythoncom.IID_IDispatch]):
        return w32.Dispatch(r)
    return r


class MacroError(RuntimeError):
    def __init__(self, cause, source):
        desc = ""
        try:
            desc = cause.args[2][2].strip()
        except Exception:
            pass
        super().__init__(f"{desc or cause}\n--- macro ---\n{source}\n-------------")
        self.cause, self.source = cause, source


def macro_str(s: str) -> str:
    """一太郎マクロの文字列リテラルにする (二重引用符は "" でエスケープ)。"""
    return '"' + s.replace('"', '""') + '"'


class Taro:
    def __init__(self, visible: bool = False, startup_wait: float = 3.0):
        self.app = w32.Dispatch("JXW.Application")
        time.sleep(startup_wait)
        try:
            self.app.Visible = visible
        except Exception:
            pass
        self.docs = self.app.Documents

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.quit()

    # --- 文書の開閉 ---
    @property
    def count(self) -> int:
        return int(self.docs.Count)

    @property
    def active(self):
        return self.app.ActiveDocument

    def add(self):
        """新規文書を開いてアクティブにする。"""
        call(self.docs, "Add")
        return self.active

    def open(self, path: Path | str, forced: int = 1):
        """Documents.Open(FileName, PassWord, FileType, TextCodeSet, AsForm, FastLoad, Forced, OpenMode)"""
        call(self.docs, "Open", str(Path(path).resolve()), "", 0, "Auto", 0, 1, forced)
        return self.active

    def close(self, doc=None, save: int = TR_NOTSAVE):
        call(doc or self.active, "Close", save)

    def quit(self, save: int = TR_NOTSAVE):
        try:
            call(self.app, "Quit", save)
        except Exception:
            pass

    # --- 保存 ---
    def save_as(self, path: Path | str, fmt: int = FMT_NORMAL, codeset: str = "ShiftJIS", doc=None):
        """Document.SaveAs(FileName, Title, PassWord, KeyWord, FileType, TextCodeSet, Fast, ReadOnly, LinkItemNumber)"""
        p = Path(path).resolve()
        if p.exists():
            p.unlink()
        call(doc or self.active, "SaveAs", str(p), "", "", "", fmt, codeset, 1, 0, 5)
        return p

    def save_document(self, path: Path | str, fmt: int = FMT_NORMAL, codeset: str = "ShiftJIS",
                      cell_mode: int = 0, newline_sub: str = ""):
        """TaroLibrary.SaveDocument をマクロ経由で。テキスト形式の罫線内文字列=セル単位 など、
        SaveAs にない引数が要るときに使う。"""
        p = Path(path).resolve()
        ext = FMT_EXT.get(fmt)
        expected = p.with_suffix("." + ext) if ext else None
        # 一太郎は形式に応じて拡張子を書き換える。消すのは「これから書くファイル」だけ。
        # (同じ stem の別形式は直前の SaveDocument で一太郎がロックしている可能性がある)
        if expected is not None and expected.exists():
            expected.unlink()
        self.run(f'SaveDocument({macro_str(str(p))}, "", "", {fmt}, 1, 0, 5, 0, 1, '
                 f'{macro_str(codeset)}, "", 1, {cell_mode}, {macro_str(newline_sub)})')
        return self._find_saved(p, expected)

    @staticmethod
    def _find_saved(p: Path, expected: Path | None, timeout: float = 5.0) -> Path | None:
        t0 = time.time()
        while time.time() - t0 < timeout:
            if expected is not None:
                if expected.exists() and expected.stat().st_size > 0:
                    return expected
            else:  # 未知の形式コード: stem 一致で最新のものを拾う
                hits = [q for q in p.parent.glob(p.stem + ".*") if q.stem == p.stem and q.stat().st_size > 0]
                if hits:
                    return max(hits, key=lambda q: q.stat().st_mtime)
            time.sleep(0.2)
        return None

    # --- マクロ ---
    def run(self, source: str):
        """マクロ文 (複数行可) をアクティブ文書に対して実行する。
        一太郎は全体をコンパイルしてから実行する: 構文エラーなら 1 文も実行されない。
        実行時エラーならそこまでは実行済み。失敗時はソースを例外に添える。"""
        try:
            call(self.app, "Run", source, TR_STATEMENT_MACRO)
        except Exception as e:
            raise MacroError(e, source) from e

    def run_capture(self, body: str, timeout: float = 10.0) -> list[str]:
        """body の中で LinePrint(1, ...) したものを行リストで返す。ファイル番号 1 は予約。"""
        with tempfile.TemporaryDirectory() as d:
            result = Path(d) / "capture.txt"
            self.run(f'Create(1, {macro_str(str(result))}, "NEWOLD", "TEXT")\n{body}\nClose(1)\n')
            t0 = time.time()
            while not result.exists() and time.time() - t0 < timeout:
                time.sleep(0.1)
            if not result.exists():
                raise RuntimeError("macro produced no output file")
            raw = result.read_bytes()
        for enc in ("cp932", "utf-8"):
            try:
                return raw.decode(enc).splitlines()
            except UnicodeDecodeError:
                continue
        return raw.decode("cp932", errors="replace").splitlines()

    def get_strings(self, mode: int = 2) -> list[str]:
        """全選択して GetString(mode)。mode 2 = 罫線セル内の文字列も配列要素にする。"""
        lines = self.run_capture(f'''SelectAll(1)
%s = GetString({mode})
LinePrint(1, "count=" & Size(%s))
for %i = 1 to Size(%s)
    LinePrint(1, %s(%i))
next
CancelAllRange()''')
        return lines

    # --- 保存形式の総当たり ---
    def scan_formats(self, out_dir: Path, codes=range(0, 100)) -> dict[int, str]:
        """SaveDocument の保存形式コードを総当たりし、出力ファイルの先頭バイトで種類を判定する。"""
        out_dir.mkdir(parents=True, exist_ok=True)
        found = {}
        for fmt in codes:
            p = out_dir / f"fmt{fmt:02d}.x"
            try:
                p = self.save_document(p, fmt)
            except Exception as e:
                try:
                    desc = e.args[2][2] or str(e)
                except Exception:
                    desc = str(e)
                found[fmt] = "ERR " + desc.strip()[:40]
                continue
            if p is None:
                found[fmt] = "no file"
                continue
            head = p.read_bytes()[:8]
            kind = {b"PK": "zip(docx/odt/pptx)", b"{\\rtf": "rtf", b"%PDF": "pdf", b"\xd0\xcf\x11\xe0": "cfb(jtd/doc)",
                    b"<!DO": "html", b"<htm": "html", b"\xef\xbb\xbf": "utf8-bom text", b"\xff\xfe": "utf16le text"}
            k = next((v for m, v in kind.items() if head.startswith(m)), f"other {head!r}")
            found[fmt] = f"{p.suffix:<6} {k} {p.stat().st_size}B"
        return found

    def insert_table(self, rows: list[list[str]]):
        """タブ区切りで挿入したテキストを罫線表に変換する。rows[0][0] は文書内で一意であること。"""
        # ヘルプ用例に倣う: 区切り文字は省略 (タブ->カンマ->改行の順に自動判定)、
        # 範囲は SelectRangeStart(2) (行単位) + CursorDown。文字列リテラル内に生のタブは置かない。
        lines = "\n".join("Insert(" + " & Char(9) & ".join(macro_str(c) for c in r) + ", 1)" for r in rows)
        self.run(f'''{lines}
JumpSearchString({macro_str(rows[0][0])}, 3, 0)
SelectRangeStart(2)
CursorDown({len(rows) - 1})
ConvertStringToKeisenTable(, {len(rows[0])}, , 0)
JumpEnd()''')

    def version(self) -> list[str]:
        return self.run_capture('LinePrint(1, GetVersion(1))\nLinePrint(1, GetVersion(2))\nLinePrint(1, GetVersion(3))')


# ---------------------------------------------------------------------------
def selftest(out_dir: Path, scan: bool = False):
    out_dir.mkdir(parents=True, exist_ok=True)

    def step(label, fn):
        try:
            r = fn()
            print(f"   OK   {label} -> {r!r}")
            return r
        except Exception as e:
            print(f"   FAIL {label}: {e}")
            return None

    with Taro(visible=True) as t:
        print("== version (GetVersion をマクロ経由で)")
        step("version()", t.version)
        print("== add / insert / table / ruby")
        step("add()", t.add)
        step("count", lambda: t.count)
        step("Insert x3", lambda: t.run('Insert("一段落目", 1)\nInsert("二段落目", 1)'))
        step("表 (insert_table)", lambda: t.insert_table([["見出しA", "見出しB"], ["a1", "b1"]]))
        step("表の後の段落", lambda: t.run('Insert("表の後", 1)'))
        step("ルビ", lambda: t.run('Insert("銀河鉄道", 1)\nSelectString("銀河", 3, 0)\nRuby(1, "ぎんが")\nJumpEnd()'))
        print("== get_strings(2)")
        lines = step("get_strings()", t.get_strings)
        if lines:
            for l in lines:
                print("      |", l)
        print("== save_document 各形式 (拡張子は一太郎が決める)")
        for name, fmt, cs in (("st", FMT_NORMAL, "ShiftJIS"), ("st_c", FMT_NORMAL_COMPRESSED, "ShiftJIS"),
                              ("st_w", FMT_WORD, "ShiftJIS"), ("st_o", FMT_ODF, "ShiftJIS"),
                              ("st_t", FMT_TEXT, "UTF8")):
            p = step(f"save_document({name}, {fmt})", lambda: t.save_document(out_dir / (name + ".x"), fmt, cs))
            print(f"        -> {p.name if p else None} size={p.stat().st_size if p else '-'} head={p.read_bytes()[:6] if p else '-'}")
        p = step("save_document(st_cells, セル単位)", lambda: t.save_document(out_dir / "st_cells.x", FMT_TEXT, "UTF8", 1, "\uE000"))
        if p:
            print("        cells:", p.read_text(encoding="utf-8-sig").splitlines())
        if scan:
            print("== SaveDocument 保存形式スキャン")
            for fmt, kind in (step("scan_formats", lambda: t.scan_formats(out_dir / "fmtscan",
                                   (0, 1, 5, 6, 10, 11, 12, 15, 16, 17, 20, 21, 22, 23, 40, 50, 51, 60, 61, 70, 71))) or {}).items():
                print(f"      {fmt:3d}: {kind}")
        print("== close / reopen")
        step("close()", t.close)
        step("count", lambda: t.count)
        for name in ("st.jtd", "st_c.jtdc"):
            if (out_dir / name).exists():
                step(f"open({name})", lambda: t.open(out_dir / name))
                step("active.FullName", lambda: t.active.FullName)
                step("get_strings() (再読込後)", t.get_strings)
                step("close()", t.close)
        input("画面を確認したら Enter で Quit ...")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "selftest":
        selftest(Path("selftest_out"), scan="scan" in sys.argv[2:])
    else:
        print(__doc__)
