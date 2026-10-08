"""インラインセグメントの長さと終端。固定上限 (旧 300/310 語) で黙って切れないこと。"""
from jtdlib.parser import TextParser


def inline(utf16, text: str, selector: int = 0x0001) -> list[int]:
    return [0x1C, 0x0001, 0x0007, 0x0000, 0x0000, selector, 0x1D, *utf16(text), 0x1E, 0x1F]


def test_long_inline_is_extracted_in_full(synth):
    utf16, words = synth.utf16, synth.words
    for n in (299, 300, 350, 1000):
        body = "あ" * n
        data = words(0x1F, *utf16("前"), *inline(utf16, body), *utf16("後"), 0x0A)
        blocks = TextParser(data).parse()
        runs = blocks[0].runs
        assert [r.text for r in runs] == ["前", body, "後"], n
        assert not any(r.truncated for r in runs)


def test_long_hidden_inline_does_not_leak_into_body(synth):
    utf16, words = synth.utf16, synth.words
    body = "x" * 350
    data = words(0x1F, *utf16("前"), *inline(utf16, body, selector=0x0000), *utf16("後"), 0x0A)   # 0x0000 = テンプレート指示文 (非表示)
    blocks = TextParser(data).parse()
    assert blocks[0].text == "前後"                      # hidden は text に出ない
    hidden = [r for r in blocks[0].runs if r.hidden]
    assert hidden and hidden[0].text == body


def test_unterminated_inline_is_flagged_not_silently_truncated(synth):
    utf16, words = synth.utf16, synth.words
    # 0x1E が無いまま領域が終わる
    data = words(0x1F, *utf16("前"), 0x1C, 0x0001, 0x0007, 0x0000, 0x0000, 0x0001, 0x1D, *utf16("切れた"))
    parser = TextParser(data)
    blocks = parser.parse()
    assert parser.truncated == 1
    last = blocks[0].runs[-1]
    assert last.truncated and last.text == "切れた"
