"""サロゲートペア (U+10000 以上) の復元。合成した TextV.01 相当の語列で検査する。"""
from jtdlib.parser import TextParser, decode_words


def test_decode_words_joins_surrogate_pairs(synth):
    utf16 = synth.utf16
    assert decode_words(utf16("𠮷野家")) == "𠮷野家"
    assert decode_words([0xD842]) == "\ufffd"          # 孤立サロゲートは置換


def test_body_text_keeps_astral_characters(synth):
    utf16, words = synth.utf16, synth.words
    # 0x1F ラン開始 … 0x0A 段落終端 (領域ヘッダなし → region() の fallback を使う)
    data = words(0x1F, *utf16("𠮷野家"), 0x09, *utf16("😀"), 0x0A)
    blocks = TextParser(data).parse()
    text = blocks[0].text
    assert text == "𠮷野家\t😀"
    text.encode("utf-8")                              # 孤立サロゲートなら失敗する


def test_inline_text_keeps_astral_characters(synth):
    utf16, words = synth.utf16, synth.words
    # 001c 0001 0007 0000 000x [sel] 001d <text> 001e … 001f
    body = utf16("𠮷")
    data = words(0x1F, *utf16("前"),
                 0x1C, 0x0001, 0x0007, 0x0000, 0x0000, 0x0001, 0x1D, *body, 0x1E, 0x1F,
                 *utf16("後"), 0x0A)
    blocks = TextParser(data).parse()
    assert "𠮷" in blocks[0].text
    blocks[0].text.encode("utf-8")
