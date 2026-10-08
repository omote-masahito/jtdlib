"""DocumentText / LayoutBoxText / Footnote ストリーム (SsmgV.01 / TextV.01) の制御コード。

根拠: OpenJTD RFC 0001/0003/0005/0009 と、一太郎2023 で生成したコーパスの観測 (RFC 0010 草稿)。
"""

# 本文ストリームの制御コード (UTF-16BE の 1 ワード)
TEXT_RUN = 0x1F          # テキストラン開始
INLINE_START = 0x1D      # インラインレコードの本文開始
INLINE_END = 0x1E        # インラインレコードの本文終了
ROW_END = 0x0E           # 表の物理行終端
PARA_END = 0x0A          # 段落終端
PAGE_BREAK = 0x0C        # 改ページ
ENTRY_END = 0x0D         # 脚注エントリ終端 (Footnote ストリームのみ)
TAB = 0x09
RECORD = 0x1C            # レコード開始: 001c [class] [len] ...

# レコードクラス
REC_CONTEXT = 0x0000     # コンテキストマーカー (脚注マーカーの位置など)
REC_INLINE = 0x0001      # インラインセグメント
REC_LINE = 0x0010        # 段落属性 / 表の物理行スペック (サブエントリ列)
REC_SECTION = 0x0020     # 表セクション遷移
REC_CELL = 0x0030        # 表セル断片 (12 ワード固定)

# 0x0010 レコードのサブエントリタグ
SUB_INDENT = 0x0026      # [_, left, _, first, _] インデント (単位: カラム)
SUB_ROWSPEC = 0x008F     # [width, 0, fmt, spans..., ffff] 表の物理行スペック

# 0x0001 インラインレコードのセレクタ
SEL_TEMPLATE_INSTRUCTION = 0x0000
SEL_AUTO_TEXT = 0x0001   # 自動生成テキスト (脚注番号 "*1" など。本文に出す)
SEL_RUBY_BASE = 0x0003   # 装飾付き本文 (直後に 0x0082 が続けばルビ親字)
SEL_RUBY_TEXT = 0x0082   # ルビ文字列

# 座標単位: 表のセル座標 (0x0030 の x0/x1) と行スペックの幅は 1 文字幅の 1/4
COORD_UNITS_PER_CHAR = 4

# 一太郎の「テキスト保存 (罫線内文字列=セル単位)」がセル内改行に使う私用文字
CELL_NEWLINE = "\ue000"
