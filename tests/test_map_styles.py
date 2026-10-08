"""表の罫線スパン対応付け (_map_styles)。旧実装 (総当たり) と同じ結果を返し、列数に対して急増しないこと。"""
import random
import time

import pytest

from jtdlib.model import _Chunk
from jtdlib.parser import TextParser


def _old_map_styles(chunks, ents):
    """0.1.0 の実装。等価性の基準として保持。"""
    if not ents or not chunks:
        return [None] * len(chunks)
    best = (-1, [None] * len(chunks))
    widths = [e[3] for e in ents]
    for gap in (2, 0):
        prefix = [0]
        for w in widths:
            prefix.append(prefix[-1] + w + gap)
        for c in chunks:
            for k in range(len(ents)):
                origin = c.x0 - prefix[k]
                styles, score = [], 0
                for ch in chunks:
                    hit = None
                    for j, e in enumerate(ents):
                        if widths[j] == 0:
                            continue
                        xs = origin + prefix[j]
                        if xs - 1 <= ch.x0 <= xs + widths[j]:
                            hit = e[2]
                            score += 2 + (widths[j] in (ch.x1 - ch.x0, ch.x1 - ch.x0 + 2))
                            break
                    styles.append(hit)
                if score > best[0]:
                    best = (score, styles)
    return best[1]


def _row(n: int, rng: random.Random, gap: int, origin: int, ragged: bool):
    """n 列の物理行: スパン (… , style, width) と、それに乗るセル断片。"""
    ents, chunks, x = [], [], origin
    for j in range(n):
        w = rng.choice([8, 12, 16, 20, 0]) if ragged else 12
        ents.append((0, 0, 1 + j % 5, w))
        if w:
            if ragged and rng.random() < 0.2:      # 断片を幅の途中で切る / 欠ける
                chunks.append(_Chunk(x0=x + 1, x1=x + w // 2, flag=0, para=None))
            else:
                chunks.append(_Chunk(x0=x, x1=x + w, flag=0, para=None))
        x += w + gap
    rng.shuffle(chunks)
    return chunks, ents


@pytest.mark.parametrize("seed", range(40))
def test_matches_bruteforce(seed):
    rng = random.Random(seed)
    n = rng.randint(1, 12)
    chunks, ents = _row(n, rng, rng.choice([0, 2]), rng.randint(0, 40), ragged=True)
    assert TextParser._map_styles(chunks, ents) == _old_map_styles(chunks, ents)


def test_scales_to_many_columns():
    rng = random.Random(0)
    chunks, ents = _row(80, rng, 2, 10, ragged=False)
    t0 = time.perf_counter()
    styles = TextParser._map_styles(chunks, ents)
    dt = time.perf_counter() - t0
    assert all(s is not None for s in styles)
    assert dt < 0.5, f"80 columns took {dt:.2f}s"        # 旧実装は約 5 秒
