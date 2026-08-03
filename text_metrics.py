"""テキストを描画したときの分量を概算する

python-pptx はレンダリングエンジンを持たないため、実際の折り返し位置や
描画後の高さを知ることができない。ここでは文字幅（全角/半角）とフォントサイズから
行数と高さを概算し、枠からのはみ出し判定に用いる。

あくまで概算であり、実際のフォントのメトリクスとは一致しない。
はみ出しを「完全に防ぐ」ことではなく「起こりにくくする」ことが目的である。
"""

from __future__ import annotations

import math
import unicodedata
from dataclasses import dataclass

#: 全角として扱う East Asian Width の区分（W=Wide, F=Fullwidth, A=Ambiguous）
FULLWIDTH_CATEGORIES = ('W', 'F', 'A')

#: フォントサイズに対する1文字の幅の比率
FULLWIDTH_RATIO = 1.0
HALFWIDTH_RATIO = 0.5

#: 箇条書き1レベルあたりのインデント（インチ）
INDENT_PER_LEVEL_INCHES = 0.3

#: フォントサイズが取得できない場合の既定値
DEFAULT_FONT_SIZE_PT = 18.0

#: 行送りの既定倍率
DEFAULT_LINE_SPACING = 1.0

#: 縮小率の下限（これ以上小さくすると読めなくなるため）
MIN_SHRINK_SCALE = 0.6

#: 縮小率を探索する刻み幅
SHRINK_STEP = 0.05

POINTS_PER_INCH = 72.0


@dataclass(frozen=True)
class ParagraphMetrics:
    """高さの概算に必要な、1段落ぶんの情報"""

    text: str
    font_size_pt: float
    level: int = 0
    line_spacing: float = DEFAULT_LINE_SPACING
    space_after_pt: float = 0.0


def char_width_ratio(ch: str) -> float:
    """1文字の幅を、フォントサイズに対する比率で返す"""
    if unicodedata.east_asian_width(ch) in FULLWIDTH_CATEGORIES:
        return FULLWIDTH_RATIO
    return HALFWIDTH_RATIO


def estimate_text_width_pt(text: str, font_size_pt: float) -> float:
    """テキストを1行で描画したときの幅（ポイント）を概算する"""
    return sum(char_width_ratio(ch) for ch in text) * font_size_pt


def estimate_line_count(text: str, font_size_pt: float, available_width_pt: float) -> int:
    """折り返しを考慮した行数を概算する（空行も1行として数える）"""
    if available_width_pt <= 0 or not text:
        return 1
    width = estimate_text_width_pt(text, font_size_pt)
    return max(1, math.ceil(width / available_width_pt))


def estimate_height_pt(
    paragraphs: list[ParagraphMetrics], available_width_pt: float, scale: float = 1.0
) -> float:
    """段落群を描画したときに必要な高さ（ポイント）を概算する"""
    total = 0.0
    for paragraph in paragraphs:
        font_size = paragraph.font_size_pt * scale
        indent = paragraph.level * INDENT_PER_LEVEL_INCHES * POINTS_PER_INCH
        # インデントで幅が無くなっても、最低1文字は入る前提で計算する
        width = max(available_width_pt - indent, font_size)

        lines = estimate_line_count(paragraph.text, font_size, width)
        total += lines * font_size * paragraph.line_spacing
        total += paragraph.space_after_pt * scale
    return total


def fit_scale(
    paragraphs: list[ParagraphMetrics],
    available_width_pt: float,
    available_height_pt: float,
    minimum: float = MIN_SHRINK_SCALE,
    step: float = SHRINK_STEP,
) -> float:
    """枠に収まる最大の縮小率（1.0=縮小不要）を返す

    フォントを縮めると行の高さだけでなく折り返し行数も減るため、
    必要な縮小率は単純な比では求まらない。段階的に試して収まる値を採用する。
    """
    if available_height_pt <= 0 or not paragraphs:
        return 1.0

    scale = 1.0
    while scale > minimum:
        if estimate_height_pt(paragraphs, available_width_pt, scale) <= available_height_pt:
            return scale
        scale = round(scale - step, 4)
    return minimum
