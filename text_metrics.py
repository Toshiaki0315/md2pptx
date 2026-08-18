"""テキストを描画したときの分量を概算する

python-pptx はレンダリングエンジンを持たないため、実際の折り返し位置や
描画後の高さを知ることができない。ここではフォントサイズと文字幅から
行数と高さを概算し、枠からのはみ出し判定に用いる。

文字幅と行の高さは、指定されたフォントが手元にあれば font_metrics が読む
実測値を使い、無ければ全角/半角からの概算に落とす。

いずれにせよ概算であり、はみ出しを「完全に防ぐ」ことではなく
「起こりにくくする」ことが目的である。
"""

from __future__ import annotations

import math
import unicodedata
from collections.abc import Callable
from dataclasses import dataclass

from font_metrics import FontMetrics, metrics_for

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

#: 1行が占める高さの、フォントサイズに対する比率（実測値が使えない場合の既定）
#
# PowerPoint の行の高さはフォントサイズそのものではなく、フォントが持つ
# 行高（ascent + descent + line gap）に基づく。本ツールが主に使う和文フォント
# （Meiryo / Yu Gothic 等）はこれが約 1.3em あるため、フォントサイズと同一と
# みなすと 3 割ほど過小評価になる。
# フォントファイルが手元にある場合は font_metrics が実測値を返すため、
# この値は使われない。
LINE_HEIGHT_RATIO = 1.3

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
    space_before_pt: float = 0.0
    font_name: str | None = None


def char_width_ratio(ch: str, metrics: FontMetrics | None = None) -> float:
    """1文字の幅を、フォントサイズに対する比率で返す

    フォントの実測値が使える場合はそれを優先する。フォントにその文字が
    無い場合（欧文フォントに和文を書いた場合など）は全角/半角から概算する。
    """
    if metrics is not None:
        measured = metrics.advance(ch)
        if measured is not None:
            return measured
    if unicodedata.east_asian_width(ch) in FULLWIDTH_CATEGORIES:
        return FULLWIDTH_RATIO
    return HALFWIDTH_RATIO


def line_height_ratio(font_name: str | None = None) -> float:
    """1行が占める高さの、フォントサイズに対する比率"""
    metrics = metrics_for(font_name)
    return metrics.line_height_ratio if metrics is not None else LINE_HEIGHT_RATIO


def estimate_text_width_pt(
    text: str, font_size_pt: float, font_name: str | None = None
) -> float:
    """テキストを1行で描画したときの幅（ポイント）を概算する"""
    metrics = metrics_for(font_name)
    return sum(char_width_ratio(ch, metrics) for ch in text) * font_size_pt


def estimate_line_count(
    text: str, font_size_pt: float, available_width_pt: float,
    font_name: str | None = None,
) -> int:
    """折り返しを考慮した行数を概算する（空行も1行として数える）"""
    if available_width_pt <= 0 or not text:
        return 1
    width = estimate_text_width_pt(text, font_size_pt, font_name)
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

        lines = estimate_line_count(paragraph.text, font_size, width, paragraph.font_name)
        total += (lines * font_size * line_height_ratio(paragraph.font_name)
                  * paragraph.line_spacing)
        total += (paragraph.space_before_pt + paragraph.space_after_pt) * scale
    return total


def _search_scale(
    measure: Callable[[float], float],
    available_height_pt: float,
    minimum: float,
    step: float,
) -> float:
    """measure(縮小率) が枠に収まる最大の縮小率を段階的に探す

    フォントを縮めると行の高さだけでなく折り返し行数も減るため、
    必要な縮小率は単純な比では求まらない。
    """
    if available_height_pt <= 0:
        return 1.0

    scale = 1.0
    while scale > minimum:
        if measure(scale) <= available_height_pt:
            return scale
        scale = round(scale - step, 4)
    return minimum


def fit_scale(
    paragraphs: list[ParagraphMetrics],
    available_width_pt: float,
    available_height_pt: float,
    minimum: float = MIN_SHRINK_SCALE,
    step: float = SHRINK_STEP,
) -> float:
    """枠に収まる最大の縮小率（1.0=縮小不要）を返す"""
    if not paragraphs:
        return 1.0
    return _search_scale(
        lambda scale: estimate_height_pt(paragraphs, available_width_pt, scale),
        available_height_pt,
        minimum,
        step,
    )


@dataclass(frozen=True)
class TableRowMetrics:
    """高さの概算に必要な、表1行ぶんの情報"""

    texts: list[str]
    font_size_pt: float
    font_name: str | None = None


def estimate_row_heights_pt(
    rows: list[TableRowMetrics],
    column_width_pt: float,
    vertical_margin_pt: float,
    scale: float = 1.0,
) -> list[float]:
    """表の各行に必要な高さ（ポイント）を概算する

    セルの内容が折り返す場合はその行数ぶん高くなる。
    """
    heights = []
    for row in rows:
        font_size = row.font_size_pt * scale
        lines = max(
            (estimate_line_count(text, font_size, column_width_pt, row.font_name)
             for text in row.texts),
            default=1,
        )
        heights.append(
            lines * font_size * line_height_ratio(row.font_name) + vertical_margin_pt
        )
    return heights


def estimate_table_height_pt(
    rows: list[TableRowMetrics],
    column_width_pt: float,
    vertical_margin_pt: float,
    scale: float = 1.0,
) -> float:
    """表全体の高さ（ポイント）を概算する"""
    return sum(estimate_row_heights_pt(rows, column_width_pt, vertical_margin_pt, scale))


def paginate_row_heights(
    row_heights: list[float],
    available_height_pt: float,
    repeat_first_row: bool = False,
) -> list[list[int]]:
    """表の行を、枠に収まるページ単位のインデックス列に分割する

    repeat_first_row が真の場合、2ページ目以降の先頭に1行目（見出し行）を繰り返し、
    その高さも各ページの消費として数える。
    1行が単独で枠を超える場合でも、そのページには必ず1行は載せる（無限分割の防止）。
    """
    if not row_heights:
        return []

    data_start = 1 if repeat_first_row else 0
    if available_height_pt <= 0 or len(row_heights) <= data_start:
        return [list(range(len(row_heights)))]

    header_height = row_heights[0] if repeat_first_row else 0.0
    pages: list[list[int]] = []
    current: list[int] = []
    used = header_height

    for index in range(data_start, len(row_heights)):
        height = row_heights[index]
        if current and used + height > available_height_pt:
            pages.append(current)
            current = []
            used = header_height
        current.append(index)
        used += height

    if current:
        pages.append(current)

    if repeat_first_row:
        return [[0] + page for page in pages]
    return pages


def fit_table_scale(
    rows: list[TableRowMetrics],
    column_width_pt: float,
    vertical_margin_pt: float,
    available_height_pt: float,
    minimum: float = MIN_SHRINK_SCALE,
    step: float = SHRINK_STEP,
) -> float:
    """表が枠に収まる最大の縮小率（1.0=縮小不要）を返す"""
    if not rows:
        return 1.0
    return _search_scale(
        lambda scale: estimate_table_height_pt(rows, column_width_pt, vertical_margin_pt, scale),
        available_height_pt,
        minimum,
        step,
    )
