"""フォントファイルから、実際の字幅と行の高さを読み取る

python-pptx は描画を行わないため、はみ出し判定は概算に頼るしかない。
文字幅を「全角=1em / 半角=0.5em」と決め打ちすると、実際の字幅
（Arial の 'i' は 0.22em、'W' は 0.94em）とは大きくずれる。

手元にフォントファイルがある場合は、そこから実測値を読んで概算の精度を上げる。
見つからない場合は None を返し、呼び出し側の既定値に任せる。

**手元にあるかどうか**が前提である点に注意。PowerPoint での描画は別の
環境で行われることがあり、たとえば macOS で Meiryo を指定した場合、
このモジュールはフォントを見つけられず既定値のままになる。
"""

from __future__ import annotations

import glob
import os
import sys
from functools import lru_cache
from typing import Any

try:
    from fontTools.ttLib import TTCollection, TTFont
except ImportError:  # pragma: no cover - fontTools が無い環境でも変換自体は動かす
    TTCollection = TTFont = None  # type: ignore[assignment, misc]

#: 探索するフォントファイルの拡張子
FONT_EXTENSIONS = ('ttf', 'otf', 'ttc', 'otc')

#: name テーブルのうち、フォント名として照合する ID
#: 1=ファミリー名, 4=フルネーム, 16=優先ファミリー名（"Yu Gothic" など）
NAME_IDS = (1, 4, 16)


def font_directories() -> list[str]:
    """OSごとのフォント置き場を返す"""
    home = os.path.expanduser('~')
    if sys.platform == 'darwin':
        return [
            '/System/Library/Fonts',
            '/System/Library/Fonts/Supplemental',
            '/Library/Fonts',
            os.path.join(home, 'Library/Fonts'),
        ]
    if os.name == 'nt':
        windir = os.environ.get('WINDIR', r'C:\Windows')
        local = os.environ.get('LOCALAPPDATA', '')
        directories = [os.path.join(windir, 'Fonts')]
        if local:
            directories.append(os.path.join(local, r'Microsoft\Windows\Fonts'))
        return directories
    return [
        '/usr/share/fonts',
        '/usr/local/share/fonts',
        os.path.join(home, '.fonts'),
        os.path.join(home, '.local/share/fonts'),
    ]


def font_files() -> list[str]:
    """フォント置き場にあるファイルを列挙する（サブフォルダも見る）"""
    paths = []
    for directory in font_directories():
        if not os.path.isdir(directory):
            continue
        for extension in FONT_EXTENSIONS:
            paths += glob.glob(
                os.path.join(directory, '**', f'*.{extension}'), recursive=True
            )
    return paths


def normalize(name: str) -> str:
    """照合用にフォント名をそろえる（空白・記号を落として小文字に）"""
    return ''.join(ch for ch in name.lower() if ch.isalnum())


def _open_fonts(path: str) -> list[Any]:
    """フォントファイルを開く（.ttc は複数のフォントを含む）"""
    if path.lower().endswith(('.ttc', '.otc')):
        return list(TTCollection(path, lazy=True).fonts)
    return [TTFont(path, lazy=True)]


def _names_of(font: Any) -> list[str]:
    """フォントが名乗っている名前を集める"""
    names = []
    for name_id in NAME_IDS:
        try:
            value = font['name'].getDebugName(name_id)
        except Exception:
            value = None
        if value:
            names.append(normalize(value))
    return names


@lru_cache(maxsize=1)
def _font_index() -> dict[str, tuple[str, int]]:
    """フォント名から (ファイル, コレクション内の番号) を引く索引を作る

    全ファイルの name テーブルを読むため、初回だけ時間がかかる（数百ミリ秒）。
    ファイル名からは分からない名前（"Yu Gothic" が YuGothR.ttc など）を
    引けるようにするために必要。
    """
    if TTFont is None:
        return {}

    index: dict[str, tuple[str, int]] = {}
    for path in font_files():
        try:
            fonts = _open_fonts(path)
        except Exception:
            continue  # 壊れたフォントや未対応の形式は飛ばす
        for number, font in enumerate(fonts):
            for name in _names_of(font):
                index.setdefault(name, (path, number))
    return index


class FontMetrics:
    """1つのフォントの実測値

    advance() はフォントサイズに対する比率で返すので、
    ポイント数に直すときはフォントサイズを掛ける。
    """

    def __init__(self, font: Any) -> None:
        self._units_per_em = font['head'].unitsPerEm
        self._cmap = font.getBestCmap()
        self._metrics = font['hmtx'].metrics
        header = font['hhea']
        #: 1行が占める高さの、フォントサイズに対する比率
        self.line_height_ratio = (
            header.ascender - header.descender + header.lineGap
        ) / self._units_per_em
        self._widths: dict[str, float] = {}

    def advance(self, ch: str) -> float | None:
        """1文字ぶんの送り幅。フォントに無い文字は None"""
        if ch in self._widths:
            return self._widths[ch]

        glyph = self._cmap.get(ord(ch))
        if glyph is None or glyph not in self._metrics:
            return None
        width = self._metrics[glyph][0] / self._units_per_em
        self._widths[ch] = width
        return width


@lru_cache(maxsize=32)
def metrics_for(name: str | None) -> FontMetrics | None:
    """フォント名から実測値を得る。手元にフォントが無ければ None"""
    if not name or TTFont is None:
        return None

    found = _font_index().get(normalize(name))
    if found is None:
        return None

    path, number = found
    try:
        return FontMetrics(_open_fonts(path)[number])
    except Exception:
        return None  # 読めないフォントは既定値に任せる
