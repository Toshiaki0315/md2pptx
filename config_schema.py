"""config.yaml の内容を起動時に検証する

python-pptx まで届いてしまうと、原因とかけ離れたエラーになることが多い。
例えば `size_pt: "20"`（文字列）は
「Exceeds the limit (4300 digits) for integer string conversion」という
設定ミスとは読み取れない例外になる。
またキー名の綴り違いは、そもそもエラーにならず黙って無視される。

ここで変換を始める前に検査し、どこをどう直せばよいかを日本語で示す。

外部ライブラリ（pydantic / jsonschema）は使っていない。
検査項目が限られており、それらの英語の汎用メッセージを日本語に
組み替える手間を考えると、直接書いた方が要点を伝えやすいと判断した。
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from typing import Any

#: 最上位で認識するキー
KNOWN_TOP_LEVEL = ('slides', 'fonts', 'images', 'theme', 'mermaid')

KNOWN_SLIDES = (
    'template_path', 'layout', 'show_slide_number', 'h3_as', 'footer',
    'layouts', 'use_template_fonts',
)
KNOWN_IMAGES = ('default_height_inches', 'position_inches', 'downscale', 'dpi')
KNOWN_THEME = (
    'accent_color', 'text_color', 'code_bg_color',
    'dark_background_color', 'dark_text_color',
)
KNOWN_MERMAID = (
    'renderer', 'endpoint', 'warn_on_external', 'fallback_to_public', 'cli_path',
)

#: フォント設定の項目
KNOWN_FONT_FIELDS = ('name', 'size_pt', 'bold', 'color_rgb')

#: コードが参照するフォント設定のキー（bullet_level_N / ordered_level_N は別途パターンで判定）
KNOWN_FONT_KEYS = (
    'title', 'title_h1', 'title_h2', 'title_h3', 'body',
    'inline_code', 'code_block', 'table_header', 'table_body', 'footer',
)
BULLET_LEVEL_PATTERN = re.compile(r'^(bullet|ordered)_level_\d+$')

#: フッターの項目
KNOWN_FOOTER = ('text', 'date', 'show_on_title')

#: スライドレイアウトの用途
KNOWN_LAYOUT_KINDS = ('title', 'content')

#: 選択肢が決まっている項目
VALID_LAYOUTS = ('16:9', '4:3', '16:10', 'A4')
VALID_H3_AS = ('subheading', 'slide')
VALID_MERMAID_RENDERERS = ('kroki', 'local', 'off')


@dataclass
class ValidationResult:
    """検証結果

    errors があれば変換を中止する。warnings は指摘のみで変換は継続する
    （そのままでも動作はするが、意図どおりでない可能性が高いもの）。
    """

    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.errors


def _describe(value: Any) -> str:
    """メッセージ中に値を示すための表記"""
    return f"現在: {value!r}"


def _suggestion(name: str, candidates: tuple[str, ...]) -> str:
    """綴り違いの可能性がある場合に、候補を示す文言を返す"""
    close = difflib.get_close_matches(name, candidates, n=1, cutoff=0.7)
    return f"（'{close[0]}' の誤りではありませんか？）" if close else ""


def _check_bool(path: str, value: Any, result: ValidationResult) -> None:
    if not isinstance(value, bool):
        result.errors.append(f"{path}: true か false で指定してください（{_describe(value)}）")


def _check_str(path: str, value: Any, result: ValidationResult) -> None:
    if not isinstance(value, str):
        result.errors.append(f"{path}: 文字列で指定してください（{_describe(value)}）")


def _check_positive_number(path: str, value: Any, result: ValidationResult) -> None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        result.errors.append(f"{path}: 数値で指定してください（{_describe(value)}）")
    elif value <= 0:
        result.errors.append(f"{path}: 0より大きい値で指定してください（{_describe(value)}）")


def _check_rgb(path: str, value: Any, result: ValidationResult) -> None:
    """色指定（0〜255の整数を3つ並べたもの）を検査する"""
    if not isinstance(value, (list, tuple)):
        result.errors.append(
            f"{path}: [赤, 緑, 青] の形式で指定してください"
            f"（例: [0, 112, 192]） （{_describe(value)}）"
        )
        return
    if len(value) != 3:
        result.errors.append(
            f"{path}: 赤・緑・青の3つの値が必要です（{len(value)}個指定されています）"
        )
        return
    for index, component in enumerate(value):
        name = ('赤', '緑', '青')[index]
        if isinstance(component, bool) or not isinstance(component, int):
            result.errors.append(
                f"{path}: {name}は整数で指定してください（{_describe(component)}）"
            )
        elif not 0 <= component <= 255:
            result.errors.append(
                f"{path}: {name}は0〜255の範囲で指定してください（{_describe(component)}）"
            )


def _check_mapping(path: str, value: Any, result: ValidationResult) -> bool:
    """マッピングであることを確認する（検査を続けられる場合に True）"""
    if value is None:
        return False
    if not isinstance(value, dict):
        result.errors.append(f"{path}: 設定の入れ子（キー: 値）で指定してください（{_describe(value)}）")
        return False
    return True


def _check_unknown_keys(
    prefix: str, conf: dict[str, Any], known: tuple[str, ...], result: ValidationResult
) -> None:
    for key in conf:
        if key not in known:
            result.warnings.append(
                f"{prefix}{key}: 認識できない設定のため無視されます{_suggestion(key, known)}"
            )


def _validate_slides(conf: Any, result: ValidationResult) -> None:
    if not _check_mapping('slides', conf, result):
        return
    _check_unknown_keys('slides.', conf, KNOWN_SLIDES, result)

    if conf.get('layout') is not None:
        layout = conf['layout']
        if layout not in VALID_LAYOUTS:
            result.warnings.append(
                f"slides.layout: '{layout}' は未対応のため 16:9 が使われます"
                f"（指定できる値: {' / '.join(VALID_LAYOUTS)}）"
            )
    if conf.get('template_path') is not None:
        _check_str('slides.template_path', conf['template_path'], result)
    if conf.get('show_slide_number') is not None:
        _check_bool('slides.show_slide_number', conf['show_slide_number'], result)
    _validate_footer(conf.get('footer'), result)
    _validate_layouts(conf.get('layouts'), result)
    if conf.get('use_template_fonts') is not None:
        _check_bool('slides.use_template_fonts', conf['use_template_fonts'], result)
    if conf.get('h3_as') is not None:
        h3_as = conf['h3_as']
        if h3_as not in VALID_H3_AS:
            result.errors.append(
                f"slides.h3_as: '{h3_as}' は指定できません"
                f"（{' / '.join(VALID_H3_AS)} のいずれか）"
                f"{_suggestion(str(h3_as), VALID_H3_AS)}"
            )


def _validate_layouts(conf: Any, result: ValidationResult) -> None:
    if not _check_mapping('slides.layouts', conf, result):
        return
    _check_unknown_keys('slides.layouts.', conf, KNOWN_LAYOUT_KINDS, result)

    for kind in KNOWN_LAYOUT_KINDS:
        if conf.get(kind) is not None:
            _check_str(f'slides.layouts.{kind}', conf[kind], result)


def _validate_footer(conf: Any, result: ValidationResult) -> None:
    if not _check_mapping('slides.footer', conf, result):
        return
    _check_unknown_keys('slides.footer.', conf, KNOWN_FOOTER, result)

    if conf.get('text') is not None:
        _check_str('slides.footer.text', conf['text'], result)
    if conf.get('show_on_title') is not None:
        _check_bool('slides.footer.show_on_title', conf['show_on_title'], result)

    date = conf.get('date')
    if date is not None and not isinstance(date, (bool, str)):
        result.errors.append(
            "slides.footer.date: true（変換日を表示）か、表示したい文字列で"
            f"指定してください（{_describe(date)}）"
        )

def _validate_fonts(conf: Any, result: ValidationResult) -> None:
    if not _check_mapping('fonts', conf, result):
        return

    for key, font_conf in conf.items():
        path = f"fonts.{key}"
        if key not in KNOWN_FONT_KEYS and not BULLET_LEVEL_PATTERN.match(str(key)):
            result.warnings.append(
                f"{path}: 認識できない設定のため無視されます{_suggestion(str(key), KNOWN_FONT_KEYS)}"
            )
        if not _check_mapping(path, font_conf, result):
            continue

        _check_unknown_keys(f"{path}.", font_conf, KNOWN_FONT_FIELDS, result)
        if font_conf.get('name') is not None:
            _check_str(f"{path}.name", font_conf['name'], result)
        if font_conf.get('size_pt') is not None:
            _check_positive_number(f"{path}.size_pt", font_conf['size_pt'], result)
        if font_conf.get('bold') is not None:
            _check_bool(f"{path}.bold", font_conf['bold'], result)
        if font_conf.get('color_rgb') is not None:
            _check_rgb(f"{path}.color_rgb", font_conf['color_rgb'], result)


def _validate_images(conf: Any, result: ValidationResult) -> None:
    if not _check_mapping('images', conf, result):
        return
    _check_unknown_keys('images.', conf, KNOWN_IMAGES, result)

    if conf.get('default_height_inches') is not None:
        _check_positive_number(
            'images.default_height_inches', conf['default_height_inches'], result
        )
    if conf.get('dpi') is not None:
        _check_positive_number('images.dpi', conf['dpi'], result)
    if conf.get('downscale') is not None:
        _check_bool('images.downscale', conf['downscale'], result)

    position = conf.get('position_inches')
    if position is not None:
        if not isinstance(position, (list, tuple)) or len(position) < 2:
            result.errors.append(
                "images.position_inches: [左からの距離, 上からの距離] の形式で"
                f"指定してください（例: [5.2, 1.8]） （{_describe(position)}）"
            )
        else:
            for index, axis in enumerate(('左からの距離', '上からの距離')):
                value = position[index]
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    result.errors.append(
                        f"images.position_inches: {axis}は数値で指定してください"
                        f"（{_describe(value)}）"
                    )


def _validate_theme(conf: Any, result: ValidationResult) -> None:
    if not _check_mapping('theme', conf, result):
        return
    _check_unknown_keys('theme.', conf, KNOWN_THEME, result)

    for key in KNOWN_THEME:
        if conf.get(key) is not None:
            _check_rgb(f"theme.{key}", conf[key], result)


def _validate_mermaid(conf: Any, result: ValidationResult) -> None:
    if not _check_mapping('mermaid', conf, result):
        return
    _check_unknown_keys('mermaid.', conf, KNOWN_MERMAID, result)

    renderer = conf.get('renderer')
    if renderer is not None and renderer not in VALID_MERMAID_RENDERERS:
        result.errors.append(
            f"mermaid.renderer: '{renderer}' は指定できません"
            f"（{' / '.join(VALID_MERMAID_RENDERERS)} のいずれか）"
            f"{_suggestion(str(renderer), VALID_MERMAID_RENDERERS)}"
        )
    for key in ('endpoint', 'cli_path'):
        if conf.get(key) is not None:
            _check_str(f"mermaid.{key}", conf[key], result)
    for key in ('warn_on_external', 'fallback_to_public'):
        if conf.get(key) is not None:
            _check_bool(f"mermaid.{key}", conf[key], result)


def validate_config(config: Any) -> ValidationResult:
    """設定内容を検証し、エラーと警告をまとめて返す

    1件目で止めず、直すべき箇所をまとめて提示する。
    """
    result = ValidationResult()

    if config is None:
        return result
    if not isinstance(config, dict):
        result.errors.append(
            f"設定ファイルの最上位は「キー: 値」の形式である必要があります（{_describe(config)}）"
        )
        return result

    _check_unknown_keys('', config, KNOWN_TOP_LEVEL, result)

    validators = {
        'slides': _validate_slides,
        'fonts': _validate_fonts,
        'images': _validate_images,
        'theme': _validate_theme,
        'mermaid': _validate_mermaid,
    }
    for key, validate in validators.items():
        if key in config:
            validate(config[key], result)

    return result
