"""GUI（gui.py）の設定値と、config.yaml 相当の辞書との相互変換

画面部品（tkinter）から切り離してあるため、GUIを起動しなくても
設定の組み立て・読み込みを検証できる。
"""

from __future__ import annotations

import os.path
from dataclasses import dataclass, field
from typing import Any

#: 色指定（赤・緑・青）
RGB = tuple[int, int, int]

#: 画面で選べるスライドの画角（config_schema.VALID_LAYOUTS と対応）
ASPECTS = ('16:9', '4:3', '16:10', 'A4')

#: ### の扱い（表示名, 設定値）
H3_CHOICES = (('スライド内の小見出しにする', 'subheading'), ('新しいスライドを作る', 'slide'))

#: Mermaid図の生成方法（表示名, 設定値）
MERMAID_CHOICES = (
    ('Kroki（外部サービスへ図を送信）', 'kroki'),
    ('mermaid-cli（オフライン生成）', 'local'),
    ('生成しない', 'off'),
)

#: 画面でフォントを設定する要素（設定キー, 表示名）
FONT_LABELS = (
    ('title_h1', 'タイトル（#）'),
    ('title_h2', '見出し（##）'),
    ('body', '本文'),
    ('bullet_level_1', '箇条書き'),
    ('code_block', 'コード'),
)


@dataclass
class FontSetting:
    """1要素分のフォント設定（空欄は「指定しない」として扱う）"""

    name: str = ''
    size_pt: float = 0.0

    def to_config(self) -> dict[str, Any]:
        conf: dict[str, Any] = {}
        if self.name.strip():
            conf['name'] = self.name.strip()
        if self.size_pt > 0:
            conf['size_pt'] = self.size_pt
        return conf


def default_fonts() -> dict[str, FontSetting]:
    """config.yaml の既定値に合わせたフォント設定"""
    return {
        'title_h1': FontSetting('Meiryo', 44),
        'title_h2': FontSetting('Meiryo', 32),
        'body': FontSetting('Meiryo', 18),
        'bullet_level_1': FontSetting('Yu Gothic', 18),
        'code_block': FontSetting('Consolas', 12),
    }


@dataclass
class GuiSettings:
    """GUIの入力内容

    基本設定（入力・テンプレート・出力）と、詳細設定をまとめて保持する。
    既定値は config.yaml の内容に合わせてある。
    """

    # --- 基本設定 ---
    input_path: str = ''
    use_template: bool = False
    template_path: str = ''
    output_path: str = 'output.pptx'

    # --- 詳細設定: スライド ---
    aspect: str = '16:9'
    h3_as: str = 'subheading'
    show_slide_number: bool = True
    use_template_fonts: bool = False
    layout_title: str = ''
    layout_content: str = ''

    # --- 詳細設定: フッター ---
    footer_text: str = ''
    footer_date: bool = False
    footer_on_title: bool = False

    # --- 詳細設定: 配色 ---
    accent_color: RGB = (0, 112, 192)
    text_color: RGB = (50, 50, 50)
    code_bg_color: RGB = (40, 44, 52)

    # --- 詳細設定: フォント ---
    fonts: dict[str, FontSetting] = field(default_factory=default_fonts)

    # --- 詳細設定: Mermaid ---
    mermaid_renderer: str = 'kroki'
    mermaid_endpoint: str = 'https://kroki.io'

    # --- 詳細設定: 画像 ---
    image_height_inches: float = 3.5
    image_downscale: bool = True
    image_dpi: float = 150


def build_config(settings: GuiSettings) -> dict[str, Any]:
    """GUIの入力内容から config.yaml 相当の辞書を組み立てる

    指定されていない項目は書き出さず、md2pptx 側の既定値に任せる。
    """
    slides: dict[str, Any] = {
        'h3_as': settings.h3_as,
        'show_slide_number': settings.show_slide_number,
    }

    if settings.use_template and settings.template_path.strip():
        slides['template_path'] = settings.template_path.strip()
        slides['use_template_fonts'] = settings.use_template_fonts
        layouts = {
            kind: name.strip()
            for kind, name in (('title', settings.layout_title), ('content', settings.layout_content))
            if name.strip()
        }
        if layouts:
            slides['layouts'] = layouts
    else:
        # テンプレート指定時は無視される設定なので、使わないときだけ書き出す
        slides['layout'] = settings.aspect

    footer: dict[str, Any] = {}
    if settings.footer_text.strip():
        footer['text'] = settings.footer_text.strip()
    if settings.footer_date:
        footer['date'] = True
    if footer:
        footer['show_on_title'] = settings.footer_on_title
        slides['footer'] = footer

    fonts: dict[str, Any] = {}
    # テンプレートのフォントを使う場合、フォント指定は md2pptx 側で無視される
    if not (settings.use_template and settings.use_template_fonts):
        for key, font in settings.fonts.items():
            conf = font.to_config()
            if conf:
                fonts[key] = conf
        # インラインコードもコードブロックと同じ書体に揃える
        code_name = settings.fonts.get('code_block', FontSetting()).name.strip()
        if code_name:
            fonts['inline_code'] = {'name': code_name}

    mermaid: dict[str, Any] = {'renderer': settings.mermaid_renderer}
    if settings.mermaid_renderer == 'kroki' and settings.mermaid_endpoint.strip():
        mermaid['endpoint'] = settings.mermaid_endpoint.strip()

    config: dict[str, Any] = {
        'slides': slides,
        'theme': {
            'accent_color': list(settings.accent_color),
            'text_color': list(settings.text_color),
            'code_bg_color': list(settings.code_bg_color),
        },
        'images': {
            'default_height_inches': settings.image_height_inches,
            'downscale': settings.image_downscale,
            'dpi': settings.image_dpi,
        },
        'mermaid': mermaid,
    }
    if fonts:
        config['fonts'] = fonts
    return config


def _as_str(value: Any, default: str = '') -> str:
    return value if isinstance(value, str) else default


def _as_bool(value: Any, default: bool) -> bool:
    return value if isinstance(value, bool) else default


def _as_float(value: Any, default: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return default
    return float(value)


def _as_rgb(value: Any, default: RGB) -> RGB:
    if not isinstance(value, (list, tuple)) or len(value) != 3:
        return default
    if not all(isinstance(c, int) and not isinstance(c, bool) and 0 <= c <= 255 for c in value):
        return default
    return (int(value[0]), int(value[1]), int(value[2]))


def _as_choice(value: Any, choices: tuple[str, ...], default: str) -> str:
    return value if value in choices else default


def _mapping(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def settings_from_config(config: Any, base: GuiSettings | None = None) -> GuiSettings:
    """config.yaml の内容をGUIの入力内容に読み込む

    型が想定と違う項目・未対応の項目は既定値のままにし、読み込み自体は失敗させない
    （設定の妥当性は変換実行時に config_schema が検査する）。
    """
    settings = base or GuiSettings()
    conf = _mapping(config)

    slides = _mapping(conf.get('slides'))
    template_path = _as_str(slides.get('template_path'))
    settings.template_path = template_path
    settings.use_template = bool(template_path)
    settings.aspect = _as_choice(slides.get('layout'), ASPECTS, settings.aspect)
    settings.h3_as = _as_choice(
        slides.get('h3_as'), tuple(value for _, value in H3_CHOICES), settings.h3_as
    )
    settings.show_slide_number = _as_bool(slides.get('show_slide_number'), settings.show_slide_number)
    settings.use_template_fonts = _as_bool(slides.get('use_template_fonts'), settings.use_template_fonts)

    layouts = _mapping(slides.get('layouts'))
    settings.layout_title = _as_str(layouts.get('title'), settings.layout_title)
    settings.layout_content = _as_str(layouts.get('content'), settings.layout_content)

    footer = _mapping(slides.get('footer'))
    settings.footer_text = _as_str(footer.get('text'), settings.footer_text)
    date = footer.get('date')
    # 文字列指定（固定の日付表記）もチェックONとして扱う
    settings.footer_date = bool(date) if isinstance(date, (bool, str)) else settings.footer_date
    settings.footer_on_title = _as_bool(footer.get('show_on_title'), settings.footer_on_title)

    theme = _mapping(conf.get('theme'))
    settings.accent_color = _as_rgb(theme.get('accent_color'), settings.accent_color)
    settings.text_color = _as_rgb(theme.get('text_color'), settings.text_color)
    settings.code_bg_color = _as_rgb(theme.get('code_bg_color'), settings.code_bg_color)

    fonts = _mapping(conf.get('fonts'))
    for key, _label in FONT_LABELS:
        font_conf = _mapping(fonts.get(key))
        if not font_conf:
            continue
        current = settings.fonts.get(key, FontSetting())
        settings.fonts[key] = FontSetting(
            name=_as_str(font_conf.get('name'), current.name),
            size_pt=_as_float(font_conf.get('size_pt'), current.size_pt),
        )

    mermaid = _mapping(conf.get('mermaid'))
    settings.mermaid_renderer = _as_choice(
        mermaid.get('renderer'), tuple(value for _, value in MERMAID_CHOICES), settings.mermaid_renderer
    )
    settings.mermaid_endpoint = _as_str(mermaid.get('endpoint'), settings.mermaid_endpoint)

    images = _mapping(conf.get('images'))
    settings.image_height_inches = _as_float(
        images.get('default_height_inches'), settings.image_height_inches
    )
    settings.image_downscale = _as_bool(images.get('downscale'), settings.image_downscale)
    settings.image_dpi = _as_float(images.get('dpi'), settings.image_dpi)

    return settings


def default_output_path(input_path: str) -> str:
    """Markdownのファイル名から出力先の既定値を決める（input.md → input.pptx）"""
    if not input_path:
        return 'output.pptx'
    base = input_path.rsplit('.', 1)[0] if '.' in input_path.rsplit('/', 1)[-1] else input_path
    return f"{base}.pptx"


def validate_inputs(settings: GuiSettings) -> list[str]:
    """変換を始める前に、入力内容の不備を日本語で列挙する"""
    errors: list[str] = []
    if not settings.input_path.strip():
        errors.append('Markdownファイルを選んでください。')
    elif not os.path.isfile(settings.input_path):
        errors.append(f"Markdownファイル '{settings.input_path}' が見つかりません。")

    if settings.use_template:
        if not settings.template_path.strip():
            errors.append('テンプレートファイルを選んでください。')
        elif not os.path.isfile(settings.template_path):
            errors.append(f"テンプレートファイル '{settings.template_path}' が見つかりません。")

    if not settings.output_path.strip():
        errors.append('出力ファイル名を入力してください。')
    elif not settings.output_path.lower().endswith('.pptx'):
        errors.append('出力ファイル名は .pptx で終わるようにしてください。')

    return errors
