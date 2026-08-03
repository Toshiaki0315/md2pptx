"""描画・レイアウトのヘルパー関数群"""

from __future__ import annotations

import collections
import collections.abc
import os
from io import BytesIO
from typing import IO, TYPE_CHECKING, Any, Union

from PIL import Image
from pptx.util import Emu, Inches, Length, Pt
from pptx.dml.color import RGBColor
from bs4 import NavigableString, Tag
from pygments import lex
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.styles import get_style_by_name

if TYPE_CHECKING:
    from pptx.shapes.picture import Picture
    from pptx.slide import Slide
    from pptx.text.text import _Paragraph, _Run

    from generator import PPTXGenerator

#: config.yaml のフォント設定（name / size_pt / bold / color_rgb）
FontConfig = dict[str, Any]

#: 画像データとして受け付ける型（ローカルパス、またはメモリ上のバイト列）
ImageSource = Union[str, IO[bytes]]

#: 画像を再サンプリングする際の既定解像度（PowerPointの「図の圧縮」の印刷品質相当）
DEFAULT_IMAGE_DPI = 150

#: JPEGを再エンコードする際の品質
JPEG_QUALITY = 90


def hex_to_rgb(hex_str: str | None) -> RGBColor | None:
    if not hex_str: return None
    hex_str = hex_str.lstrip('#')
    return RGBColor(*tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

def apply_syntax_highlight(
    p: _Paragraph, text: str, language: str | None, font_conf: FontConfig | None
) -> None:
    try:
        if language:
            lexer = get_lexer_by_name(language, stripall=False)
        else:
            lexer = guess_lexer(text)
    except Exception:
        # 未知の言語や判定不能なコードはハイライトなしのプレーンテキストとして扱う
        lexer = get_lexer_by_name('text')

    style = get_style_by_name('monokai') # 濃い背景に合うmonokaiを使用

    for token, content in lex(text, lexer):
        if not content: continue
        run = p.add_run()
        run.text = content
        apply_font_style(run, font_conf)

        token_style = style.style_for_token(token)
        if token_style['color']:
            run.font.color.rgb = hex_to_rgb(token_style['color'])
        else:
            run.font.color.rgb = RGBColor(248, 248, 242) # 背景に合う白をデフォルトに

        if token_style['bold']: run.font.bold = True
        if token_style['italic']: run.font.italic = True

def apply_font_style(run: _Run, font_config: FontConfig | None) -> None:
    """フォントスタイルの適用"""
    if not font_config: return
    font = run.font
    if 'name' in font_config: font.name = font_config['name']
    if 'size_pt' in font_config: font.size = Pt(font_config['size_pt'])
    if 'bold' in font_config: font.bold = font_config['bold']
    if 'color_rgb' in font_config:
        rgb = font_config['color_rgb']
        font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

def _rewind(img_data: ImageSource) -> ImageSource:
    """ファイルライクな画像データを先頭に巻き戻して返す"""
    if hasattr(img_data, 'seek'):
        img_data.seek(0)
    return img_data

def _source_size(img_data: ImageSource) -> int | None:
    """元画像のバイト数を返す（取得できない場合は None）"""
    try:
        if isinstance(img_data, str):
            return os.path.getsize(img_data)
        position = img_data.tell()
        img_data.seek(0, os.SEEK_END)
        size = img_data.tell()
        img_data.seek(position)
        return size
    except Exception:
        return None

def downscale_image(
    img_data: ImageSource,
    max_width: Length,
    max_height: Length,
    dpi: int = DEFAULT_IMAGE_DPI,
) -> ImageSource:
    """表示サイズに対して過大な画像を再サンプリングし、埋め込みサイズを削減する

    スライド上の表示サイズと指定DPIから必要な画素数を求め、それを超える分だけ縮小する。
    拡大は行わない（元が小さい画像はそのまま返す）。

    再エンコードでかえってサイズが増える画像（色数の少ない図版など、元のPNGが
    よく圧縮されているケース）では元データをそのまま使う。
    画像として読めない・保存できない場合も同様に元データを返し、変換自体は止めない。
    """
    target_w = max(1, int(Emu(int(max_width)).inches * dpi))
    target_h = max(1, int(Emu(int(max_height)).inches * dpi))

    try:
        _rewind(img_data)

        with Image.open(img_data) as img:
            # 既に十分小さい画像は再エンコードしない（無駄な劣化とCPUを避ける）
            if img.width <= target_w and img.height <= target_h:
                return _rewind(img_data)

            image_format = img.format or 'PNG'
            resized = img.copy()
            resized.thumbnail((target_w, target_h), Image.Resampling.LANCZOS)

            buffer = BytesIO()
            save_options: dict[str, Any] = {'dpi': (dpi, dpi)}
            if image_format == 'JPEG':
                save_options['quality'] = JPEG_QUALITY
            elif image_format == 'PNG':
                save_options['optimize'] = True
            resized.save(buffer, format=image_format, **save_options)

        original_size = _source_size(img_data)
        if original_size is not None and buffer.getbuffer().nbytes >= original_size:
            return _rewind(img_data)

        buffer.seek(0)
        return buffer
    except Exception as e:
        print(f"Warning: 画像の縮小に失敗したため元の画像を使用します: {e}")
        return _rewind(img_data)

def insert_image_fit(
    slide: Slide,
    img_data: ImageSource,
    left: Length,
    top: Length,
    max_width: Length,
    max_height: Length,
) -> Picture:
    """画像を最大枠に収まるようにアスペクト比を保って自動縮小・中央配置する"""
    pic = slide.shapes.add_picture(img_data, left, top)
    ratio_w = max_width / pic.width
    ratio_h = max_height / pic.height
    ratio = min(ratio_w, ratio_h)
    ratio = min(ratio, 1.5) # 極端な拡大を防止

    pic.width = Emu(int(pic.width * ratio))
    pic.height = Emu(int(pic.height * ratio))
    pic.left = Emu(int(left + (max_width - pic.width) / 2))
    pic.top = Emu(int(top + (max_height - pic.height) / 2))
    return pic

def add_runs_from_tag(
    generator: PPTXGenerator,
    element: Tag,
    paragraph: _Paragraph,
    default_font_conf: FontConfig | None,
) -> None:
    """インライン装飾を解釈しながらテキストを追加（再帰処理）"""
    for child in element:
        if isinstance(child, NavigableString):
            text = str(child).replace('\n', ' ')
            if text.strip() or text == ' ':
                run = paragraph.add_run()
                run.text = text
                apply_font_style(run, default_font_conf)
        elif isinstance(child, Tag):
            if child.name in ['ul', 'ol', 'pre', 'img', 'table', 'blockquote']: continue
            if child.name in ['p', 'div', 'span', 'li', 'th', 'td']:
                add_runs_from_tag(generator, child, paragraph, default_font_conf)
            else:
                run = paragraph.add_run()
                run.text = child.get_text().replace('\n', ' ')
                apply_font_style(run, default_font_conf)

                if child.name in ['strong', 'b']: run.font.bold = True
                elif child.name in ['em', 'i']: run.font.italic = True
                elif child.name == 'code':
                    run.font.name = generator.fonts_conf.get('inline_code', {}).get('name', 'Consolas')
                    rgb = generator.fonts_conf.get('inline_code', {}).get('color_rgb', [220, 20, 60])
                    run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

def shrink_body_shape(
    generator: PPTXGenerator,
    width: Length,
    max_height: Length | None = None,
) -> None:
    """テキスト枠を指定サイズ（EMU）に縮める（レイアウト調整用ヘルパー）

    注意: 下の left/top の自己代入は削除しないこと。プレースホルダーの位置・サイズは
    スライドレイアウトからの継承値であり、一部だけを書き換えると継承が切れて
    残りの値が 0 になる。継承値を明示的に書き戻してから変更する必要がある。
    """
    try:
        body_shape = generator.current_slide.placeholders[1]
        body_shape.left, body_shape.top = body_shape.left, body_shape.top
        body_shape.width = width
        if max_height:
            body_shape.height = max_height
    except Exception:
        pass

def append_text_block(
    generator: PPTXGenerator,
    content: Tag,
    level: int = 0,
    font_conf: FontConfig | None = None,
) -> None:
    """段落オブジェクトを追加し、テキストまたはタグ構造を書き込むヘルパー"""
    if not generator.slide_has_text and len(generator.current_body.paragraphs) == 1 and not generator.current_body.paragraphs[0].text:
        p = generator.current_body.paragraphs[0]
    else:
        p = generator.current_body.add_paragraph()

    p.level = level
    p.space_after = Pt(12)  # 段落後の余白を追加してレイアウトを美しく
    p.line_spacing = 1.2    # 行間を1.2倍に設定

    add_runs_from_tag(generator, content, p, font_conf)

def append_code_textbox(
    generator: PPTXGenerator, content: str, language: str | None = None
) -> None:
    """独立したテキストボックスを作成し、背景色付きでコードを挿入する"""
    layout = generator.layout
    if generator.slide_has_text or generator.forced_layout == '2-column':
        shrink_body_shape(generator, layout.code_split_body_width)
        box_left = layout.code_split_left
        box_top = layout.content_top
        box_width = layout.code_split_width
        box_height = layout.content_height
    elif generator.forced_layout == 'center':
        box_left = layout.code_center_left
        box_top = layout.content_top
        box_width = layout.code_center_width
        box_height = layout.content_height
    else:
        box_left = layout.content_left
        box_top = layout.code_full_top
        box_width = layout.content_width
        box_height = layout.code_full_height

    textbox = generator.current_slide.shapes.add_textbox(box_left, box_top, box_width, box_height)
    textbox.fill.solid()

    bg_color = generator.config.get('theme', {}).get('code_bg_color', [40, 44, 52]) if hasattr(generator, 'config') else [40, 44, 52]
    textbox.fill.fore_color.rgb = RGBColor(*bg_color)

    tf = textbox.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_bottom = Inches(0.2)

    p = tf.paragraphs[0]
    p.line_spacing = 1.1

    conf = generator.fonts_conf.get('code_block', {'name': 'Consolas', 'size_pt': 12})
    apply_syntax_highlight(p, content, language, conf)

    generator.slide_has_text = True

def auto_shrink_text(slide: Slide | None) -> None:
    """スライド内のテキスト行数が多い場合、フォントサイズと余白を自動で縮小してはみ出しを防ぐ"""
    if not slide: return
    try:
        if len(slide.placeholders) > 1:
            body = slide.placeholders[1]
            if not body.has_text_frame: return

            tf = body.text_frame
            line_count = sum(1 for p in tf.paragraphs if p.text.strip())

            if line_count > 6: # 6行を超えたら縮小を開始（より早めに設定）
                shrink_factor = max(0.6, 1.0 - (line_count - 6) * 0.06) # 縮小率を強化
                for p in tf.paragraphs:
                    if p.space_after:
                        p.space_after = Pt(p.space_after.pt * shrink_factor)
                    for run in p.runs:
                        if run.font.size:
                            run.font.size = Pt(int(run.font.size.pt * shrink_factor))
    except Exception:
        pass
