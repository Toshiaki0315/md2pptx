"""Markdown（HTML）の各タグをスライド上の要素へ変換する処理"""

from __future__ import annotations

import requests
from io import BytesIO
from typing import TYPE_CHECKING, cast

from bs4 import Tag
from pptx.util import Inches, Length
from pptx.dml.color import RGBColor

from mermaid_renderer import mermaid_conf, render_mermaid
from utils import (
    DEFAULT_IMAGE_DPI,
    FontConfig,
    ImageSource,
    apply_font_style,
    downscale_image,
    insert_image_fit,
    shrink_body_shape,
    add_runs_from_tag,
    append_text_block,
    create_code_textbox,
    auto_shrink_text
)

if TYPE_CHECKING:
    from generator import PPTXGenerator

# 画像取得（HTTP）のタイムアウト秒数
HTTP_TIMEOUT_SEC = 15

# スピーカーノートで1つのまとまりとして扱うブロック要素
NOTE_BLOCK_TAGS = ['p', 'li']

# 箇条書きのインデントレベルの上限（PowerPointの仕様）
MAX_BULLET_LEVEL = 8

# config.yaml に該当設定が無い場合のコードブロックの既定値
DEFAULT_CODE_BLOCK_FONT = {'name': 'Consolas', 'size_pt': 12}
DEFAULT_CODE_BG_COLOR = [40, 44, 52]

def process_heading(generator: PPTXGenerator, tag: Tag) -> None:
    """見出しタグの処理とスライド作成"""
    if generator.current_slide:
        auto_shrink_text(generator.current_slide)
        
    layout_idx = 0 if tag.name == 'h1' else 1
    generator.current_slide = generator.prs.slides.add_slide(generator.prs.slide_layouts[layout_idx])
    
    from pptx.enum.text import MSO_AUTO_SIZE
    title_shape = generator.current_slide.shapes.title
    if title_shape:
        title_shape.text_frame.word_wrap = True
        title_shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        title_shape.text = tag.get_text()
        
        style_key = 'title_h1' if tag.name == 'h1' else 'title_h2'
        for run in title_shape.text_frame.paragraphs[0].runs:
            apply_font_style(run, generator.fonts_conf.get(style_key, generator.fonts_conf.get('title')))

    generator.current_body = generator.current_slide.placeholders[1].text_frame
    generator.current_body.word_wrap = True
    generator.current_body.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    generator.current_body.text = "" 
    generator.slide_has_text = False

    # デフォルト枠のはみ出し補正
    if not generator.slides_conf.get('template_path'):
        try:
            body_shape = generator.current_slide.placeholders[1]
            o_left, o_top, o_width = body_shape.left, body_shape.top, body_shape.width
            new_height = generator.layout.body_height_for(o_top)
            body_shape.left, body_shape.top, body_shape.width, body_shape.height = o_left, o_top, o_width, new_height
        except Exception:
            pass

def process_h3(generator: PPTXGenerator, tag: Tag) -> None:
    """H3見出し（スライド内セクション区切り）の処理"""
    if not tag.get_text(strip=True): return
    from pptx.util import Pt
    
    p = generator.current_body.add_paragraph()
    # 箇条書きを完全に無効化
    from pptx.oxml.xmlchemy import OxmlElement
    p.level = 0
    p_pr = p._element.get_or_add_pPr()
    buNone = OxmlElement('a:buNone')
    p_pr.insert(0, buNone)
    
    p.space_before = Pt(10)
    p.space_after = Pt(2)
    
    run = p.add_run()
    run.text = tag.get_text()
    
    font_conf = generator.fonts_conf.get('title_h3', {'name': 'Meiryo', 'size_pt': 20, 'bold': True})
    apply_font_style(run, font_conf)
    
    # 段落レベルでフォントサイズを固定（はみ出し防止のベース）
    if 'size_pt' in font_conf:
        p.font.size = Pt(font_conf['size_pt'])
    generator.slide_has_text = True

def process_hr(generator: PPTXGenerator, tag: Tag) -> None:
    """水平線（---）による新しいスライド（タイトルなし）の生成"""
    if generator.current_slide:
        from utils import auto_shrink_text
        auto_shrink_text(generator.current_slide)
        
    generator.current_slide = generator.prs.slides.add_slide(generator.prs.slide_layouts[1])
    
    # タイトルシェイプを削除して上部から広く使えるようにする
    if generator.current_slide.shapes.title:
        sp = generator.current_slide.shapes.title._element
        sp.getparent().remove(sp)
        
    generator.current_body = generator.current_slide.placeholders[1].text_frame
    generator.current_body.word_wrap = True
    
    from pptx.enum.text import MSO_AUTO_SIZE
    generator.current_body.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    generator.current_body.text = "" 
    generator.slide_has_text = False
    
    if not generator.slides_conf.get('template_path'):
        try:
            body_shape = generator.current_slide.placeholders[1]
            o_left, o_top, o_width = body_shape.left, body_shape.top, body_shape.width
            new_top = Inches(0.5)
            new_height = generator.layout.body_height_for(new_top)
            body_shape.left, body_shape.top, body_shape.width, body_shape.height = o_left, new_top, o_width, new_height
        except Exception:
            pass

def extract_note_text(tag: Tag) -> str:
    """引用ブロックから、段落構造を保ったままノート用のテキストを取り出す

    tag.get_text() は段落や箇条書きの区切りを落として全体を連結してしまうため、
    ブロック単位で取り出して改行でつなぐ。
    """
    # リスト項目の中の p は、項目側でまとめて扱うため除外する
    blocks = [
        block for block in tag.find_all(NOTE_BLOCK_TAGS)
        if not (block.name == 'p' and block.find_parent('li'))
    ]
    if not blocks:
        return tag.get_text(strip=True)

    parts: list[str] = []
    for block in blocks:
        text = block.get_text().strip()
        if not text:
            continue
        if parts:
            # 箇条書きは詰めて、段落どうしは空行を挟む
            parts.append('\n' if block.name == 'li' else '\n\n')
        parts.append(text)
    return ''.join(parts)

def process_blockquote(generator: PPTXGenerator, tag: Tag) -> None:
    """スピーカーノートの処理"""
    text_frame = generator.current_slide.notes_slide.notes_text_frame
    note_text = extract_note_text(tag)
    if not note_text:
        return
    text_frame.text = text_frame.text + "\n\n" + note_text if text_frame.text else note_text

def inline_code_conf(generator: PPTXGenerator) -> FontConfig | None:
    """インラインコード（`code`）のフォント設定を取り出す"""
    return generator.fonts_conf.get('inline_code')

def append_code_textbox(
    generator: PPTXGenerator, content: str, language: str | None = None
) -> None:
    """レイアウトを決めて、背景色付きのコード枠をスライドに追加する"""
    layout = generator.layout

    if generator.slide_has_text or generator.forced_layout == '2-column':
        shrink_body_shape(generator.current_slide, layout.code_split_body_width)
        box = (layout.code_split_left, layout.content_top,
               layout.code_split_width, layout.content_height)
    elif generator.forced_layout == 'center':
        box = (layout.code_center_left, layout.content_top,
               layout.code_center_width, layout.content_height)
    else:
        box = (layout.content_left, layout.code_full_top,
               layout.content_width, layout.code_full_height)

    theme = generator.config.get('theme') or {}
    create_code_textbox(
        generator.current_slide, *box,
        content=content,
        language=language,
        font_conf=generator.fonts_conf.get('code_block', DEFAULT_CODE_BLOCK_FONT),
        background_rgb=theme.get('code_bg_color', DEFAULT_CODE_BG_COLOR),
    )
    generator.slide_has_text = True

def image_dpi(generator: PPTXGenerator) -> int | None:
    """埋め込み画像の解像度を返す（images.downscale が false の場合は None＝縮小しない）"""
    if not generator.images_conf.get('downscale', True):
        return None
    return generator.images_conf.get('dpi', DEFAULT_IMAGE_DPI)

def place_image(
    generator: PPTXGenerator,
    img_data: ImageSource,
    left: Length,
    top: Length,
    width: Length,
    height: Length,
) -> None:
    """設定に応じて画像を縮小したうえで、指定した枠に収めて配置する"""
    dpi = image_dpi(generator)
    if dpi:
        img_data = downscale_image(img_data, width, height, dpi)
    insert_image_fit(generator.current_slide, img_data, left, top, width, height)

def place_image_full(generator: PPTXGenerator, img_data: ImageSource) -> None:
    """コンテンツ領域いっぱいに図を配置する"""
    layout = generator.layout
    place_image(
        generator, img_data,
        layout.content_left, layout.content_top, layout.content_width, layout.content_height,
    )

def place_image_split(generator: PPTXGenerator, img_data: ImageSource) -> None:
    """本文枠を左に縮め、図を右半分に配置する（2カラム）"""
    layout = generator.layout
    shrink_body_shape(generator.current_slide, layout.split_body_width)
    place_image(
        generator, img_data,
        layout.split_image_left, layout.content_top,
        layout.split_image_width, layout.content_height,
    )

def load_image(src: str) -> ImageSource:
    """画像URLならダウンロードし、ローカルパスならそのまま返す"""
    if src.startswith(('http://', 'https://')):
        response = requests.get(src, timeout=HTTP_TIMEOUT_SEC)
        response.raise_for_status()
        return BytesIO(response.content)
    return src

def process_image(generator: PPTXGenerator, tag: Tag) -> None:
    """画像の挿入処理"""
    img_url = tag.get('src')
    if not img_url:
        print("Warning: src属性が無い画像をスキップしました。")
        return

    try:
        img_data = load_image(cast(str, img_url))
        pos = generator.images_conf.get('position_inches')
        
        if pos and len(pos) >= 2:
            # YAMLの固定位置（幅は縦横比に従うため、上限はスライド幅とする）
            fixed_height = Inches(generator.images_conf.get('default_height_inches', 3.5))
            dpi = image_dpi(generator)
            if dpi:
                img_data = downscale_image(img_data, generator.prs.slide_width, fixed_height, dpi)
            generator.current_slide.shapes.add_picture(img_data, Inches(pos[0]), Inches(pos[1]), height=fixed_height)
        elif generator.forced_layout == 'center':
            place_image_full(generator, img_data)
        else:
            # オートレイアウト
            if generator.slide_has_text or generator.forced_layout == '2-column':
                place_image_split(generator, img_data)
            else:
                place_image_full(generator, img_data)
    except Exception as e:
        print(f"Warning: 画像の挿入に失敗しました: {e}")

def process_table(generator: PPTXGenerator, tag: Tag) -> None:
    """表の挿入処理"""
    layout = generator.layout
    rows = tag.find_all('tr')
    if not rows: return
    
    num_rows = len(rows)
    num_cols = max(len(row.find_all(['th', 'td'])) for row in rows)
    
    if generator.slide_has_text:
        shrink_body_shape(
            generator.current_slide,
            layout.content_width,
            max_height=layout.table_split_body_height,
        )
        table_top = layout.table_split_top
    else:
        table_top = layout.content_top

    table_shape = generator.current_slide.shapes.add_table(
        num_rows, num_cols,
        layout.content_left, table_top, layout.content_width, layout.table_height,
    )
    table = table_shape.table
    
    for row_idx, row in enumerate(rows):
        cols = row.find_all(['th', 'td'])
        for col_idx, col in enumerate(cols):
            if col_idx < num_cols:
                cell = table.cell(row_idx, col_idx)
                cell.text = "" 
                p = cell.text_frame.paragraphs[0]
                
                if col.name == 'th':
                    font_conf = generator.fonts_conf.get('table_header', {'name': 'Meiryo', 'size_pt': 14, 'bold': True, 'color_rgb': [255, 255, 255]})
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(31, 73, 125) # 濃い青をデフォルトに設定
                else:
                    font_conf = generator.fonts_conf.get('table_body', {'name': 'Meiryo', 'size_pt': 12})
                    
                add_runs_from_tag(col, p, font_conf, inline_code_conf(generator))
    generator.slide_has_text = True

def process_code_or_mermaid(generator: PPTXGenerator, tag: Tag) -> None:
    """コードブロックまたはMermaid図形の処理"""
    code_tag = tag.find('code')
    # class属性が無いコードブロックでは get('class') が None を返すため、必ず空リストへ倒す
    classes: list[str] = cast(
        "list[str]", (code_tag.get('class') or []) if isinstance(code_tag, Tag) else []
    )
    is_mermaid = 'language-mermaid' in classes or 'mermaid' in classes

    language = None
    for cls in classes:
        if cls.startswith('language-'):
            language = cls.replace('language-', '')
            break

    if is_mermaid and isinstance(code_tag, Tag):
        try:
            print("INFO: Mermaid図形を生成中...")
            image = render_mermaid(mermaid_conf(generator), code_tag.get_text())
            if image is None:  # renderer: off
                return

            if generator.slide_has_text:
                place_image_split(generator, BytesIO(image))
            else:
                place_image_full(generator, BytesIO(image))
        except Exception as e:
            print(f"Warning: Mermaid図形の生成に失敗しました: {e}")
    else:
        append_code_textbox(generator, tag.get_text(), language=language)

def bullet_font_conf(fonts_conf: dict[str, FontConfig], level: int) -> FontConfig | None:
    """指定レベルの箇条書きのフォント設定を返す

    設定が無いレベルは、より浅いレベルの設定を引き継ぐ。
    いきなり body へフォールバックすると、ネストした途端に書体やサイズが
    変わってしまうため（body は本文用で、箇条書きとは役割が異なる）。
    """
    for candidate in range(level + 1, 0, -1):
        conf = fonts_conf.get(f'bullet_level_{candidate}')
        if conf:
            return conf
    return fonts_conf.get('body')

def process_text(generator: PPTXGenerator, tag: Tag) -> None:
    """段落・リストの処理"""
    if not tag.get_text(strip=True): return

    if tag.name == 'li':
        level = min(len(tag.find_parents(['ul', 'ol'])) - 1, MAX_BULLET_LEVEL)
        font_conf = bullet_font_conf(generator.fonts_conf, level)
    else:
        level = 0
        font_conf = generator.fonts_conf.get('body')


    append_text_block(
        generator.current_body, tag,
        reuse_first_paragraph=not generator.slide_has_text,
        level=level,
        font_conf=font_conf,
        inline_code_conf=inline_code_conf(generator),
    )
    generator.slide_has_text = True
