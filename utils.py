import collections
import collections.abc
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from bs4 import NavigableString, Tag
from pygments import lex
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.styles import get_style_by_name

def hex_to_rgb(hex_str):
    if not hex_str: return None
    hex_str = hex_str.lstrip('#')
    return RGBColor(*tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

def apply_syntax_highlight(p, text, language, font_conf):
    try:
        if language:
            lexer = get_lexer_by_name(language, stripall=False)
        else:
            lexer = guess_lexer(text)
    except:
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

def apply_font_style(run, font_config):
    """フォントスタイルの適用"""
    if not font_config: return
    font = run.font
    if 'name' in font_config: font.name = font_config['name']
    if 'size_pt' in font_config: font.size = Pt(font_config['size_pt'])
    if 'bold' in font_config: font.bold = font_config['bold']
    if 'color_rgb' in font_config:
        rgb = font_config['color_rgb']
        font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

def insert_image_fit(slide, img_data, left, top, max_width, max_height):
    """画像を最大枠に収まるようにアスペクト比を保って自動縮小・中央配置する"""
    pic = slide.shapes.add_picture(img_data, left, top)
    ratio_w = max_width / pic.width
    ratio_h = max_height / pic.height
    ratio = min(ratio_w, ratio_h)
    ratio = min(ratio, 1.5) # 極端な拡大を防止
    
    pic.width = int(pic.width * ratio)
    pic.height = int(pic.height * ratio)
    pic.left = int(left + (max_width - pic.width) / 2)
    pic.top = int(top + (max_height - pic.height) / 2)
    return pic

def add_runs_from_tag(generator, element, paragraph, default_font_conf):
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

def shrink_body_shape(generator, width_inches=4.8, max_height_inches=None):
    """テキスト枠を指定サイズに縮める（レイアウト調整用ヘルパー）"""
    try:
        body_shape = generator.current_slide.placeholders[1]
        body_shape.left, body_shape.top = body_shape.left, body_shape.top
        body_shape.width = Inches(width_inches)
        if max_height_inches:
            body_shape.height = Inches(max_height_inches)
    except Exception:
        pass

def append_text_block(generator, content, level=0, font_conf=None):
    """段落オブジェクトを追加し、テキストまたはタグ構造を書き込むヘルパー"""
    if not generator.slide_has_text and len(generator.current_body.paragraphs) == 1 and not generator.current_body.paragraphs[0].text:
        p = generator.current_body.paragraphs[0]
    else:
        p = generator.current_body.add_paragraph()
        
    p.level = level
    p.space_after = Pt(12)  # 段落後の余白を追加してレイアウトを美しく
    p.line_spacing = 1.2    # 行間を1.2倍に設定
    
    add_runs_from_tag(generator, content, p, font_conf)

def append_code_textbox(generator, content, language=None):
    """独立したテキストボックスを作成し、背景色付きでコードを挿入する"""
    if generator.slide_has_text or generator.forced_layout == '2-column':
        shrink_body_shape(generator, width_inches=4.5)
        box_left = Inches(5.0)
        box_top = Inches(1.5)
        box_width = Inches(4.5)
        box_height = Inches(3.8)
    elif generator.forced_layout == 'center':
        box_left = Inches(1.5)
        box_top = Inches(1.5)
        box_width = Inches(7.0)
        box_height = Inches(3.8)
    else:
        box_left = Inches(1.0)
        box_top = Inches(2.0)
        box_width = Inches(8.0)
        box_height = Inches(3.0)
        
    textbox = generator.current_slide.shapes.add_textbox(box_left, box_top, box_width, box_height)
    textbox.fill.solid()
    textbox.fill.fore_color.rgb = RGBColor(40, 44, 52) # 濃いグレー（Monokai風）
    
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
