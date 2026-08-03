import requests
from io import BytesIO
import zlib
import base64
from pptx.util import Inches
from pptx.dml.color import RGBColor

from utils import (
    apply_font_style,
    insert_image_fit,
    shrink_body_shape,
    add_runs_from_tag,
    append_text_block,
    append_code_textbox,
    auto_shrink_text
)

def process_heading(generator, tag):
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
            new_height = generator.prs.slide_height - o_top - Inches(0.5)
            body_shape.left, body_shape.top, body_shape.width, body_shape.height = o_left, o_top, o_width, new_height
        except Exception:
            pass

def process_h3(generator, tag):
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

def process_hr(generator, tag):
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
            new_height = generator.prs.slide_height - new_top - Inches(0.5)
            body_shape.left, body_shape.top, body_shape.width, body_shape.height = o_left, new_top, o_width, new_height
        except Exception:
            pass

def process_blockquote(generator, tag):
    """スピーカーノートの処理"""
    text_frame = generator.current_slide.notes_slide.notes_text_frame
    note_text = tag.get_text(strip=True)
    text_frame.text = text_frame.text + "\n\n" + note_text if text_frame.text else note_text

def process_image(generator, tag):
    """画像の挿入処理"""
    img_url = tag.get('src')
    try:
        img_data = BytesIO(requests.get(img_url).content) if img_url.startswith('http') else img_url
        pos = generator.images_conf.get('position_inches')
        
        if pos and len(pos) >= 2:
            # YAMLの固定位置
            generator.current_slide.shapes.add_picture(img_data, Inches(pos[0]), Inches(pos[1]), height=Inches(generator.images_conf.get('default_height_inches', 3.5)))
        elif generator.forced_layout == 'center':
            insert_image_fit(generator.current_slide, img_data, Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8))
        else:
            # オートレイアウト
            if generator.slide_has_text or generator.forced_layout == '2-column':
                shrink_body_shape(generator, width_inches=4.8)
                insert_image_fit(generator.current_slide, img_data, Inches(5.2), Inches(1.5), Inches(4.5), Inches(3.8))
            else:
                insert_image_fit(generator.current_slide, img_data, Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8))
    except Exception as e:
        print(f"Warning: 画像の挿入に失敗しました: {e}")

def process_table(generator, tag):
    """表の挿入処理"""
    rows = tag.find_all('tr')
    if not rows: return
    
    num_rows = len(rows)
    num_cols = max(len(row.find_all(['th', 'td'])) for row in rows)
    
    if generator.slide_has_text:
        shrink_body_shape(generator, width_inches=8.0, max_height_inches=2.0)
        table_top = Inches(2.8) 
    else:
        table_top = Inches(1.5)
        
    table_shape = generator.current_slide.shapes.add_table(num_rows, num_cols, Inches(1.0), table_top, Inches(8.0), Inches(0.8))
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
                    
                add_runs_from_tag(generator, col, p, font_conf)
    generator.slide_has_text = True

def process_code_or_mermaid(generator, tag):
    """コードブロックまたはMermaid図形の処理"""
    code_tag = tag.find('code')
    classes = code_tag.get('class') if code_tag else []
    is_mermaid = 'language-mermaid' in classes or 'mermaid' in classes
    
    language = None
    for cls in classes:
        if cls.startswith('language-'):
            language = cls.replace('language-', '')
            break

    if is_mermaid:
        try:
            print("INFO: Mermaid図形をAPIで生成中...")
            text = code_tag.get_text()
            
            # Kroki API
            compressed = zlib.compress(text.encode('utf-8'), 9)
            encoded = base64.urlsafe_b64encode(compressed).decode('ascii')
            
            response = None
            try:
                response = requests.get(f"https://kroki.io/mermaid/png/{encoded}", timeout=15)
                response.raise_for_status()
            except Exception as e_kroki:
                print(f"INFO: Kroki APIが応答しませんでした。代替API(mermaid.ink)を試行します... ({e_kroki})")
                # mermaid.ink へのフォールバック
                encoded_ink = base64.urlsafe_b64encode(text.encode('utf-8')).decode('ascii')
                response = requests.get(f"https://mermaid.ink/img/{encoded_ink}", timeout=15)
                response.raise_for_status()
            
            if generator.slide_has_text:
                shrink_body_shape(generator, width_inches=4.8)
                insert_image_fit(generator.current_slide, BytesIO(response.content), Inches(5.2), Inches(1.5), Inches(4.5), Inches(3.8))
            else:
                insert_image_fit(generator.current_slide, BytesIO(response.content), Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8))
        except Exception as e:
            print(f"Warning: Mermaid図形の生成に失敗しました: {e}")
    else:
        append_code_textbox(generator, tag.get_text(), language=language)
        # generator.slide_has_text は append_code_textbox 内で True になります

def process_text(generator, tag):
    """段落・リストの処理"""
    if not tag.get_text(strip=True): return
    
    level = min(len(tag.find_parents(['ul', 'ol'])) - 1, 8) if tag.name == 'li' else 0
    font_conf = generator.fonts_conf.get(f'bullet_level_{level + 1}' if tag.name == 'li' else 'body', generator.fonts_conf.get('body'))
    
    append_text_block(generator, tag, level=level, font_conf=font_conf)
    generator.slide_has_text = True
