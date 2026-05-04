import collections
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import markdown
from bs4 import BeautifulSoup, NavigableString, Tag
import requests
from io import BytesIO
import os
import argparse
import yaml
import traceback # エラー解析用に追加

def get_slide_size_from_layout(layout_str):
    sizes = {
        "16:9": (Inches(10), Inches(5.625)),
        "4:3": (Inches(10), Inches(7.5)),
        "16:10": (Inches(10), Inches(6.25)),
        "A4": (Inches(11.69), Inches(8.27))
    }
    return sizes.get(layout_str, sizes["16:9"])

def apply_font_style(run, font_config):
    if not font_config:
        return
    font = run.font
    if 'name' in font_config:
        font.name = font_config['name']
    if 'size_pt' in font_config:
        font.size = Pt(font_config['size_pt'])
    if 'bold' in font_config:
        font.bold = font_config['bold']
    if 'color_rgb' in font_config:
        rgb = font_config['color_rgb']
        font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

def add_runs_from_tag(element, paragraph, default_font_conf, fonts_conf):
    for child in element:
        if isinstance(child, NavigableString):
            text = str(child).replace('\n', ' ')
            if text.strip() or text == ' ':
                run = paragraph.add_run()
                run.text = text
                apply_font_style(run, default_font_conf)
        elif isinstance(child, Tag):
            if child.name in ['ul', 'ol', 'pre', 'img', 'table']: continue
            if child.name in ['p', 'div', 'span', 'li', 'th', 'td']:
                add_runs_from_tag(child, paragraph, default_font_conf, fonts_conf)
            else:
                run = paragraph.add_run()
                run.text = child.get_text().replace('\n', ' ')
                apply_font_style(run, default_font_conf)
                if child.name in ['strong', 'b']: run.font.bold = True
                elif child.name in ['em', 'i']: run.font.italic = True
                elif child.name == 'code':
                    run.font.name = fonts_conf.get('inline_code', {}).get('name', 'Consolas')
                    rgb = fonts_conf.get('inline_code', {}).get('color_rgb', [220, 20, 60])
                    run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

def markdown_to_pptx_v2(md_text, output_file, config):
    html = markdown.markdown(md_text, extensions=['extra', 'fenced_code', 'sane_lists'])
    soup = BeautifulSoup(html, 'html.parser')

    slides_conf = config.get('slides', {})
    prs = Presentation(slides_conf.get('template_path')) if slides_conf.get('template_path') and os.path.exists(slides_conf.get('template_path')) else Presentation()
    if not slides_conf.get('template_path'):
        width, height = get_slide_size_from_layout(slides_conf.get('layout', '16:9'))
        prs.slide_width, prs.slide_height = width, height

    LAYOUT_TITLE = 0
    LAYOUT_CONTENT = 1
    
    current_slide = None
    current_body = None
    slide_has_text = False

    fonts_conf = config.get('fonts', {})
    images_conf = config.get('images', {})

    for tag in soup.find_all(['h1', 'h2', 'p', 'li', 'img', 'pre', 'table']):
        if tag.name == 'p' and tag.find_parent('li'): continue

        # --- 新しいスライドの作成 ---
        if tag.name in ['h1', 'h2']:
            layout = prs.slide_layouts[LAYOUT_TITLE] if tag.name == 'h1' else prs.slide_layouts[LAYOUT_CONTENT]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = tag.get_text()
            
            style_key = 'title_h1' if tag.name == 'h1' else 'title_h2'
            for run in slide.shapes.title.text_frame.paragraphs[0].runs:
                apply_font_style(run, fonts_conf.get(style_key, fonts_conf.get('title')))

            current_slide = slide
            current_body = slide.placeholders[1].text_frame
            current_body.text = "" 
            slide_has_text = False

        # --- 画像の処理 ---
        elif tag.name == 'img' and current_slide:
            img_url = tag.get('src')
            try:
                img_data = BytesIO(requests.get(img_url).content) if img_url.startswith('http') else img_url
                
                pos = images_conf.get('position_inches')
                if pos and len(pos) >= 2:
                    left, top = Inches(pos[0]), Inches(pos[1])
                    img_height = Inches(images_conf.get('default_height_inches', 3.5))
                    current_slide.shapes.add_picture(img_data, left, top, height=img_height)
                else:
                    if slide_has_text:
                        try:
                            body_shape = current_slide.placeholders[1]
                            body_shape.left, body_shape.top = body_shape.left, body_shape.top
                            body_shape.height, body_shape.width = body_shape.height, Inches(4.8)
                        except Exception as e: pass
                        current_slide.shapes.add_picture(img_data, Inches(5.2), Inches(1.8), width=Inches(4.3))
                    else:
                        current_slide.shapes.add_picture(img_data, Inches(1.0), Inches(1.5), width=Inches(8.0))
            except Exception as e: print(f"Warning: 画像の挿入に失敗しました: {e}")

        # --- 表 (Table) の処理 ---
        elif tag.name == 'table' and current_slide:
            rows = tag.find_all('tr')
            if not rows: continue
            
            num_rows = len(rows)
            num_cols = max(len(row.find_all(['th', 'td'])) for row in rows)
            
            if slide_has_text:
                try:
                    body_shape = current_slide.placeholders[1]
                    body_shape.left, body_shape.top = body_shape.left, body_shape.top
                    # 高さを2.0インチに緩和してテキスト消失を防ぐ
                    body_shape.width, body_shape.height = body_shape.width, Inches(2.0)
                except Exception as e: pass
                table_top = Inches(2.8) # 少しゆとりを持たせる
            else:
                table_top = Inches(1.5)
                
            table_left = Inches(1.0)
            table_width = Inches(8.0)
            table_height = Inches(0.8)
            
            table_shape = current_slide.shapes.add_table(num_rows, num_cols, table_left, table_top, table_width, table_height)
            table = table_shape.table
            
            for row_idx, row in enumerate(rows):
                cols = row.find_all(['th', 'td'])
                for col_idx, col in enumerate(cols):
                    if col_idx < num_cols:
                        cell = table.cell(row_idx, col_idx)
                        cell.text = "" 
                        p = cell.text_frame.paragraphs[0]
                        
                        if col.name == 'th':
                            font_conf = fonts_conf.get('table_header', {'name': 'Meiryo', 'size_pt': 14, 'bold': True, 'color_rgb': [255, 255, 255]})
                        else:
                            font_conf = fonts_conf.get('table_body', {'name': 'Meiryo', 'size_pt': 12})
                            
                        add_runs_from_tag(col, p, font_conf, fonts_conf)
                        
            slide_has_text = True

        # --- コードブロックの処理 ---
        elif tag.name == 'pre' and current_body:
            if not slide_has_text and len(current_body.paragraphs) == 1 and not current_body.paragraphs[0].text:
                p = current_body.paragraphs[0]
            else:
                p = current_body.add_paragraph()
                
            run = p.add_run()
            run.text = tag.get_text()
            font_conf = fonts_conf.get('code_block', {'name': 'Consolas', 'size_pt': 12})
            apply_font_style(run, font_conf)
            if 'color_rgb' not in font_conf: run.font.color.rgb = RGBColor(0, 80, 160)
            slide_has_text = True

        # --- テキスト・リストの処理 ---
        elif tag.name in ['li', 'p'] and current_body:
            if not tag.get_text(strip=True):
                continue

            if not slide_has_text and len(current_body.paragraphs) == 1 and not current_body.paragraphs[0].text:
                p = current_body.paragraphs[0]
            else:
                p = current_body.add_paragraph()
                
            if tag.name == 'li':
                p.level = min(len(tag.find_parents(['ul', 'ol'])) - 1, 8)
                font_conf = fonts_conf.get(f'bullet_level_{p.level + 1}', fonts_conf.get('body'))
            else:
                p.level = 0
                font_conf = fonts_conf.get('body')
            
            add_runs_from_tag(tag, p, font_conf, fonts_conf)
            slide_has_text = True 

    prs.save(output_file)

def main():
    print("INFO: 変換処理を開始します...")
    parser = argparse.ArgumentParser()
    parser.add_argument("input")
    parser.add_argument("-o", "--output", default="output.pptx")
    parser.add_argument("-c", "--config", default="config.yaml")
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"Error: 入力ファイル '{args.input}' が見つかりません。")
        return

    if not os.path.exists(args.config):
        print(f"Error: 設定ファイル '{args.config}' が見つかりません。")
        return

    try:
        with open(args.config, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        with open(args.input, "r", encoding="utf-8") as f:
            content = f.read()
            
        markdown_to_pptx_v2(content, args.output, config)
        print(f"Success: '{args.output}' の生成が完了しました！")
        
    except PermissionError:
        print(f"Error: '{args.output}' に書き込めません！PowerPointでファイルを開いたままにしていませんか？閉じてから再実行してください。")
    except Exception as e:
        print(f"Error: 予期せぬエラーが発生しました: {e}")
        traceback.print_exc()

if __name__ == "__main__": 
    main()