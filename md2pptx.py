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
import traceback
import base64
import zlib

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

def insert_image_fit(slide, img_data, left, top, max_width, max_height):
    """画像を最大枠(max_width, max_height)に収まるようにアスペクト比を保って自動縮小・中央配置する"""
    # 一度画像をネイティブサイズで挿入して縦横サイズを取得
    pic = slide.shapes.add_picture(img_data, left, top)
    
    # 縦横の「枠に収まるための縮小率」を計算し、厳しい（小さい）方を採用
    ratio_w = max_width / pic.width
    ratio_h = max_height / pic.height
    ratio = min(ratio_w, ratio_h)
    
    # 極端な拡大を防ぐため、元のサイズの1.5倍までを上限とする
    ratio = min(ratio, 1.5)
    
    # リサイズ適用
    pic.width = int(pic.width * ratio)
    pic.height = int(pic.height * ratio)
    
    # 用意した枠の中央に美しく配置する
    pic.left = int(left + (max_width - pic.width) / 2)
    pic.top = int(top + (max_height - pic.height) / 2)
    
    return pic

def add_runs_from_tag(element, paragraph, default_font_conf, fonts_conf):
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

    for tag in soup.find_all(['h1', 'h2', 'p', 'li', 'img', 'pre', 'table', 'blockquote']):
        if tag.name == 'p' and (tag.find_parent('li') or tag.find_parent('blockquote')): 
            continue

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

            if not slides_conf.get('template_path'):
                try:
                    body_shape = current_slide.placeholders[1]
                    # 変更前に、現在の4つのプロパティ全てを保存する
                    original_left = body_shape.left
                    original_top = body_shape.top
                    original_width = body_shape.width
                    
                    # 高さを再計算
                    new_height = prs.slide_height - original_top - Inches(0.5)
                    
                    # 座標リセットを防ぐため、leftとtopも含めて全て同時にセットし直す
                    body_shape.left = original_left
                    body_shape.top = original_top
                    body_shape.width = original_width
                    body_shape.height = new_height
                except Exception:
                    pass
            
        # --- スピーカーノートの処理 ---
        elif tag.name == 'blockquote' and current_slide:
            notes_slide = current_slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            note_text = tag.get_text(strip=True)
            if text_frame.text: text_frame.text += "\n\n" + note_text
            else: text_frame.text = note_text

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
                        # 最大枠を指定してフィット配置
                        insert_image_fit(current_slide, img_data, Inches(5.2), Inches(1.5), Inches(4.5), Inches(3.8))
                    else:
                        insert_image_fit(current_slide, img_data, Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8))
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
                    body_shape.width, body_shape.height = body_shape.width, Inches(2.0)
                except Exception as e: pass
                table_top = Inches(2.8) 
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

        # --- コードブロック / Mermaid の処理 ---
        elif tag.name == 'pre' and current_body:
            code_tag = tag.find('code')
            is_mermaid = False
            
            if code_tag and code_tag.get('class'):
                classes = code_tag.get('class')
                if 'language-mermaid' in classes or 'mermaid' in classes:
                    is_mermaid = True

            if is_mermaid:
                mermaid_text = code_tag.get_text()
                try:
                    print("INFO: Mermaid図形をAPIで生成中...")
                    compressed = zlib.compress(mermaid_text.encode('utf-8'), 9)
                    encoded = base64.urlsafe_b64encode(compressed).decode('ascii')
                    img_url = f"https://kroki.io/mermaid/png/{encoded}"

                    response = requests.get(img_url, timeout=15)
                    response.raise_for_status()
                    img_data = BytesIO(response.content)

                    # Mermaid画像にも「はみ出し防止」のリサイズを適用
                    if slide_has_text:
                        try:
                            body_shape = current_slide.placeholders[1]
                            body_shape.left, body_shape.top = body_shape.left, body_shape.top
                            body_shape.height, body_shape.width = body_shape.height, Inches(4.8)
                        except Exception as e: pass
                        insert_image_fit(current_slide, img_data, Inches(5.2), Inches(1.5), Inches(4.5), Inches(3.8))
                    else:
                        insert_image_fit(current_slide, img_data, Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8))
                except Exception as e:
                    print(f"Warning: Mermaid図形の生成に失敗しました: {e}")
            else:
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