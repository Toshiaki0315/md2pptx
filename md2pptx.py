import collections
import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
import markdown
from bs4 import BeautifulSoup
import requests
from io import BytesIO
import os
import argparse
import yaml # YAML読み込み用

def get_slide_size_from_layout(layout_str):
    """YAMLの文字列からpython-pptxのスライドサイズ（インチ）を返す"""
    sizes = {
        "16:9": (Inches(10), Inches(5.625)),
        "4:3": (Inches(10), Inches(7.5)),
        "16:10": (Inches(10), Inches(6.25)),
        "A4": (Inches(11.69), Inches(8.27))
    }
    return sizes.get(layout_str, sizes["16:9"]) # デフォルトは16:9

def apply_font_style(run, font_config):
    """指定されたrunにフォント設定を適用する"""
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

def markdown_to_pptx_v2(md_text, output_file, config):
    # MarkdownをHTMLに変換
    html = markdown.markdown(md_text, extensions=['extra'])
    soup = BeautifulSoup(html, 'html.parser')

    # --- プレゼンテーションの初期化（設定を反映） ---
    slides_conf = config.get('slides', {})
    template_path = slides_conf.get('template_path')

    if template_path and os.path.exists(template_path):
        # テンプレートがある場合は読み込む（サイズはテンプレートに従う）
        prs = Presentation(template_path)
        print(f"INFO: テンプレート '{template_path}' を使用します。")
    else:
        # テンプレートがない場合は新規作成し、画角を設定
        prs = Presentation()
        width, height = get_slide_size_from_layout(slides_conf.get('layout', '16:9'))
        prs.slide_width = width
        prs.slide_height = height
        print(f"INFO: 新規作成します。画角: {slides_conf.get('layout', '16:9')}")

    # レイアウト設定 (1: タイトルとコンテンツ)
    bullet_slide_layout = prs.slide_layouts[1]
    
    current_slide = None
    current_body = None

    fonts_conf = config.get('fonts', {})
    images_conf = config.get('images', {})

    for tag in soup.find_all(['h1', 'h2', 'p', 'li', 'img']):
        # --- 新しいスライドの作成 (見出し) ---
        if tag.name in ['h1', 'h2']:
            slide = prs.slides.add_slide(bullet_slide_layout)
            title_shape = slide.shapes.title
            title_shape.text = tag.get_text()
            
            # タイトルのフォント装飾
            for paragraph in title_shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    apply_font_style(run, fonts_conf.get('title'))

            current_slide = slide
            current_body = slide.placeholders[1].text_frame
            current_body.text = "" 

        # --- 画像の処理 ---
        elif tag.name == 'img' and current_slide:
            img_url = tag.get('src')
            try:
                if img_url.startswith('http'):
                    response = requests.get(img_url, timeout=10)
                    img_data = BytesIO(response.content)
                else:
                    img_data = img_url
                
                # 配置（設定ファイルから読み込み）
                pos = images_conf.get('position_inches', [0, 0])
                left = Inches(pos[0])
                top = Inches(pos[1])
                height = Inches(images_conf.get('default_height_inches', 4.0))

                current_slide.shapes.add_picture(img_data, left, top, height=height)
            except Exception as e:
                print(f"Warning: 画像の挿入に失敗しました ({img_url}): {e}")

        # --- テキストの処理 ---
        elif tag.name in ['li', 'p'] and current_body:
            p = current_body.add_paragraph()
            p.text = tag.get_text()
            
            if tag.name == 'li':
                p.level = 1
                font_conf = fonts_conf.get('bullet_level_1', fonts_conf.get('body'))
            else:
                p.level = 0
                font_conf = fonts_conf.get('body')

            # 本文のフォント設定を適用
            if p.runs:
                apply_font_style(p.runs[0], font_conf)

    prs.save(output_file)

def main():
    parser = argparse.ArgumentParser(description="MarkdownファイルをPowerPoint (.pptx) に変換します（設定ファイル対応）。")
    parser.add_argument("input", help="変換するMarkdownファイルのパス")
    parser.add_argument("-o", "--output", help="出力するファイル名 (デフォルト: output.pptx)", default="output.pptx")
    parser.add_argument("-c", "--config", help="使用するYAML設定ファイルのパス (デフォルト: config.yaml)", default="config.yaml")

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: 入力ファイル '{args.input}' が見つかりません。")
        return

    if not os.path.exists(args.config):
        print(f"Error: 設定ファイル '{args.config}' が見つかりません。")
        return

    # 設定ファイルの読み込み
    try:
        with open(args.config, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        print(f"INFO: 設定ファイル '{args.config}' を読み込みました。")
    except Exception as e:
        print(f"Error: 設定ファイルの読み込みに失敗しました: {e}")
        return

    # 変換の実行
    try:
        with open(args.input, "r", encoding="utf-8") as f:
            content = f.read()
        
        markdown_to_pptx_v2(content, args.output, config)
        print(f"Success: '{args.input}' を変換し、'{args.output}' を作成しました。")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error: 変換中にエラーが発生しました: {e}")

if __name__ == "__main__":
    main()