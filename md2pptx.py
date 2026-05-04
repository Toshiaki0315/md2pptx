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

class PPTXGenerator:
    """MarkdownをPowerPointに変換するジェネレータークラス"""

    def __init__(self, config):
        self.config = config
        self.slides_conf = config.get('slides', {})
        self.fonts_conf = config.get('fonts', {})
        self.images_conf = config.get('images', {})
        
        self.prs = None
        self.current_slide = None
        self.current_body = None
        self.slide_has_text = False
        
        self._init_presentation()

    def _init_presentation(self):
        """プレゼンテーションの初期化（テンプレート読み込み・サイズ設定）"""
        template_path = self.slides_conf.get('template_path')
        if template_path and os.path.exists(template_path):
            self.prs = Presentation(template_path)
        else:
            self.prs = Presentation()
            width, height = self._get_slide_size(self.slides_conf.get('layout', '16:9'))
            self.prs.slide_width = width
            self.prs.slide_height = height

    def _get_slide_size(self, layout_str):
        sizes = {
            "16:9": (Inches(10), Inches(5.625)),
            "4:3":  (Inches(10), Inches(7.5)),
            "16:10": (Inches(10), Inches(6.25)),
            "A4":   (Inches(11.69), Inches(8.27))
        }
        return sizes.get(layout_str, sizes["16:9"])

    @staticmethod
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

    @staticmethod
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

    def _add_runs_from_tag(self, element, paragraph, default_font_conf):
        """インライン装飾を解釈しながらテキストを追加（再帰処理）"""
        for child in element:
            if isinstance(child, NavigableString):
                text = str(child).replace('\n', ' ')
                if text.strip() or text == ' ':
                    run = paragraph.add_run()
                    run.text = text
                    self.apply_font_style(run, default_font_conf)
            elif isinstance(child, Tag):
                if child.name in ['ul', 'ol', 'pre', 'img', 'table', 'blockquote']: continue
                if child.name in ['p', 'div', 'span', 'li', 'th', 'td']:
                    self._add_runs_from_tag(child, paragraph, default_font_conf)
                else:
                    run = paragraph.add_run()
                    run.text = child.get_text().replace('\n', ' ')
                    self.apply_font_style(run, default_font_conf)
                    
                    if child.name in ['strong', 'b']: run.font.bold = True
                    elif child.name in ['em', 'i']: run.font.italic = True
                    elif child.name == 'code':
                        run.font.name = self.fonts_conf.get('inline_code', {}).get('name', 'Consolas')
                        rgb = self.fonts_conf.get('inline_code', {}).get('color_rgb', [220, 20, 60])
                        run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

    def _shrink_body_shape(self, width_inches=4.8, max_height_inches=None):
        """テキスト枠を指定サイズに縮める（レイアウト調整用ヘルパー）"""
        try:
            body_shape = self.current_slide.placeholders[1]
            body_shape.left, body_shape.top = body_shape.left, body_shape.top
            body_shape.width = Inches(width_inches)
            if max_height_inches:
                body_shape.height = Inches(max_height_inches)
        except Exception:
            pass

    def _process_heading(self, tag):
        """見出しタグの処理とスライド作成"""
        layout_idx = 0 if tag.name == 'h1' else 1
        self.current_slide = self.prs.slides.add_slide(self.prs.slide_layouts[layout_idx])
        self.current_slide.shapes.title.text = tag.get_text()
        
        style_key = 'title_h1' if tag.name == 'h1' else 'title_h2'
        for run in self.current_slide.shapes.title.text_frame.paragraphs[0].runs:
            self.apply_font_style(run, self.fonts_conf.get(style_key, self.fonts_conf.get('title')))

        self.current_body = self.current_slide.placeholders[1].text_frame
        self.current_body.text = "" 
        self.slide_has_text = False

        # デフォルト枠のはみ出し補正
        if not self.slides_conf.get('template_path'):
            try:
                body_shape = self.current_slide.placeholders[1]
                o_left, o_top, o_width = body_shape.left, body_shape.top, body_shape.width
                new_height = self.prs.slide_height - o_top - Inches(0.5)
                body_shape.left, body_shape.top, body_shape.width, body_shape.height = o_left, o_top, o_width, new_height
            except Exception:
                pass

    def _process_blockquote(self, tag):
        """スピーカーノートの処理"""
        text_frame = self.current_slide.notes_slide.notes_text_frame
        note_text = tag.get_text(strip=True)
        text_frame.text = text_frame.text + "\n\n" + note_text if text_frame.text else note_text

    def _process_image(self, tag):
        """画像の挿入処理"""
        img_url = tag.get('src')
        try:
            img_data = BytesIO(requests.get(img_url).content) if img_url.startswith('http') else img_url
            pos = self.images_conf.get('position_inches')
            
            if pos and len(pos) >= 2:
                # YAMLの固定位置
                self.current_slide.shapes.add_picture(img_data, Inches(pos[0]), Inches(pos[1]), height=Inches(self.images_conf.get('default_height_inches', 3.5)))
            else:
                # オートレイアウト
                if self.slide_has_text:
                    self._shrink_body_shape(width_inches=4.8)
                    self.insert_image_fit(self.current_slide, img_data, Inches(5.2), Inches(1.5), Inches(4.5), Inches(3.8))
                else:
                    self.insert_image_fit(self.current_slide, img_data, Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8))
        except Exception as e:
            print(f"Warning: 画像の挿入に失敗しました: {e}")

    def _process_table(self, tag):
        """表の挿入処理"""
        rows = tag.find_all('tr')
        if not rows: return
        
        num_rows = len(rows)
        num_cols = max(len(row.find_all(['th', 'td'])) for row in rows)
        
        if self.slide_has_text:
            self._shrink_body_shape(width_inches=8.0, max_height_inches=2.0)
            table_top = Inches(2.8) 
        else:
            table_top = Inches(1.5)
            
        table_shape = self.current_slide.shapes.add_table(num_rows, num_cols, Inches(1.0), table_top, Inches(8.0), Inches(0.8))
        table = table_shape.table
        
        for row_idx, row in enumerate(rows):
            cols = row.find_all(['th', 'td'])
            for col_idx, col in enumerate(cols):
                if col_idx < num_cols:
                    cell = table.cell(row_idx, col_idx)
                    cell.text = "" 
                    p = cell.text_frame.paragraphs[0]
                    
                    if col.name == 'th':
                        font_conf = self.fonts_conf.get('table_header', {'name': 'Meiryo', 'size_pt': 14, 'bold': True, 'color_rgb': [255, 255, 255]})
                    else:
                        font_conf = self.fonts_conf.get('table_body', {'name': 'Meiryo', 'size_pt': 12})
                        
                    self._add_runs_from_tag(col, p, font_conf)
        self.slide_has_text = True

    def _process_code_or_mermaid(self, tag):
        """コードブロックまたはMermaid図形の処理"""
        code_tag = tag.find('code')
        classes = code_tag.get('class') if code_tag else []
        is_mermaid = 'language-mermaid' in classes or 'mermaid' in classes

        if is_mermaid:
            try:
                print("INFO: Mermaid図形をAPIで生成中...")
                compressed = zlib.compress(code_tag.get_text().encode('utf-8'), 9)
                encoded = base64.urlsafe_b64encode(compressed).decode('ascii')
                response = requests.get(f"https://kroki.io/mermaid/png/{encoded}", timeout=15)
                response.raise_for_status()
                
                if self.slide_has_text:
                    self._shrink_body_shape(width_inches=4.8)
                    self.insert_image_fit(self.current_slide, BytesIO(response.content), Inches(5.2), Inches(1.5), Inches(4.5), Inches(3.8))
                else:
                    self.insert_image_fit(self.current_slide, BytesIO(response.content), Inches(1.0), Inches(1.5), Inches(8.0), Inches(3.8))
            except Exception as e:
                print(f"Warning: Mermaid図形の生成に失敗しました: {e}")
        else:
            self._append_text_block(tag.get_text(), is_code=True)
            self.slide_has_text = True

    def _process_text(self, tag):
        """段落・リストの処理"""
        if not tag.get_text(strip=True): return
        
        level = min(len(tag.find_parents(['ul', 'ol'])) - 1, 8) if tag.name == 'li' else 0
        font_conf = self.fonts_conf.get(f'bullet_level_{level + 1}' if tag.name == 'li' else 'body', self.fonts_conf.get('body'))
        
        self._append_text_block(tag, is_code=False, level=level, font_conf=font_conf)
        self.slide_has_text = True

    def _append_text_block(self, content, is_code=False, level=0, font_conf=None):
        """段落オブジェクトを追加し、テキストまたはタグ構造を書き込むヘルパー"""
        if not self.slide_has_text and len(self.current_body.paragraphs) == 1 and not self.current_body.paragraphs[0].text:
            p = self.current_body.paragraphs[0]
        else:
            p = self.current_body.add_paragraph()
            
        p.level = level
        
        if is_code:
            run = p.add_run()
            run.text = content
            conf = self.fonts_conf.get('code_block', {'name': 'Consolas', 'size_pt': 12})
            self.apply_font_style(run, conf)
            if 'color_rgb' not in conf: run.font.color.rgb = RGBColor(0, 80, 160)
        else:
            self._add_runs_from_tag(content, p, font_conf)

    def generate(self, md_text, output_file):
        """MarkdownをパースしてPPTXを生成するメイン処理"""
        html = markdown.markdown(md_text, extensions=['extra', 'fenced_code', 'sane_lists'])
        soup = BeautifulSoup(html, 'html.parser')

        for tag in soup.find_all(['h1', 'h2', 'p', 'li', 'img', 'pre', 'table', 'blockquote']):
            if tag.name == 'p' and (tag.find_parent('li') or tag.find_parent('blockquote')):
                continue

            if tag.name in ['h1', 'h2']:
                self._process_heading(tag)
            elif tag.name == 'blockquote' and self.current_slide:
                self._process_blockquote(tag)
            elif tag.name == 'img' and self.current_slide:
                self._process_image(tag)
            elif tag.name == 'table' and self.current_slide:
                self._process_table(tag)
            elif tag.name == 'pre' and self.current_body:
                self._process_code_or_mermaid(tag)
            elif tag.name in ['li', 'p'] and self.current_body:
                self._process_text(tag)

        self.prs.save(output_file)

def main():
    print("INFO: 変換処理を開始します...")
    parser = argparse.ArgumentParser(description="MarkdownファイルをPowerPointに変換します。")
    parser.add_argument("input", help="変換するMarkdownファイルのパス")
    parser.add_argument("-o", "--output", help="出力ファイル名", default="output.pptx")
    parser.add_argument("-c", "--config", help="YAML設定ファイルのパス", default="config.yaml")
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
            
        generator = PPTXGenerator(config)
        generator.generate(content, args.output)
        
        print(f"Success: '{args.output}' の生成が完了しました！")
        
    except PermissionError:
        print(f"Error: '{args.output}' に書き込めません！PowerPointでファイルを開いたままにしていませんか？閉じてから再実行してください。")
    except Exception as e:
        print(f"Error: 予期せぬエラーが発生しました: {e}")
        traceback.print_exc()

if __name__ == "__main__": 
    main()