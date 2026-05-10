import os
import yaml
from pptx import Presentation
from pptx.util import Inches
import markdown
from bs4 import BeautifulSoup, Comment
from utils import apply_font_style

from processors import (
    process_heading,
    process_blockquote,
    process_image,
    process_table,
    process_code_or_mermaid,
    process_text
)

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
        self.forced_layout = None
        
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

    def generate(self, md_text, output_file):
        """MarkdownをパースしてPPTXを生成するメイン処理"""
        
        # 1. フロントマターの解析
        front_matter = {}
        if md_text.startswith('---'):
            parts = md_text.split('---', 2)
            if len(parts) >= 3:
                try:
                    front_matter = yaml.safe_load(parts[1])
                    md_text = parts[2]
                except Exception as e:
                    print(f"Warning: フロントマターの解析に失敗しました: {e}")
        
        # タイトルスライドの自動生成（フロントマターがある場合）
        if front_matter and front_matter.get('title'):
            self.current_slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
            self.current_slide.shapes.title.text = str(front_matter.get('title'))
            subtitle_text = []
            if 'subtitle' in front_matter: subtitle_text.append(str(front_matter['subtitle']))
            if 'author' in front_matter: subtitle_text.append(str(front_matter['author']))
            if 'date' in front_matter: subtitle_text.append(str(front_matter['date']))
            if subtitle_text and len(self.current_slide.placeholders) > 1:
                self.current_slide.placeholders[1].text = "\n".join(subtitle_text)
            
            for run in self.current_slide.shapes.title.text_frame.paragraphs[0].runs:
                apply_font_style(run, self.fonts_conf.get('title_h1', self.fonts_conf.get('title')))

        html = markdown.markdown(md_text, extensions=['extra', 'fenced_code', 'sane_lists'])
        soup = BeautifulSoup(html, 'html.parser')

        # 2. タグとコメントの処理
        for element in soup.find_all(['h1', 'h2', 'p', 'li', 'img', 'pre', 'table', 'blockquote', lambda tag: isinstance(tag, Comment)]):
            if isinstance(element, Comment):
                text = element.strip()
                if text.startswith('layout:'):
                    layout_val = text.split(':', 1)[1].strip()
                    self.forced_layout = layout_val
                    if layout_val == '2-column':
                        self.slide_has_text = True # 画像や表を右側に寄せるためのフラグ
                continue

            tag = element
            if tag.name == 'p' and (tag.find_parent('li') or tag.find_parent('blockquote')):
                continue

            if tag.name in ['h1', 'h2']:
                self.forced_layout = None # 新しいスライドでリセット
                process_heading(self, tag)
            elif tag.name == 'blockquote' and self.current_slide:
                process_blockquote(self, tag)
            elif tag.name == 'img' and self.current_slide:
                process_image(self, tag)
            elif tag.name == 'table' and self.current_slide:
                process_table(self, tag)
            elif tag.name == 'pre' and self.current_body:
                process_code_or_mermaid(self, tag)
            elif tag.name in ['li', 'p'] and self.current_body:
                process_text(self, tag)

        self.prs.save(output_file)
