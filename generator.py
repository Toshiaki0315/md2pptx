import os
import yaml
from pptx import Presentation
from pptx.util import Inches
import markdown
from bs4 import BeautifulSoup, Comment
from utils import apply_font_style

from processors import (
    process_heading,
    process_h3,
    process_hr,
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
            from pptx.enum.text import MSO_AUTO_SIZE
            
            self.current_slide = self.prs.slides.add_slide(self.prs.slide_layouts[0])
            
            title_shape = self.current_slide.shapes.title
            if title_shape:
                title_shape.text_frame.word_wrap = True
                title_shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                title_shape.text = str(front_matter.get('title'))
                
                for run in title_shape.text_frame.paragraphs[0].runs:
                    apply_font_style(run, self.fonts_conf.get('title_h1', self.fonts_conf.get('title')))
            
            subtitle_text = []
            if 'subtitle' in front_matter: subtitle_text.append(str(front_matter['subtitle']))
            if 'author' in front_matter: subtitle_text.append(str(front_matter['author']))
            if 'date' in front_matter: subtitle_text.append(str(front_matter['date']))
            
            if subtitle_text and len(self.current_slide.placeholders) > 1:
                sub_shape = self.current_slide.placeholders[1]
                sub_shape.text_frame.word_wrap = True
                sub_shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                sub_shape.text = "\n".join(subtitle_text)
                
                # サブタイトルには少し小さめのフォントサイズを適用
                sub_font_conf = self.fonts_conf.get('body', {'name': 'Meiryo', 'size_pt': 20})
                for p in sub_shape.text_frame.paragraphs:
                    for run in p.runs:
                        apply_font_style(run, sub_font_conf)

        html = markdown.markdown(md_text, extensions=['extra', 'fenced_code', 'sane_lists'])
        soup = BeautifulSoup(html, 'html.parser')

        # 2. タグとコメントの処理
        for element in soup.find_all(['h1', 'h2', 'h3', 'hr', 'p', 'li', 'img', 'pre', 'table', 'blockquote', lambda tag: isinstance(tag, Comment)]):
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
            elif tag.name == 'h3' and self.current_body:
                process_h3(self, tag)
            elif tag.name == 'hr':
                self.forced_layout = None
                process_hr(self, tag)
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

        if self.current_slide:
            from utils import auto_shrink_text
            auto_shrink_text(self.current_slide)
            
        # スライド番号の自動挿入 (Task 4)
        if self.slides_conf.get('show_slide_number', True):
            from pptx.enum.text import PP_ALIGN
            from pptx.util import Pt
            from pptx.dml.color import RGBColor
            
            for i, slide in enumerate(self.prs.slides):
                if i == 0: continue # タイトルスライドは除外
                
                left = self.prs.slide_width - Inches(1.0)
                top = self.prs.slide_height - Inches(0.5)
                width = Inches(0.8)
                height = Inches(0.3)
                
                txBox = slide.shapes.add_textbox(left, top, width, height)
                p = txBox.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.RIGHT
                run = p.add_run()
                run.text = str(i)
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(128, 128, 128)
            
        self.prs.save(output_file)
