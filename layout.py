"""スライドサイズから各要素の配置寸法を導出する

従来は 16:9（10 × 5.625インチ）を前提とした絶対値が processors / utils に散在しており、
config.yaml で `4:3` や `A4` を選んでも本文の高さ以外は追従しなかった。
ここでスライドサイズから寸法を導出し、どの画角でも同じ見た目になるようにする。

各係数は **16:9 のときに従来とまったく同じ寸法になるよう較正**してある
（テスト TestSlideLayout で 16:9 の実測値との一致を検証している）。
"""

from __future__ import annotations

from dataclasses import dataclass

from pptx.util import Emu, Inches, Length

# --- 余白（絶対値）---
# 余白は「紙面が大きくなっても広げない」ほうが自然なため、割合ではなく絶対値で持つ。

#: 左右の余白
SIDE_MARGIN_INCHES = 1.0
#: コンテンツ領域の上端（タイトルプレースホルダーの下端に相当）
CONTENT_TOP_INCHES = 1.5
#: コンテンツ領域の下端余白
BOTTOM_MARGIN_INCHES = 0.325
#: 本文プレースホルダーの下端余白
BODY_BOTTOM_MARGIN_INCHES = 0.5
#: 2カラム時、図の右端に残す余白
SPLIT_RIGHT_MARGIN_INCHES = 0.3

# --- 分割比（スライド幅・コンテンツ領域に対する割合）---

#: 2カラム時の本文枠の幅
SPLIT_BODY_WIDTH_RATIO = 0.48
#: 2カラム時の図の左端
SPLIT_IMAGE_LEFT_RATIO = 0.52

#: コードブロックを右に置くときの本文枠の幅
CODE_SPLIT_BODY_WIDTH_RATIO = 0.45
#: 同、コード枠の左端
CODE_SPLIT_LEFT_RATIO = 0.50
#: 中央寄せレイアウト時のコード枠の左端と幅
CODE_CENTER_LEFT_RATIO = 0.15
CODE_CENTER_WIDTH_RATIO = 0.70
#: 全幅レイアウト時、コード枠をコンテンツ領域より少し内側に置くための調整値
CODE_FULL_TOP_OFFSET_INCHES = 0.5
CODE_FULL_HEIGHT_REDUCTION_INCHES = 0.8

#: テキストと表を上下に分割するときの比率（コンテンツ領域の高さに対する割合）
TABLE_SPLIT_TOP_RATIO = 1.3 / 3.8
TABLE_SPLIT_BODY_HEIGHT_RATIO = 2.0 / 3.8
#: 表を生成するときの初期の高さ（行数に応じてPowerPoint側で伸びる）
TABLE_ROW_HEIGHT_INCHES = 0.8


@dataclass(frozen=True)
class SlideLayout:
    """スライドサイズに応じた配置寸法（すべてEMU）"""

    width: Length
    height: Length

    @classmethod
    def from_presentation(cls, prs) -> SlideLayout:
        return cls(Emu(int(prs.slide_width)), Emu(int(prs.slide_height)))

    # --- コンテンツ領域（画像・表を置ける範囲）---

    @property
    def content_left(self) -> Length:
        return Inches(SIDE_MARGIN_INCHES)

    @property
    def content_top(self) -> Length:
        return Inches(CONTENT_TOP_INCHES)

    @property
    def content_width(self) -> Length:
        return Emu(int(self.width - Inches(SIDE_MARGIN_INCHES) * 2))

    @property
    def content_height(self) -> Length:
        return Emu(int(self.height - self.content_top - Inches(BOTTOM_MARGIN_INCHES)))

    # --- 本文プレースホルダー ---

    def body_height_for(self, top: Length) -> Length:
        """本文枠がスライド下端からはみ出さない高さ"""
        return Emu(int(self.height - top - Inches(BODY_BOTTOM_MARGIN_INCHES)))

    @property
    def split_body_width(self) -> Length:
        """図・表と横に並べるときの本文枠の幅"""
        return Emu(int(self.width * SPLIT_BODY_WIDTH_RATIO))

    @property
    def code_split_body_width(self) -> Length:
        """コード枠と横に並べるときの本文枠の幅"""
        return Emu(int(self.width * CODE_SPLIT_BODY_WIDTH_RATIO))

    # --- 図（画像・Mermaid）---

    @property
    def split_image_left(self) -> Length:
        return Emu(int(self.width * SPLIT_IMAGE_LEFT_RATIO))

    @property
    def split_image_width(self) -> Length:
        return Emu(
            int(self.width - Inches(SPLIT_RIGHT_MARGIN_INCHES) - self.split_image_left)
        )

    # --- 表 ---

    @property
    def table_split_top(self) -> Length:
        """テキストの下に表を置くときの表の上端"""
        return Emu(int(self.content_top + self.content_height * TABLE_SPLIT_TOP_RATIO))

    @property
    def table_split_body_height(self) -> Length:
        """表と上下に並べるときの本文枠の高さ"""
        return Emu(int(self.content_height * TABLE_SPLIT_BODY_HEIGHT_RATIO))

    @property
    def table_height(self) -> Length:
        return Inches(TABLE_ROW_HEIGHT_INCHES)

    # --- コードブロック枠 ---

    @property
    def code_split_left(self) -> Length:
        return Emu(int(self.width * CODE_SPLIT_LEFT_RATIO))

    @property
    def code_split_width(self) -> Length:
        return Emu(int(self.width * (1 - CODE_SPLIT_LEFT_RATIO) - Inches(0.5)))

    @property
    def code_center_left(self) -> Length:
        return Emu(int(self.width * CODE_CENTER_LEFT_RATIO))

    @property
    def code_center_width(self) -> Length:
        return Emu(int(self.width * CODE_CENTER_WIDTH_RATIO))

    @property
    def code_full_top(self) -> Length:
        return Emu(int(self.content_top + Inches(CODE_FULL_TOP_OFFSET_INCHES)))

    @property
    def code_full_height(self) -> Length:
        return Emu(
            int(self.content_height - Inches(CODE_FULL_HEIGHT_REDUCTION_INCHES))
        )
