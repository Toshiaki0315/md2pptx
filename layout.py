"""スライドの配置寸法を導出する

配置の基準は**本文プレースホルダー**とする。タイトルも本文もテンプレートが
定めた位置にあるため、画像や表だけが別の余白で配置されると左端が揃わない。
プレースホルダーに合わせることで、社内テンプレートを使う場合も
そのテンプレートの設計どおりに揃う。

プレースホルダーが取得できない場合は、スライドサイズからの既定値にフォールバックする。
"""

from __future__ import annotations

from dataclasses import dataclass

from pptx.util import Emu, Inches, Length

# --- プレースホルダーが取得できない場合の既定値（絶対値） ---
# 余白は「紙面が大きくなっても広げない」ほうが自然なため、割合ではなく絶対値で持つ。

#: 左右の余白
SIDE_MARGIN_INCHES = 0.5
#: コンテンツ領域の上端（タイトル領域の下端に相当）
CONTENT_TOP_INCHES = 1.75

#: コンテンツ領域および本文枠の下端に残す余白
BOTTOM_MARGIN_INCHES = 0.5

# --- 分割 ---

#: 左右に分割するときの、本文と図の間隔
SPLIT_GUTTER_INCHES = 0.3

#: テキストの下に表を置くときの、本文が使える高さの割合
TABLE_SPLIT_BODY_HEIGHT_RATIO = 0.35
#: 同、本文と表の間隔
TABLE_SPLIT_GUTTER_INCHES = 0.15

#: 中央寄せレイアウトで、コンテンツ領域の左右をさらに詰める割合
CENTER_INSET_RATIO = 0.1


@dataclass(frozen=True)
class SlideLayout:
    """スライドに応じた配置寸法（すべてEMU）

    body_* は本文プレースホルダーの位置・大きさ。取得できない場合は None とし、
    スライドサイズからの既定値を用いる。
    """

    width: Length
    height: Length
    body_left: Length | None = None
    body_top: Length | None = None
    body_width: Length | None = None

    @classmethod
    def from_presentation(cls, prs, content_layout=None) -> SlideLayout:
        """スライドサイズと本文プレースホルダーから配置寸法を導く

        content_layout には実際に本文で使うレイアウトを渡す。省略した場合は
        1つ目のマスターの2番目のレイアウトを見る（従来どおりの動作）。
        """
        left = top = body_width = None
        try:
            layout = content_layout if content_layout is not None else prs.slide_layouts[1]
            body = layout.placeholders[1]
            left = Emu(int(body.left))
            top = Emu(int(body.top))
            body_width = Emu(int(body.width))
        except Exception:
            # 本文プレースホルダーを持たないテンプレートでは既定値を使う
            pass

        return cls(
            Emu(int(prs.slide_width)), Emu(int(prs.slide_height)), left, top, body_width
        )

    # --- コンテンツ領域（画像・表・コード枠を置ける範囲） ---

    @property
    def content_left(self) -> Length:
        return self.body_left if self.body_left is not None else Inches(SIDE_MARGIN_INCHES)

    @property
    def content_top(self) -> Length:
        return self.body_top if self.body_top is not None else Inches(CONTENT_TOP_INCHES)

    @property
    def content_width(self) -> Length:
        """コンテンツ領域の幅（左余白を右側にも同じだけ取る）

        プレースホルダーの幅をそのまま使わないのは、既定テンプレートの
        プレースホルダーがスライドサイズを変えても追従しないため。
        左余白だけを基準にすることで、A4など横長の画角でも紙面を使い切る。
        """
        return Emu(int(self.width - self.content_left * 2))

    @property
    def content_height(self) -> Length:
        return Emu(int(self.height - self.content_top - Inches(BOTTOM_MARGIN_INCHES)))

    # --- 本文プレースホルダー ---

    def body_height_for(self, top: Length) -> Length:
        """本文枠がスライド下端からはみ出さない高さ"""
        return Emu(int(self.height - top - Inches(BOTTOM_MARGIN_INCHES)))

    # --- 左右分割（本文と、図またはコード枠） ---

    @property
    def split_body_width(self) -> Length:
        """図・表と横に並べるときの本文枠の幅（コンテンツ領域を等分する）"""
        return Emu(int((self.content_width - Inches(SPLIT_GUTTER_INCHES)) / 2))

    @property
    def split_right_left(self) -> Length:
        """左右分割で右側に置く要素の左端"""
        return Emu(int(self.content_left + self.split_body_width + Inches(SPLIT_GUTTER_INCHES)))

    @property
    def split_right_width(self) -> Length:
        """同、右側に置く要素の幅（右端はコンテンツ領域の右端に揃う）"""
        return Emu(int(self.content_left + self.content_width - self.split_right_left))

    # --- 上下分割（テキストと表） ---

    @property
    def table_split_body_height(self) -> Length:
        """表と上下に並べるときの本文枠の高さ"""
        return Emu(int(self.content_height * TABLE_SPLIT_BODY_HEIGHT_RATIO))

    @property
    def table_split_top(self) -> Length:
        """テキストの下に表を置くときの表の上端（本文枠と重ならない位置）"""
        return Emu(int(
            self.content_top + self.table_split_body_height + Inches(TABLE_SPLIT_GUTTER_INCHES)
        ))

    # --- 中央寄せ ---

    @property
    def center_left(self) -> Length:
        return Emu(int(self.content_left + self.content_width * CENTER_INSET_RATIO))

    @property
    def center_width(self) -> Length:
        return Emu(int(self.content_width * (1 - CENTER_INSET_RATIO * 2)))
