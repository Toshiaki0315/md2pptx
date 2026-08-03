"""md2pptx のユニットテスト・結合テスト"""

import base64
import os
from io import BytesIO
from unittest.mock import MagicMock, patch

import markdown
import pytest
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

import md2pptx
from md2pptx import LAYOUT, PPTXGenerator, load_config, main, parse_args, read_text_file

# 1x1ピクセルの最小PNG（画像挿入テスト用のダミーデータ）
TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAE"
    "hQGAhKmMIQAAAABJRU5ErkJggg=="
)


# --- フィクスチャ（テスト用の共通設定） ---


@pytest.fixture
def base_config():
    """テスト用の基本的なYAML設定のモック"""
    return {
        "slides": {"layout": "16:9"},
        "fonts": {
            "title_h1": {"name": "Meiryo", "size_pt": 44, "bold": True},
            "body": {"name": "Meiryo", "size_pt": 20},
            "inline_code": {"name": "Consolas", "color_rgb": [220, 20, 60]},
        },
        "images": {"default_height_inches": 3.5},
    }


@pytest.fixture
def gen(base_config):
    """スライド未作成のジェネレーター"""
    return PPTXGenerator(base_config)


@pytest.fixture
def gen_with_slide(gen):
    """コンテンツスライドを1枚作成済みのジェネレーター"""
    gen._process_heading(parse_md("## 見出し").find("h2"))
    return gen


@pytest.fixture
def png_file(tmp_path):
    """ローカル画像ファイルのパス"""
    path = tmp_path / "tiny.png"
    path.write_bytes(TINY_PNG)
    return str(path)


@pytest.fixture
def mock_response():
    """requests.get の戻り値のモック"""
    response = MagicMock()
    response.content = TINY_PNG
    response.status_code = 200
    response.raise_for_status.return_value = None
    return response


def parse_md(md_text):
    """Markdown文字列をBeautifulSoupに変換するヘルパー"""
    html = markdown.markdown(md_text, extensions=["extra", "fenced_code", "sane_lists"])
    return BeautifulSoup(html, "html.parser")


def parse_html(html):
    """HTML文字列をBeautifulSoupに変換するヘルパー"""
    return BeautifulSoup(html, "html.parser")


def new_paragraph():
    """スタイル検証用の実物のparagraphオブジェクトを作る"""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    return box.text_frame.paragraphs[0]


# --- 初期化・設定まわり ---


class TestInitialization:
    @pytest.mark.parametrize(
        "layout, expected",
        [
            ("16:9", (Inches(10), Inches(5.625))),
            ("4:3", (Inches(10), Inches(7.5))),
            ("16:10", (Inches(10), Inches(6.25))),
            ("A4", (Inches(11.69), Inches(8.27))),
            ("unknown", (Inches(10), Inches(5.625))),
        ],
    )
    def test_get_slide_size(self, gen, layout, expected):
        """スライドサイズの計算ロジック（未知の値は16:9にフォールバック）"""
        assert gen._get_slide_size(layout) == expected

    def test_slide_size_applied_from_config(self, base_config):
        """config.yamlの画角がプレゼンテーションに反映される"""
        base_config["slides"]["layout"] = "4:3"
        gen = PPTXGenerator(base_config)
        assert gen.prs.slide_width == Inches(10)
        assert gen.prs.slide_height == Inches(7.5)

    def test_empty_config_is_accepted(self):
        """設定が空・Noneでも既定値で初期化できる"""
        for config in (None, {}, {"slides": None, "fonts": None, "images": None}):
            gen = PPTXGenerator(config)
            assert gen.prs.slide_height == Inches(5.625)

    def test_template_is_loaded_when_exists(self, base_config, tmp_path):
        """テンプレートが存在する場合は読み込み、画角設定を上書きしない"""
        template = tmp_path / "template.pptx"
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        prs.save(str(template))

        base_config["slides"]["template_path"] = str(template)
        gen = PPTXGenerator(base_config)
        assert gen.prs.slide_width == Inches(13.333)

    def test_missing_template_falls_back_to_default(self, base_config):
        """テンプレートのパスが不正な場合はデフォルトのプレゼンテーションを使う"""
        base_config["slides"]["template_path"] = "no_such_template.pptx"
        gen = PPTXGenerator(base_config)
        assert gen.prs.slide_height == Inches(5.625)


class TestFontResolution:
    def test_builtin_default_is_used_when_unset(self, gen):
        """config.yamlに無いキーは組み込みのデフォルト値で補完される"""
        assert gen._font("table_header") == {
            "name": "Meiryo",
            "size_pt": 14,
            "bold": True,
            "color_rgb": [255, 255, 255],
        }

    def test_user_config_overrides_per_key(self, base_config):
        """ユーザー設定はキー単位でデフォルトを上書きする"""
        base_config["fonts"]["code_block"] = {"size_pt": 30}
        gen = PPTXGenerator(base_config)
        conf = gen._font("code_block")
        assert conf["size_pt"] == 30
        assert conf["name"] == "Consolas"
        assert conf["color_rgb"] == [0, 80, 160]

    def test_unknown_key_returns_empty(self, gen):
        """定義の無いキーは空の設定（＝スタイル未適用）になる"""
        assert gen._font("nothing") == {}


class TestApplyFontStyle:
    def test_all_properties_applied(self):
        """name / size_pt / bold / color_rgb がすべて反映される"""
        run = new_paragraph().add_run()
        PPTXGenerator.apply_font_style(
            run, {"name": "Meiryo", "size_pt": 24, "bold": True, "color_rgb": [1, 2, 3]}
        )
        assert run.font.name == "Meiryo"
        assert run.font.size == Pt(24)
        assert run.font.bold is True
        assert run.font.color.rgb == RGBColor(1, 2, 3)

    def test_partial_config_leaves_others_untouched(self):
        """指定されていないプロパティは変更されない"""
        run = new_paragraph().add_run()
        PPTXGenerator.apply_font_style(run, {"name": "Meiryo"})
        assert run.font.name == "Meiryo"
        assert run.font.size is None
        assert run.font.bold is None

    @pytest.mark.parametrize("config", [None, {}])
    def test_empty_config_is_noop(self, config):
        """設定が空の場合は何も適用しない"""
        run = new_paragraph().add_run()
        PPTXGenerator.apply_font_style(run, config)
        assert run.font.name is None


# --- 画像配置 ---


class TestInsertImageFit:
    def _mock_slide(self, width, height):
        mock_slide = MagicMock()
        mock_pic = MagicMock()
        mock_pic.width, mock_pic.height = width, height
        mock_slide.shapes.add_picture.return_value = mock_pic
        return mock_slide

    def test_small_image_is_scaled_up_to_cap(self):
        """小さい画像は最大1.5倍までしか拡大されない"""
        slide = self._mock_slide(100, 200)
        pic = PPTXGenerator.insert_image_fit(slide, b"dummy", 0, 0, 500, 500)
        assert (pic.width, pic.height) == (150, 300)

    def test_large_image_is_shrunk_to_fit(self):
        """大きい画像はアスペクト比を保って枠内に縮小される"""
        slide = self._mock_slide(1000, 500)
        pic = PPTXGenerator.insert_image_fit(slide, b"dummy", 0, 0, 500, 500)
        assert (pic.width, pic.height) == (500, 250)

    def test_image_is_centered_in_frame(self):
        """縮小後の画像は枠の中央に配置される"""
        slide = self._mock_slide(1000, 500)
        pic = PPTXGenerator.insert_image_fit(slide, b"dummy", 100, 200, 500, 500)
        assert pic.left == 100 + (500 - 500) / 2
        assert pic.top == 200 + (500 - 250) / 2


class TestPlaceImage:
    @patch("md2pptx.PPTXGenerator.insert_image_fit")
    def test_full_width_when_slide_has_no_text(self, mock_fit, gen_with_slide):
        """テキストが無いスライドでは画像を中央に大きく配置する"""
        gen_with_slide._place_image(BytesIO(TINY_PNG))
        args = mock_fit.call_args[0]
        assert args[2] == Inches(LAYOUT.full_left)
        assert args[4] == Inches(LAYOUT.full_width)

    @patch("md2pptx.PPTXGenerator.insert_image_fit")
    def test_two_column_when_slide_has_text(self, mock_fit, gen_with_slide):
        """テキストがあるスライドでは本文枠を縮めて右側に配置する"""
        gen_with_slide.slide_has_text = True
        gen_with_slide._place_image(BytesIO(TINY_PNG))

        args = mock_fit.call_args[0]
        assert args[2] == Inches(LAYOUT.split_image_left)
        assert args[4] == Inches(LAYOUT.split_image_width)
        assert gen_with_slide.current_body_shape.width == Inches(LAYOUT.split_body_width)

    def test_shrink_without_body_shape_is_noop(self, gen):
        """本文枠が未設定でも例外にならない"""
        gen._shrink_body_shape(4.8)

    def test_shrink_keeps_inherited_geometry(self, gen_with_slide):
        """幅だけを変更してもレイアウトから継承した位置・高さが失われない"""
        shape = gen_with_slide.current_body_shape
        top, height = shape.top, shape.height

        gen_with_slide._shrink_body_shape(LAYOUT.split_body_width)

        assert shape.width == Inches(LAYOUT.split_body_width)
        assert shape.left == Inches(0.5)  # 既定レイアウトの本文枠の左位置
        assert (shape.top, shape.height) == (top, height)

    def test_heading_keeps_inherited_width(self, gen_with_slide):
        """高さ補正をしても本文枠の幅が0に落ちない（継承値のリセット防止）"""
        assert gen_with_slide.current_body_shape.width > 0


# --- インライン装飾 ---


class TestAddRunsFromTag:
    def _runs_for(self, gen, html):
        p = new_paragraph()
        gen._add_runs_from_tag(parse_html(html).p, p, {"name": "Meiryo"})
        return p.runs

    def test_bold_and_italic(self, gen):
        """strong/em がフォントスタイルに変換される"""
        runs = self._runs_for(gen, "<p><strong>太字</strong><em>斜体</em></p>")
        assert [r.text for r in runs] == ["太字", "斜体"]
        assert runs[0].font.bold is True
        assert runs[1].font.italic is True

    def test_inline_code_uses_code_font(self, gen):
        """codeタグにはinline_codeの設定が適用される"""
        (run,) = self._runs_for(gen, "<p><code>x = 1</code></p>")
        assert run.font.name == "Consolas"
        assert run.font.color.rgb == RGBColor(220, 20, 60)

    def test_plain_text_uses_default_font(self, gen):
        """装飾の無いテキストには既定フォントが適用される"""
        (run,) = self._runs_for(gen, "<p>ふつうの文字</p>")
        assert run.font.name == "Meiryo"

    def test_nested_container_is_flattened(self, gen):
        """spanなどのコンテナ要素は再帰的に展開される"""
        runs = self._runs_for(gen, "<p><span>外<strong>中</strong></span></p>")
        assert [r.text for r in runs] == ["外", "中"]
        assert runs[1].font.bold is True

    def test_block_tags_are_skipped(self, gen):
        """リストや表などのブロック要素はインライン処理から除外される"""
        runs = self._runs_for(gen, "<p>本文<ul><li>項目</li></ul><table></table></p>")
        assert [r.text for r in runs] == ["本文"]

    def test_newlines_are_replaced_and_blanks_skipped(self, gen):
        """改行は半角スペースに置換され、空白のみのテキストは無視される"""
        runs = self._runs_for(gen, "<p>1行目\n2行目<b>\n\n</b></p>")
        assert runs[0].text == "1行目 2行目"


# --- タグごとの処理 ---


class TestProcessHeading:
    def test_h1_creates_title_slide(self, gen):
        """h1はタイトルスライド（レイアウト0）を作る"""
        gen._process_heading(parse_md("# タイトル").find("h1"))
        assert len(gen.prs.slides) == 1
        assert gen.current_slide.shapes.title.text == "タイトル"
        assert gen.current_slide.slide_layout == gen.prs.slide_layouts[0]

    def test_h2_creates_content_slide(self, gen):
        """h2はコンテンツスライド（レイアウト1）を作る"""
        gen._process_heading(parse_md("## 中身").find("h2"))
        assert gen.current_slide.slide_layout == gen.prs.slide_layouts[1]

    def test_title_font_is_applied(self, gen):
        """title_h1 の設定がタイトルに適用される"""
        gen._process_heading(parse_md("# タイトル").find("h1"))
        run = gen.current_slide.shapes.title.text_frame.paragraphs[0].runs[0]
        assert run.font.size == Pt(44)
        assert run.font.bold is True

    def test_generic_title_config_is_fallback(self, base_config):
        """title_h2 が無い場合は title の設定にフォールバックする"""
        base_config["fonts"]["title"] = {"size_pt": 18}
        gen = PPTXGenerator(base_config)
        gen._process_heading(parse_md("## 中身").find("h2"))
        run = gen.current_slide.shapes.title.text_frame.paragraphs[0].runs[0]
        assert run.font.size == Pt(18)

    def test_state_is_reset_for_new_slide(self, gen):
        """スライド作成時に本文とテキスト有無フラグが初期化される"""
        gen.slide_has_text = True
        gen._process_heading(parse_md("## 中身").find("h2"))
        assert gen.slide_has_text is False
        assert gen.current_body.text == ""
        assert (
            gen.current_body_shape.shape_id
            == gen.current_slide.placeholders[1].shape_id
        )

    def test_body_height_is_clamped_to_slide(self, gen):
        """デフォルトテンプレートでは本文枠がスライド内に収まるよう補正される"""
        gen._process_heading(parse_md("## 中身").find("h2"))
        shape = gen.current_body_shape
        assert shape.top + shape.height == gen.prs.slide_height - Inches(
            LAYOUT.body_bottom_margin
        )

    def test_template_layout_is_preserved(self, base_config, tmp_path):
        """テンプレート利用時は本文枠の高さを変更しない"""
        template = tmp_path / "template.pptx"
        Presentation().save(str(template))
        base_config["slides"]["template_path"] = str(template)

        gen = PPTXGenerator(base_config)
        original = gen.prs.slide_layouts[1].placeholders[1].height
        gen._process_heading(parse_md("## 中身").find("h2"))
        assert gen.current_body_shape.height == original


class TestProcessBlockquote:
    def test_note_is_written(self, gen_with_slide):
        """引用ブロックはスピーカーノートに書き込まれる"""
        gen_with_slide._process_blockquote(parse_md("> メモ").find("blockquote"))
        notes = gen_with_slide.current_slide.notes_slide.notes_text_frame.text
        assert notes == "メモ"

    def test_multiple_notes_are_appended(self, gen_with_slide):
        """複数の引用ブロックは空行を挟んで追記される"""
        gen_with_slide._process_blockquote(parse_md("> 1つ目").find("blockquote"))
        gen_with_slide._process_blockquote(parse_md("> 2つ目").find("blockquote"))
        notes = gen_with_slide.current_slide.notes_slide.notes_text_frame.text
        assert notes == "1つ目\n\n2つ目"


class TestProcessImage:
    def test_local_path_is_inserted(self, gen_with_slide, png_file):
        """ローカル画像はダウンロードせずにそのまま挿入される"""
        gen_with_slide._process_image(parse_html(f'<img src="{png_file}">').img)
        assert len(gen_with_slide.current_slide.shapes) == 3  # title, body, picture

    @patch("requests.get")
    def test_remote_image_is_downloaded_with_timeout(
        self, mock_get, gen_with_slide, mock_response
    ):
        """URL画像はタイムアウト付きで取得される"""
        mock_get.return_value = mock_response
        gen_with_slide._process_image(
            parse_html('<img src="http://example.com/a.png">').img
        )
        mock_get.assert_called_once_with(
            "http://example.com/a.png", timeout=md2pptx.HTTP_TIMEOUT_SEC
        )
        mock_response.raise_for_status.assert_called_once()

    @patch("md2pptx.PPTXGenerator.insert_image_fit")
    def test_fixed_position_from_config(self, mock_fit, base_config, png_file):
        """position_inches が指定されている場合はオートレイアウトを使わない"""
        base_config["images"]["position_inches"] = [5.2, 1.8]
        gen = PPTXGenerator(base_config)
        gen._process_heading(parse_md("## 中身").find("h2"))
        gen.current_slide = MagicMock()

        gen._process_image(parse_html(f'<img src="{png_file}">').img)

        mock_fit.assert_not_called()
        gen.current_slide.shapes.add_picture.assert_called_once_with(
            png_file, Inches(5.2), Inches(1.8), height=Inches(3.5)
        )

    def test_missing_src_is_skipped(self, gen_with_slide, capsys):
        """src属性が無い画像は警告を出してスキップする"""
        gen_with_slide._process_image(parse_html("<img>").img)
        assert "Warning" in capsys.readouterr().out
        assert len(gen_with_slide.current_slide.shapes) == 2

    @patch("requests.get", side_effect=OSError("network down"))
    def test_download_failure_is_reported(self, _mock_get, gen_with_slide, capsys):
        """取得に失敗しても処理は継続し、警告のみ表示する"""
        gen_with_slide._process_image(
            parse_html('<img src="https://example.com/a.png">').img
        )
        assert "画像の挿入に失敗しました" in capsys.readouterr().out


class TestProcessTable:
    MD_TABLE = "| 列A | 列B |\n|---|---|\n| 値1 | 値2 |"

    def _table_of(self, gen):
        return [s for s in gen.current_slide.shapes if s.has_table][0].table

    def test_rows_and_cells_are_converted(self, gen_with_slide):
        """行数・列数とセルのテキストが再現される"""
        gen_with_slide._process_table(parse_md(self.MD_TABLE).find("table"))
        table = self._table_of(gen_with_slide)
        assert (len(table.rows), len(table.columns)) == (2, 2)
        assert table.cell(0, 0).text == "列A"
        assert table.cell(1, 1).text == "値2"
        assert gen_with_slide.slide_has_text is True

    def test_header_and_body_fonts(self, gen_with_slide):
        """ヘッダー行とボディ行で異なるフォント設定が適用される"""
        gen_with_slide._process_table(parse_md(self.MD_TABLE).find("table"))
        table = self._table_of(gen_with_slide)
        header_run = table.cell(0, 0).text_frame.paragraphs[0].runs[0]
        body_run = table.cell(1, 0).text_frame.paragraphs[0].runs[0]
        assert header_run.font.bold is True
        assert header_run.font.size == Pt(14)
        assert body_run.font.size == Pt(12)

    def test_inline_decoration_in_cell(self, gen_with_slide):
        """セル内の太字・インラインコードも反映される"""
        md = "| 見出し |\n|---|\n| **強調** |"
        gen_with_slide._process_table(parse_md(md).find("table"))
        run = self._table_of(gen_with_slide).cell(1, 0).text_frame.paragraphs[0].runs[0]
        assert run.font.bold is True

    def test_column_count_uses_widest_row(self, gen_with_slide):
        """列数が揃っていない場合は最大列数に合わせる"""
        html = "<table><tr><td>a</td></tr><tr><td>b</td><td>c</td></tr></table>"
        gen_with_slide._process_table(parse_html(html).table)
        assert len(self._table_of(gen_with_slide).columns) == 2

    def test_short_row_leaves_empty_cells(self, gen_with_slide):
        """セルが足りない行の余白セルは空のまま保持される"""
        html = "<table><tr><td>a</td></tr><tr><td>b</td><td>c</td></tr></table>"
        gen_with_slide._process_table(parse_html(html).table)
        assert self._table_of(gen_with_slide).cell(0, 1).text == ""

    @pytest.mark.parametrize(
        "html", ["<table></table>", "<table><tr></tr></table>"]
    )
    def test_empty_table_is_skipped(self, gen_with_slide, html):
        """行やセルが無い表は何も生成しない"""
        gen_with_slide._process_table(parse_html(html).table)
        assert len(gen_with_slide.current_slide.shapes) == 2

    def test_layout_is_split_when_slide_has_text(self, gen_with_slide):
        """テキストがある場合は本文枠を縮めて表を下半分に置く"""
        gen_with_slide.slide_has_text = True
        gen_with_slide._process_table(parse_md(self.MD_TABLE).find("table"))

        shape = [s for s in gen_with_slide.current_slide.shapes if s.has_table][0]
        assert shape.top == Inches(LAYOUT.table_split_top)
        assert gen_with_slide.current_body_shape.height == Inches(
            LAYOUT.table_split_body_height
        )

    def test_layout_is_full_when_slide_is_empty(self, gen_with_slide):
        """テキストが無い場合は表を上部から配置する"""
        gen_with_slide._process_table(parse_md(self.MD_TABLE).find("table"))
        shape = [s for s in gen_with_slide.current_slide.shapes if s.has_table][0]
        assert shape.top == Inches(LAYOUT.content_top)


class TestCodeAndMermaid:
    def test_code_block_is_written_as_text(self, gen_with_slide):
        """通常のコードブロックは等幅フォントの本文として書き込まれる"""
        md = "```python\nprint(1)\n```"
        gen_with_slide._process_code_or_mermaid(parse_md(md).find("pre"))

        run = gen_with_slide.current_body.paragraphs[0].runs[0]
        assert "print(1)" in run.text
        assert run.font.name == "Consolas"
        assert run.font.color.rgb == RGBColor(0, 80, 160)
        assert gen_with_slide.slide_has_text is True

    def test_code_block_without_language(self, gen_with_slide):
        """言語指定の無いコードブロックもMermaid扱いにならない"""
        gen_with_slide._process_code_or_mermaid(parse_html("<pre><code>x</code></pre>"))
        assert gen_with_slide.current_body.paragraphs[0].runs[0].text == "x"

    @pytest.mark.parametrize(
        "html, expected",
        [
            ('<code class="language-mermaid">g</code>', True),
            ('<code class="mermaid">g</code>', True),
            ('<code class="language-python">g</code>', False),
            ("<code>g</code>", False),
        ],
    )
    def test_is_mermaid(self, html, expected):
        """Mermaid判定はクラス属性で行う（class無しでも例外にしない）"""
        assert PPTXGenerator._is_mermaid(parse_html(html).code) is expected

    def test_is_mermaid_without_code_tag(self):
        """codeタグが無い場合はMermaidではない"""
        assert PPTXGenerator._is_mermaid(None) is False

    @patch("requests.get")
    def test_mermaid_is_requested_from_kroki(self, mock_get, mock_response):
        """Mermaid記法は圧縮・base64化してKrokiに問い合わせる"""
        mock_get.return_value = mock_response
        result = PPTXGenerator._render_mermaid("graph TD; A-->B;")

        url = mock_get.call_args[0][0]
        assert url.startswith(md2pptx.KROKI_MERMAID_PNG_URL)
        assert mock_get.call_args[1]["timeout"] == md2pptx.HTTP_TIMEOUT_SEC
        assert result.getvalue() == TINY_PNG

    @patch("md2pptx.PPTXGenerator._place_image")
    @patch("requests.get")
    def test_mermaid_image_is_placed(
        self, mock_get, mock_place, gen_with_slide, mock_response
    ):
        """生成されたMermaid画像はオートレイアウトで配置される"""
        mock_get.return_value = mock_response
        md = "```mermaid\ngraph TD; A-->B;\n```"
        gen_with_slide._process_code_or_mermaid(parse_md(md).find("pre"))
        mock_place.assert_called_once()

    @patch("requests.get", side_effect=OSError("timeout"))
    def test_mermaid_failure_is_reported(self, _mock_get, gen_with_slide, capsys):
        """API失敗時は警告のみ表示して処理を継続する"""
        md = "```mermaid\ngraph TD; A-->B;\n```"
        gen_with_slide._process_code_or_mermaid(parse_md(md).find("pre"))
        assert "Mermaid図形の生成に失敗しました" in capsys.readouterr().out


class TestProcessText:
    def test_paragraph_uses_body_font(self, gen_with_slide):
        """段落にはbodyのフォント設定が適用される"""
        gen_with_slide._process_text(parse_md("本文です").find("p"))
        run = gen_with_slide.current_body.paragraphs[0].runs[0]
        assert run.text == "本文です"
        assert run.font.size == Pt(20)
        assert gen_with_slide.slide_has_text is True

    def test_empty_text_is_skipped(self, gen_with_slide):
        """空の要素は段落を追加しない"""
        gen_with_slide._process_text(parse_html("<p>   </p>").p)
        assert gen_with_slide.slide_has_text is False
        assert len(gen_with_slide.current_body.paragraphs) == 1

    def test_bullet_level_follows_nesting(self, gen_with_slide):
        """ネストの深さがリストのレベルになる"""
        md = "* 親\n    * 子\n        * 孫"
        items = parse_md(md).find_all("li")
        for item in items:
            gen_with_slide._process_text(item)

        levels = [p.level for p in gen_with_slide.current_body.paragraphs]
        assert levels == [0, 1, 2]

    def test_bullet_level_is_capped(self, gen_with_slide):
        """レベルはPowerPointの上限(8)でクランプされる"""
        html = "<ul><li>深い</li></ul>"
        for _ in range(12):
            html = f"<ul><li>{html}</li></ul>"
        deepest = parse_html(html).find_all("li")[-1]

        gen_with_slide._process_text(deepest)
        assert gen_with_slide.current_body.paragraphs[0].level == md2pptx.MAX_BULLET_LEVEL

    def test_bullet_level_font_config(self, base_config):
        """bullet_level_N の設定がレベルごとに適用される"""
        base_config["fonts"]["bullet_level_1"] = {"size_pt": 18}
        gen = PPTXGenerator(base_config)
        gen._process_heading(parse_md("## 見出し").find("h2"))
        gen._process_text(parse_md("* 項目").find("li"))
        assert gen.current_body.paragraphs[0].runs[0].font.size == Pt(18)

    def test_first_paragraph_is_reused_then_appended(self, gen_with_slide):
        """最初の段落は空の既存段落を再利用し、以降は追加される"""
        gen_with_slide._process_text(parse_md("1つ目").find("p"))
        assert len(gen_with_slide.current_body.paragraphs) == 1

        gen_with_slide._process_text(parse_md("2つ目").find("p"))
        assert len(gen_with_slide.current_body.paragraphs) == 2


# --- 結合テスト ---


class TestGenerate:
    @patch("md2pptx.PPTXGenerator.insert_image_fit")
    @patch("requests.get")
    def test_markdown_integration(
        self, mock_get, mock_insert_image, base_config, mock_response, tmp_path
    ):
        """Markdownのパースからスライド生成までの一連の結合テスト"""
        mock_get.return_value = mock_response

        md_content = """
# タイトルスライド
ここはタイトルです。

## テキストと装飾のテスト
* 箇条書き1
* **太字** と `インラインコード`

> これはスピーカーノートです。

## 画像のテスト
![テスト画像](http://example.com/test.png)

## 表のテスト
| 列A | 列B |
|---|---|
| 値1 | 値2 |

## Mermaidのテスト
```mermaid
graph TD; A-->B;
```
        """

        gen = PPTXGenerator(base_config)
        output_path = tmp_path / "test_output.pptx"
        gen.generate(md_content, str(output_path))

        # ファイルが生成され、h1/h2の数だけスライドが作られる
        assert os.path.exists(output_path)
        assert len(gen.prs.slides) == 5

        # ノートが正しいスライドに追加される
        notes = gen.prs.slides[1].notes_slide.notes_text_frame.text
        assert "これはスピーカーノートです。" in notes

        # 外部通信と画像挿入は 画像1回 + Mermaid1回
        assert mock_get.call_count == 2
        assert mock_insert_image.call_count == 2

    def test_content_before_first_heading_is_ignored(self, gen, tmp_path):
        """見出しより前の本文は配置先が無いため無視される"""
        gen.generate("本文だけ\n\n# タイトル\n中身", str(tmp_path / "out.pptx"))
        assert len(gen.prs.slides) == 1
        assert gen.current_body.text == "中身"

    def test_list_items_are_converted(self, gen, tmp_path):
        """箇条書きは1項目につき1段落になる"""
        gen.generate("# タイトル\n\n* 項目1\n* 項目2\n", str(tmp_path / "out.pptx"))
        texts = [p.text for p in gen.current_body.paragraphs]
        assert texts == ["項目1", "項目2"]

    def test_list_item_paragraph_is_not_duplicated(self, gen, tmp_path):
        """段落を含むリスト（loose list）でもpタグが二重に出力されない"""
        md = "# タイトル\n\n* 項目1\n\n* 項目2\n"
        gen.generate(md, str(tmp_path / "out.pptx"))
        texts = [p.text.strip() for p in gen.current_body.paragraphs]
        assert texts == ["項目1", "項目2"]

    def test_blockquote_paragraph_is_not_duplicated(self, gen, tmp_path):
        """引用内のpタグは本文に出力されない"""
        gen.generate("# タイトル\n\n> ノート\n", str(tmp_path / "out.pptx"))
        assert gen.current_body.text == ""
        assert gen.current_slide.notes_slide.notes_text_frame.text == "ノート"


# --- CLI ---


class TestCli:
    def test_parse_args_defaults(self):
        """出力先と設定ファイルの既定値"""
        args = parse_args(["input.md"])
        assert args.input == "input.md"
        assert args.output == "output.pptx"
        assert args.config == "config.yaml"

    def test_parse_args_options(self):
        """オプション指定が反映される"""
        args = parse_args(["in.md", "-o", "out.pptx", "-c", "my.yaml"])
        assert (args.output, args.config) == ("out.pptx", "my.yaml")

    def test_load_config(self, tmp_path):
        """YAML設定ファイルを辞書として読み込む"""
        path = tmp_path / "c.yaml"
        path.write_text("slides:\n  layout: '4:3'\n", encoding="utf-8")
        assert load_config(str(path)) == {"slides": {"layout": "4:3"}}

    def test_load_config_empty_file(self, tmp_path):
        """空のYAMLは空の辞書として扱う"""
        path = tmp_path / "empty.yaml"
        path.write_text("", encoding="utf-8")
        assert load_config(str(path)) == {}

    def test_read_text_file(self, tmp_path):
        """UTF-8のテキストを読み込む"""
        path = tmp_path / "a.md"
        path.write_text("# 日本語", encoding="utf-8")
        assert read_text_file(str(path)) == "# 日本語"

    def _write_project(self, tmp_path):
        md = tmp_path / "in.md"
        md.write_text("# タイトル\n本文\n", encoding="utf-8")
        conf = tmp_path / "c.yaml"
        conf.write_text("slides:\n  layout: '16:9'\n", encoding="utf-8")
        return str(md), str(conf), str(tmp_path / "out.pptx")

    def test_main_success(self, tmp_path, capsys):
        """正常系では0を返しファイルを生成する"""
        md, conf, out = self._write_project(tmp_path)
        assert main([md, "-o", out, "-c", conf]) == 0
        assert os.path.exists(out)
        assert "Success" in capsys.readouterr().out

    def test_main_uses_sys_argv_by_default(self, tmp_path, mocker):
        """引数を渡さない場合はsys.argvを解析する"""
        md, conf, out = self._write_project(tmp_path)
        mocker.patch.object(md2pptx.sys, "argv", ["md2pptx.py", md, "-o", out, "-c", conf])
        assert main() == 0

    def test_main_missing_input(self, tmp_path, capsys):
        """入力ファイルが無い場合は1を返す"""
        _, conf, out = self._write_project(tmp_path)
        assert main(["no_such.md", "-o", out, "-c", conf]) == 1
        assert "入力ファイル" in capsys.readouterr().out

    def test_main_missing_config(self, tmp_path, capsys):
        """設定ファイルが無い場合は1を返す"""
        md, _, out = self._write_project(tmp_path)
        assert main([md, "-o", out, "-c", "no_such.yaml"]) == 1
        assert "設定ファイル" in capsys.readouterr().out

    def test_main_permission_error(self, tmp_path, mocker, capsys):
        """出力先に書き込めない場合は専用のメッセージを表示する"""
        md, conf, out = self._write_project(tmp_path)
        mocker.patch.object(
            md2pptx.PPTXGenerator, "generate", side_effect=PermissionError
        )
        assert main([md, "-o", out, "-c", conf]) == 1
        assert "書き込めません" in capsys.readouterr().out

    def test_main_unexpected_error(self, tmp_path, mocker, capsys):
        """想定外の例外はメッセージとトレースバックを表示して1を返す"""
        md, conf, out = self._write_project(tmp_path)
        mocker.patch.object(
            md2pptx.PPTXGenerator, "generate", side_effect=ValueError("boom")
        )
        assert main([md, "-o", out, "-c", conf]) == 1
        captured = capsys.readouterr()
        assert "予期せぬエラー" in captured.out
        assert "ValueError" in captured.err
