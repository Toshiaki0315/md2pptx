# Markdown to PowerPoint Converter

## 概要

Markdown形式（`.md`）で書かれたプレーンテキストのドキュメントを読み込み、PowerPointプレゼンテーション（`.pptx`）ファイルへ自動変換するPythonスクリプトです。
見出し（h1, h2）を新しいスライドのタイトルに、箇条書き（ul, li）を本文のコンテンツとして処理します。
外部のYAML設定ファイルにより、テーマテンプレート、画角（縦横比）、フォントスタイル（名前、サイズ、色）、画像の配置などのデフォルト値を、スクリプト本体を書き換えずにカスタマイズできます。

## 特徴

*   **簡単変換**: Markdownの見出しをスライドタイトル、箇条書きを本文コンテンツに自動割り当て。
*   **YAML設定対応**: テーマ、画角、フォント（タイトル・本文・箇条書き別）、画像の配置ルールをYAMLで外部定義可能。
*   **テンプレート対応**: 既存のPPTXファイル（会社フォーマットなど）をテンプレートとして読み込み、デザインを継承可能。
*   **画像対応**: Markdown内の `<img>` タグ（ローカルファイル・Web上のURL双方）を解析し、スライドへ埋め込み可能。
*   **コマンドライン・インターフェース**: 入力ファイル、出力ファイル名、設定ファイルを引数で指定可能。
*   **安全な環境**: 外部管理環境（PEP 668）に対応するため、仮想環境（venv）を利用した構築方法を推奨。

## 利用環境の構築方法

このプロジェクトはPython 3.11以降の環境（PEP 668による外部管理環境保護）での動作を想定しています。システム環境を汚さないよう、**仮想環境（venv）**の利用を強く推奨します。

### 1. リポジトリのクローン（またはファイルの配置）

プロジェクトファイルを任意のディレクトリに配置します。

```bash
mkdir md2pptx
cd md2pptx
# ここに md2pptx.py, config.yaml, default.yaml, sample.md などを配置
```

### 2. 仮想環境（venv）の作成

プロジェクトディレクトリ内で、独立したPython環境を作成します。

```bash
# Linux / macOS / Windows
python3 -m venv venv
```
*   *※ Linux環境でエラーが出る場合は `sudo apt install python3-venv` で必要なパッケージをインストールしてください。*

### 3. 仮想環境のアクティベート

作成した環境を有効にします。コマンドはOSによって異なります。

*   **Linux / macOS**:
    ```bash
    source venv/bin/activate
    ```
*   **Windows (PowerShell)**:
    ```powershell
    .\venv\Scripts\Activate.ps1
    ```
    *   *※ 実行ポリシーのエラーが出る場合は、管理者権限のPSで `Set-ExecutionPolicy RemoteSigned -Scope Process` を実行してください。*
*   **Windows (コマンドプロンプト)**:
    ```cmd
    .\venv\Scripts\activate.bat
    ```

アクティベートに成功すると、プロンプトの先頭に `(venv)` と表示されます。

### 4. 必要なライブラリのインストール

仮想環境内で、`pip` を使用して依存ライブラリをインストールします。
```bash
# (venv) プロンプトが表示されている状態で実行
pip install python-pptx markdown beautifulsoup4 requests PyYAML
```

*   *※ インストール中に `Requirement already satisfied...` や ROS関連の依存競合エラー（Generate-parameter-library-pyなど）が表示される場合がありますが、本スクリプトの動作には影響ないため無視して構いません。*

## 使用方法

仮想環境がアクティベートされた状態で、`md2pptx.py` を実行します。

### 基本的な使い方

最もシンプルな実行方法です。デフォルトの設定ファイル (`config.yaml`) を使用し、出力ファイル名は `output.pptx` となります。
```bash
# sample.md を変換し、output.pptx を生成
python md2pptx.py sample.md
```

### 高度な使い方（オプション指定）

入力ファイル、出力ファイル名、使用する設定ファイルを個別に指定できます。
```bash
# 入力: sample.md
# 出力: sample_presentation.pptx
# 設定: default.yaml
python md2pptx.py sample.md -c default.yaml -o sample_presentation.pptx
```

| オプション | 引数 | 説明 | デフォルト値 |
| :--- | :--- | :--- | :--- |
| `-o`, `--output` | PATH | 出力するPowerPointファイルのパス（ファイル名）。 | `output.pptx` |
| `-c`, `--config` | PATH | 使用するYAML設定ファイルのパス。 | `config.yaml` |
| `-h`, `--help` | - | ヘルプメッセージを表示して終了。 | - |

## 設定ファイル (`config.yaml`) の設定値

YAMLファイル内でプレゼンテーションのデフォルトデザインを定義します。

### 重要：文字コードについて

YAMLファイル内に日本語（游ゴシックなど）を記載する場合、ファイルは必ず **UTF-8（BOMなし）** の文字コードで保存してください。Shift-JISなどで保存すると、スクリプト実行時に読み込みエラー（`'utf-8' codec can't decode byte...`）が発生します。

### YAML構造と項目一覧

デフォルトの `config.yaml` は以下の構造になっています。
```yaml
# --- スライド全体の基本設定 (slides) ---
slides:
  # 既存のPPTXテンプレートを使用する場合のパス（指定しない場合は null）
  # 社内ロゴや背景デザインが入ったマスターファイルがある場合、ここに指定します。
  # template_path: "company_template.pptx"
  template_path: null

  # スライドの画角（縦横比）。template_path が指定されている場合は無視されます。
  # 指定可能値: "16:9" (ワイド), "4:3" (標準), "16:10", "A4" (横向きA4)
  layout: "16:9"

# --- テキストのフォント設定 (fonts) ---
# OSにインストールされている正確なフォント名を指定してください。
fonts:
  # タイトル (h1, h2タグ)
  title:
    name: "Calibri"           # フォント名 (例: "Yu Gothic", "MS Gothic", "Arial")
    size_pt: 36               # サイズ (ポイント)
    bold: true                # 太字 (true / false)
    color_rgb: [0, 51, 102]   # 色 [R, G, B] (各0-255)

  # 本文 (pタグ)
  body:
    name: "Calibri"
    size_pt: 20
    bold: false
    color_rgb: [0, 0, 0]      # 黒

  # 箇条書きレベル1 (liタグ)
  bullet_level_1:
    name: "Calibri"
    size_pt: 18
    bold: false
    color_rgb: [30, 30, 30]   # 濃い灰色

# --- 画像の配置設定 (images) ---
images:
  default_height_inches: 4.0 # 画像のデフォルトの高さ（インチ）
  
  # 画像の初期配置位置（スライドの左上を (0,0) としたインチ単位）
  # [左からの距離(left), 上からの距離(top)]
  position_inches: [5.5, 2.0] # 右側に寄せる配置例
```

### OS別フォント名の例

| OS | 和文ゴシック体 (例) | 欧文サンセリフ体 (例) |
| :--- | :--- | :--- |
| **Windows** | `"Yu Gothic"`, `"MS Gothic"` | `"Calibri"`, `"Arial"` |
| **macOS** | `"Hiragino Sans"`, `"Hiragino Kaku Gothic ProN"` | `"Helvetica"`, `"Arial"` |

*※ 和文フォントを指定した場合、欧文も自動的にそのフォントの欧文リソースが使われます。*

## トラブルシューティング

### 設定ファイルの読み込みエラー
`Error: 設定ファイルの読み込みに失敗しました: 'utf-8' codec can't decode byte 0x83...`
*   **原因**: YAML設定ファイルがUTF-8ではない文字コード（Shift-JISなど）で保存されています。
*   **対策**: VS Codeやメモ帳などのエディタで、**文字コードを「UTF-8」に指定して上書き保存**してください（「UTF-8 BOM付き」は避けてください）。

### 画像の挿入に失敗する
`Warning: 画像の挿入に失敗しました (http...): ...`
*   **原因**: Web上のURLを指定した場合のネットワークエラー、タイムアウト、または画像URLが不正です。
*   **対策**: Markdown内の `src` に記述したURLがブラウザで直接開けるか確認してください。ローカルファイルの場合はパスが正しいか確認してください。

### ライブラリの依存競合エラー
`ERROR: pip's dependency resolver... launch-ros 0.26.11 requires setuptools...`
*   **原因**: システム環境（OS）にインストールされている特定のパッケージ（ROS関連など）が、仮想環境内のライブラリ不足に文句を言っています。
*   **対策**: **無視して構いません。** 本スクリプトに必要なライブラリ自体は正常にインストールされており、ROS関連の機能は使用しないため問題なく動作します。

