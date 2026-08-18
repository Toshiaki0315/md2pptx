#!/bin/sh
#
# md2pptx の起動スクリプト
#
#   sh run.sh                        GUI（画面）を起動する
#   sh run.sh 資料.md -o 資料.pptx     コマンドラインで変換する
#
# 初回は仮想環境（.venv）を作り、必要なライブラリを導入する。
# 2回目以降はそのまま起動するだけなので、待ち時間は無い。
#
# 仮想環境を作るPythonは環境変数 PYTHON で指定できる:
#   PYTHON=/opt/homebrew/bin/python3.13 sh run.sh
# （.venv が既にある場合、この指定は使われない。作り直すには rm -rf .venv）

set -u

# 実行した場所（作業ディレクトリ）は変えない。
# Markdownや出力先を相対パスで指定できるようにするため。
APP_DIR="$(cd "$(dirname "$0")" && pwd)" || exit 1

VENV_PYTHON="$APP_DIR/.venv/bin/python"
[ -x "$VENV_PYTHON" ] || VENV_PYTHON="$APP_DIR/.venv/Scripts/python.exe"   # Windows

# GUIには Tk 8.6 以上が必要。macOS標準のPythonが持つ Tk 8.5 では
# ウィンドウが真っ白のまま表示されないことがある。
has_usable_tk() {
    "$1" -c "import tkinter; raise SystemExit(0 if tkinter.TkVersion >= 8.6 else 1)" 2>/dev/null
}

# 仮想環境を作るPythonを選ぶ（GUIが動くもの・新しいものを優先）
find_base_python() {
    fallback=''
    for name in python3.14 python3.13 python3.12 python3.11 python3 python; do
        command -v "$name" >/dev/null 2>&1 || continue
        if has_usable_tk "$name"; then
            echo "$name"
            return 0
        fi
        [ -n "$fallback" ] || fallback="$name"
    done
    [ -n "$fallback" ] || return 1
    echo "$fallback"
}

# --- 仮想環境を用意する ---
# ライブラリは必ず .venv に入れる（利用者のシステムのPythonは変更しない）
if [ ! -x "$VENV_PYTHON" ]; then
    if [ -n "${PYTHON:-}" ]; then
        base_python="$PYTHON"
    else
        base_python="$(find_base_python)" || {
            echo "Error: Python が見つかりません。python.org などから導入してください。" >&2
            exit 1
        }
    fi
    echo "==> 仮想環境（.venv）を作成します（$base_python）"
    "$base_python" -m venv "$APP_DIR/.venv" || {
        echo "Error: 仮想環境を作成できませんでした。" >&2
        exit 1
    }
    VENV_PYTHON="$APP_DIR/.venv/bin/python"
    [ -x "$VENV_PYTHON" ] || VENV_PYTHON="$APP_DIR/.venv/Scripts/python.exe"
fi
PYTHON_BIN="$VENV_PYTHON"

# --- 依存ライブラリを確認し、足りなければ導入する ---
if ! "$PYTHON_BIN" -c "import yaml, pptx, markdown, bs4, requests, PIL, pygments" >/dev/null 2>&1; then
    echo "==> 必要なライブラリを導入します（初回のみ・数分かかることがあります）"
    "$PYTHON_BIN" -m pip install --quiet --upgrade pip
    "$PYTHON_BIN" -m pip install --quiet -r "$APP_DIR/requirements.txt" || {
        echo "Error: ライブラリを導入できませんでした。ネットワーク接続を確認してください。" >&2
        exit 1
    }
fi

# --- GUIを起動する（引数なしの場合）---
if [ "$#" -eq 0 ]; then
    if ! has_usable_tk "$PYTHON_BIN"; then
        echo "Warning: このPythonのTkが古いため、ウィンドウが表示されないことがあります。" >&2
        echo "         macOSでは 'brew install python-tk@3.13' などでTk 8.6を導入し、" >&2
        echo "         仮想環境を作り直してください（rm -rf .venv && sh run.sh）。" >&2
    fi
    exec "$PYTHON_BIN" "$APP_DIR/gui.py"
fi

# --- コマンドラインで変換する ---
# 設定ファイルの指定が無く、手元にも config.yaml が無い場合は、
# md2pptx に同梱のものを使う（どこから実行しても動くように）
case " $* " in
    *" -c "*|*" --config "*) ;;
    *) [ -f config.yaml ] || set -- "$@" -c "$APP_DIR/config.yaml" ;;
esac

exec "$PYTHON_BIN" "$APP_DIR/md2pptx.py" "$@"
