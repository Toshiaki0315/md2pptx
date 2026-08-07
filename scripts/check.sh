#!/bin/sh
#
# ローカルでの品質チェック（テスト・型チェック）
#
#   sh scripts/check.sh
#
# CIを使わない代わりに、コミット前やpush前に実行する想定。
# .githooks/pre-push から呼ばれ、push時に自動実行される。
#
# 使用するPythonは環境変数 PYTHON で上書きできる:
#   PYTHON=.venv/bin/python sh scripts/check.sh

set -u

cd "$(dirname "$0")/.." || exit 1

# python3 を優先し、無ければ python を使う
if [ -z "${PYTHON:-}" ]; then
    if command -v python3 >/dev/null 2>&1; then
        PYTHON=python3
    else
        PYTHON=python
    fi
fi

if ! command -v "$PYTHON" >/dev/null 2>&1; then
    echo "Error: Python が見つかりません（PYTHON=$PYTHON）。" >&2
    exit 1
fi

for module in pytest mypy; do
    if ! "$PYTHON" -c "import $module" >/dev/null 2>&1; then
        echo "Error: $module が見つかりません。" >&2
        echo "       pip install -r requirements-dev.txt を実行してください。" >&2
        exit 1
    fi
done

status=0

# どちらか一方で失敗しても両方の結果を見たいので、途中で止めない
echo "==> テスト (pytest)"
if ! "$PYTHON" -m pytest -q; then
    status=1
fi

echo ""
echo "==> 型チェック (mypy)"
if ! "$PYTHON" -m mypy .; then
    status=1
fi

echo ""
if [ "$status" -eq 0 ]; then
    echo "すべてのチェックを通過しました。"
else
    echo "チェックに失敗しました。上記の出力を確認してください。" >&2
fi

exit "$status"
