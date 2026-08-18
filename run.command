#!/bin/sh
#
# Finderからダブルクリックで起動するためのファイル（macOS）。
# 中身は run.sh と同じで、GUIを起動する。
#
exec "$(dirname "$0")/run.sh"
