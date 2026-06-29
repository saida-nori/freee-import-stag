#!/bin/bash
# Claude Code (web) 用 SessionStart フック
# Python 依存関係をインストールし、楽楽精算→freee 仕訳CSV 生成ツールを
# Web セッションで実行・検証できる状態にする。
set -euo pipefail

# リモート（Claude Code on the web）環境でのみ実行する
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "$CLAUDE_PROJECT_DIR"

# 依存関係をインストール（再実行しても安全 / 非対話）
python3 -m pip install -r requirements.txt

echo "session-start hook: dependencies installed."
