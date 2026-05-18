#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 が見つかりません。Python 3 をインストールしてください。"
  exit 1
fi

python3 -m pip install --user openpyxl pypdf >/dev/null
python3 app.py
