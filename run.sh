#!/usr/bin/env bash

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate

if ! python3 -c "import customtkinter, PIL" >/dev/null 2>&1; then
  pip install -r requirements.txt
fi

python3 teaching_eval_app.py
