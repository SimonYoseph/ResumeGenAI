#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_DIR"

if [[ ! -d ".venv" ]]; then
  echo "[setup] Creating virtual environment..."
  python3 -m venv .venv
fi

source .venv/bin/activate

if [[ ! -f ".venv/.deps_installed" ]] || [[ "requirements.txt" -nt ".venv/.deps_installed" ]]; then
  echo "[setup] Installing dependencies..."
  python3 -m pip install --upgrade pip
  pip install -r requirements.txt
  date > .venv/.deps_installed
fi

echo "[run] Starting Streamlit app..."
exec streamlit run app.py
