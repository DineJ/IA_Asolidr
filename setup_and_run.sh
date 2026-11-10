#!/usr/bin/env bash
# Set up IA_Asolidr, run the test suite, then execute the data pipeline.
set -Eeuo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")"

OVERWRITE=0
SKIP_TESTS=0
SKIP_RUN=0

usage() {
  cat <<'HELP'
Usage: bash setup_and_run.sh [--ecraser] [--skip-tests] [--tests-only]

  --ecraser       Allow main.py to overwrite generated CSV files
  --skip-tests    Skip pytest and run only the pipeline after installation
  --tests-only    Install dependencies and run pytest, without running main.py
  -h, --help      Show this help

Default: create/reuse .env, update pip, install requirement.txt and pytest,
         run all tests, then run main.py if all tests pass.
HELP
}

for arg in "$@"; do
  case "$arg" in
    --ecraser) OVERWRITE=1 ;;
    --skip-tests) SKIP_TESTS=1 ;;
    --tests-only) SKIP_RUN=1 ;;
    -h|--help) usage; exit 0 ;;
    *) printf 'Unknown option: %s\n' "$arg" >&2; usage >&2; exit 2 ;;
  esac
done

if (( SKIP_TESTS && SKIP_RUN )); then
  echo 'Cannot combine --skip-tests and --tests-only.' >&2
  exit 2
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo 'Python 3 is required.' >&2
  exit 1
fi
if [[ ! -f requirement.txt ]]; then
  echo 'Missing requirement.txt in the project root.' >&2
  exit 1
fi
if [[ ! -f main.py && $SKIP_RUN -eq 0 ]]; then
  echo 'Missing main.py in the project root.' >&2
  exit 1
fi

if [[ ! -x .env/bin/python ]]; then
  echo '>>> Creating virtual environment .env'
  python3 -m venv .env
else
  echo '>>> Reusing existing .env'
fi

PYTHON="$PWD/.env/bin/python"
echo '>>> Updating pip'
"$PYTHON" -m pip install --upgrade pip

echo '>>> Installing project dependencies'
"$PYTHON" -m pip install -r requirement.txt

if (( ! SKIP_TESTS )); then
  echo '>>> Installing pytest'
  "$PYTHON" -m pip install pytest
  if [[ -f requirements-dev.txt ]]; then
    echo '>>> Installing optional development dependencies'
    "$PYTHON" -m pip install -r requirements-dev.txt
  fi
  echo '>>> Running all tests'
  "$PYTHON" -m pytest -v
fi

if (( ! SKIP_RUN )); then
  if [[ ! -d data/raw ]]; then
    echo 'Missing data/raw directory: add your Excel files before running the pipeline.' >&2
    exit 1
  fi
  if [[ ! -f corrections.csv ]]; then
    echo 'Missing corrections.csv in the project root.' >&2
    exit 1
  fi
  ARGS=()
  if (( OVERWRITE )); then
    ARGS+=(--ecraser)
  fi
  echo '>>> Running the full pipeline'
  "$PYTHON" main.py "${ARGS[@]}"
fi

echo '>>> Completed successfully'
