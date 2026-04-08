#!/bin/bash
set -euo pipefail

# Finder/Terminal launcher may close the parent shell with `; exit;`.
# Ignore SIGHUP so the launcher can continue reliably.
trap '' HUP

# Resolve script directory without hardcoded absolute paths.
# This keeps launcher behavior stable even when the repo is cloned elsewhere.
if [[ "$0" == */* ]]; then
  SCRIPT_PATH="$0"
else
  SCRIPT_PATH="$(command -v -- "$0" || true)"
fi

if [[ -z "${SCRIPT_PATH}" ]]; then
  echo "[ERROR] Could not resolve launcher path: $0"
  echo "[HINT] From repo root, run: ./run_gui.command"
  exit 1
fi

SCRIPT_DIR="$(cd -- "$(dirname -- "${SCRIPT_PATH}")" && pwd -P)"
cd "$SCRIPT_DIR"

LOG_DIR="$SCRIPT_DIR/.logs"
mkdir -p "$LOG_DIR"
LOG_FILE="$LOG_DIR/run_gui-$(date +%Y%m%d-%H%M%S).log"
PIP_CACHE_DIR="$SCRIPT_DIR/.pip-cache"
mkdir -p "$PIP_CACHE_DIR"
export PIP_CACHE_DIR
export PIP_DISABLE_PIP_VERSION_CHECK=1

log() {
  printf '[%s] %s\n' "$(date '+%H:%M:%S')" "$*" | tee -a "$LOG_FILE"
}

log "Launcher started"
log "Project directory: $SCRIPT_DIR"

if [ ! -d ".venv" ]; then
  log "Python virtualenv not found. Creating .venv..."
  if command -v python3 >/dev/null 2>&1; then
    python3 -m venv .venv
  elif command -v python >/dev/null 2>&1; then
    python -m venv .venv
  else
    log "[ERROR] Python 3 is required."
    exit 1
  fi
fi

if [ ! -x ".venv/bin/python" ]; then
  log "[ERROR] Missing executable: .venv/bin/python"
  exit 1
fi

source .venv/bin/activate

if ! .venv/bin/python - <<'PY' >/dev/null 2>&1
import openpyxl
import PySide6
print(openpyxl.__version__, PySide6.__version__)
PY
then
  log "Installing required packages from requirements.txt..."
  if ! .venv/bin/python -m pip install --upgrade pip >>"$LOG_FILE" 2>&1; then
    log "[WARN] pip upgrade failed. Continuing with existing pip."
  fi
  if ! .venv/bin/python -m pip install -r requirements.txt >>"$LOG_FILE" 2>&1; then
    log "[ERROR] Failed to install dependencies from requirements.txt"
    log "[HINT] Check network access and rerun from repo root: ./run_gui.command"
    log "[HINT] Detailed log: $LOG_FILE"
    tail -n 60 "$LOG_FILE" || true
    exit 1
  fi
fi

DEP_VERSIONS="$(
.venv/bin/python - <<'PY'
import openpyxl, PySide6
print(f"openpyxl={openpyxl.__version__}, PySide6={PySide6.__version__}")
PY
)"
log "Dependency check passed: ${DEP_VERSIONS}"

if [[ "${RUN_GUI_DETACH:-0}" == "1" ]]; then
  log "Launching GUI in detached mode..."
  nohup .venv/bin/python excel_diff_gui.py >>"$LOG_FILE" 2>&1 &
  GUI_PID=$!
  sleep 1

  if kill -0 "$GUI_PID" >/dev/null 2>&1; then
    log "GUI started successfully (PID: $GUI_PID)"
    log "Detailed log: $LOG_FILE"
    exit 0
  fi

  log "[ERROR] GUI process exited immediately."
  log "[HINT] Detailed log: $LOG_FILE"
  tail -n 80 "$LOG_FILE" || true
  exit 1
fi

log "Launching GUI in foreground mode..."
log "Detailed log: $LOG_FILE"
exec .venv/bin/python excel_diff_gui.py >>"$LOG_FILE" 2>&1
