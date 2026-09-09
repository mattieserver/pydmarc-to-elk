#!/bin/bash
# Read the DMARC reports from the mailbox and ship them to OpenSearch.
# Run ./install.sh first. Pass "reload" to re-ingest data/processed/ instead.

set -euo pipefail
cd "$(dirname "$0")"
VIRTUALENV="$(pwd -P)/venv"

if [ ! -d "${VIRTUALENV}" ]; then
  echo "ERROR: No virtual environment at ${VIRTUALENV}. Run ./install.sh first."
  exit 1
fi

source "${VIRTUALENV}/bin/activate"

if [ "${1:-}" = "reload" ]; then
  exec python3 pyStartReloadDataV2.py
fi
exec python3 pyStartv2.py
