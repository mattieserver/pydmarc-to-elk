#!/bin/bash
# One-time setup: create the virtual environment and install the dependencies.
# Re-running is safe; it reuses an existing venv and upgrades the packages in place.
#
# Invokes Python from the PYTHON environment variable, or falls back to "python3".

set -euo pipefail
cd "$(dirname "$0")"
VIRTUALENV="$(pwd -P)/venv"
PYTHON="${PYTHON:-python3}"

# Validate the minimum required Python version
"${PYTHON}" -c 'import sys; exit(1 if sys.version_info < (3, 10) else 0)' || {
  echo "--------------------------------------------------------------------"
  echo "ERROR: Unsupported Python version: $(${PYTHON} -V). This tool requires"
  echo "Python 3.10 or later (opensearch-py 3.x). To specify an alternate Python"
  echo "executable, set"
  echo "the PYTHON environment variable. For example:"
  echo ""
  echo "  PYTHON=/usr/bin/python3.11 ./install.sh"
  echo "--------------------------------------------------------------------"
  exit 1
}
echo "Using $(${PYTHON} -V)"

if [ ! -d "${VIRTUALENV}" ]; then
  echo "Creating a new virtual environment at ${VIRTUALENV}..."
  "${PYTHON}" -m venv "${VIRTUALENV}"
fi

source "${VIRTUALENV}/bin/activate"

echo "Updating pip..."
pip install --upgrade pip wheel

echo "Installing dependencies..."
pip install -r requirements.txt

if [ ! -f "Settings/config.ini" ]; then
  echo "Creating a default Settings/config.ini..."
  python3 writedefaultconf.py
  echo "--------------------------------------------------------------------"
  echo "Edit Settings/config.ini with your credentials before running ./run.sh"
  echo "--------------------------------------------------------------------"
fi
