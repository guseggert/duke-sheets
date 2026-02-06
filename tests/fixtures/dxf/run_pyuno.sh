#!/bin/bash
# Script to run PyUNO fixture framework.
#
# Usage:
#   ./run_pyuno.sh                    # Generate all fixtures
#   ./run_pyuno.sh pyuno_dxf_font.xlsx  # Generate specific fixture
#   ./run_pyuno.sh --list             # List available fixtures

set -e

OUTPUT_DIR="${OUTPUT_DIR:-/output}"
SCRIPT_DIR="$(dirname "$0")"
FRAMEWORK="$SCRIPT_DIR/pyuno_framework.py"

# Check for --list flag
if [[ "$1" == "--list" || "$1" == "-l" ]]; then
    echo "Starting LibreOffice for fixture listing..."
    soffice --headless --invisible --nologo --nofirststartwizard \
        --accept="socket,host=localhost,port=2002;urp;" &
    LO_PID=$!
    sleep 3
    python3 "$FRAMEWORK" --list
    kill $LO_PID 2>/dev/null || true
    exit 0
fi

echo "============================================================"
echo "PyUNO Fixture Generator"
echo "============================================================"
echo ""

echo "Starting LibreOffice in listening mode..."
soffice \
    --headless \
    --invisible \
    --nologo \
    --nofirststartwizard \
    --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" \
    &

LO_PID=$!

# Wait for LibreOffice to start
echo "Waiting for LibreOffice to start..."
sleep 5

# Check if it's running
if ! kill -0 $LO_PID 2>/dev/null; then
    echo "ERROR: LibreOffice failed to start"
    exit 1
fi

echo "LibreOffice started (PID: $LO_PID)"
echo ""

# Run the Python framework with any passed arguments
echo "Running fixture generator..."
python3 "$FRAMEWORK" --output "$OUTPUT_DIR" "$@" || true

# Cleanup
echo ""
echo "Stopping LibreOffice..."
kill $LO_PID 2>/dev/null || true
wait $LO_PID 2>/dev/null || true

echo "Done!"
