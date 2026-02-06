#!/bin/bash
# Run script for PyUNO fixture generation
#
# This script:
# 1. Starts LibreOffice in headless mode with socket listener
# 2. Waits for it to be ready
# 3. Runs the fixture generator
# 4. Cleans up
#
# Usage (inside container):
#   ./run.sh                    # Generate all fixtures
#   ./run.sh data_types.xlsx    # Generate specific fixture
#   ./run.sh --list             # List available fixtures
#   ./run.sh --verify file.xlsx # Verify a file (for write tests)

set -e

# Configuration
SOFFICE_PORT=2002
SOFFICE_TIMEOUT=30
OUTPUT_DIR="${OUTPUT_DIR:-/output}"

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

log_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

log_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Start LibreOffice in headless mode with socket listener
start_libreoffice() {
    log_info "Starting LibreOffice in headless mode..."
    
    # Kill any existing soffice processes
    pkill -9 soffice || true
    
    # Start LibreOffice with socket listener
    soffice --headless --accept="socket,host=localhost,port=${SOFFICE_PORT};urp;StarOffice.ServiceManager" &
    SOFFICE_PID=$!
    
    # Wait for LibreOffice to be ready
    log_info "Waiting for LibreOffice to be ready (timeout: ${SOFFICE_TIMEOUT}s)..."
    
    for i in $(seq 1 $SOFFICE_TIMEOUT); do
        if nc -z localhost $SOFFICE_PORT 2>/dev/null; then
            log_info "LibreOffice is ready (took ${i}s)"
            return 0
        fi
        sleep 1
    done
    
    log_error "Timeout waiting for LibreOffice to start"
    return 1
}

# Stop LibreOffice
stop_libreoffice() {
    log_info "Stopping LibreOffice..."
    pkill -9 soffice || true
}

# Cleanup on exit
cleanup() {
    stop_libreoffice
}
trap cleanup EXIT

# Main entry point
main() {
    # Parse arguments
    if [ "$1" == "--list" ]; then
        # List mode: just show fixtures without starting LibreOffice
        python3 -c "
from fixtures import *
from framework import get_fixtures

print('Available fixtures:')
for name in sorted(get_fixtures().keys()):
    print(f'  - {name}')
"
        exit 0
    fi
    
    if [ "$1" == "--verify" ]; then
        # Verification mode: verify a file written by Rust
        shift
        if [ -z "$1" ]; then
            log_error "Usage: $0 --verify <file.xlsx> [verification_spec.json]"
            exit 1
        fi
        
        start_libreoffice
        
        # Run verifier
        python3 /app/verifier.py "$@"
        exit $?
    fi
    
    # Generation mode: generate fixtures
    start_libreoffice
    
    log_info "Output directory: ${OUTPUT_DIR}"
    mkdir -p "${OUTPUT_DIR}"
    
    # Import all fixtures to register them
    log_info "Loading fixture definitions..."
    
    # Run the framework with all fixtures
    if [ $# -eq 0 ]; then
        log_info "Generating all fixtures..."
        python3 -c "
import sys
sys.path.insert(0, '/app')
from fixtures import *
from framework import run_fixtures

run_fixtures(output_dir='${OUTPUT_DIR}')
"
    else
        log_info "Generating specific fixtures: $*"
        FIXTURES=$(printf "'%s'," "$@")
        FIXTURES="[${FIXTURES%,}]"
        python3 -c "
import sys
sys.path.insert(0, '/app')
from fixtures import *
from framework import run_fixtures

run_fixtures(fixtures=${FIXTURES}, output_dir='${OUTPUT_DIR}')
"
    fi
    
    log_info "Done! Fixtures saved to ${OUTPUT_DIR}"
    ls -la "${OUTPUT_DIR}"
}

main "$@"
