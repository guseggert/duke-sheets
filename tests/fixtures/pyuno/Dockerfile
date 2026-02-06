# PyUNO Fixture Generator for duke-sheets E2E tests
#
# This Dockerfile creates an environment with LibreOffice and Python-UNO
# for generating authoritative XLSX test fixtures.
#
# Usage:
#   docker build -t duke-sheets-pyuno tests/fixtures/pyuno
#   docker run -v $(pwd)/tests/fixtures/pyuno/output:/output duke-sheets-pyuno
#
# Or use the mise tasks:
#   mise run fixtures:build
#   mise run fixtures:generate

FROM ubuntu:22.04

# Avoid interactive prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# Install LibreOffice and dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    python3 \
    python3-uno \
    && rm -rf /var/lib/apt/lists/*

# Create output directory
RUN mkdir -p /output

# Create working directory
WORKDIR /app

# Copy framework and fixtures
COPY framework.py /app/
COPY fixtures/ /app/fixtures/
COPY run.sh /app/
RUN chmod +x /app/run.sh

# Set environment
ENV OUTPUT_DIR=/output
ENV PYTHONPATH=/app

# Default command: run the fixture generator
CMD ["/app/run.sh"]
