# LibreOffice Docker image for duke-sheets E2E tests
#
# Provides LibreOffice with URP (UNO Remote Protocol) socket access.
# The E2E tests connect to LibreOffice via URP to create XLSX fixtures
# on-demand, then read them back with the XLSX reader.
#
# Usage:
#   docker build -t duke-sheets-pyuno tests/fixtures/pyuno
#   docker run --rm -d -p 2002:2002 \
#     -v /tmp/duke-sheets-urp:/tmp/duke-sheets-urp \
#     duke-sheets-pyuno \
#     bash -c 'soffice --headless --accept="socket,host=0.0.0.0,port=2002;urp;StarOffice.ComponentContext" & sleep infinity'
#
# Or use the mise tasks:
#   mise run docker:build
#   mise run urp:start

FROM ubuntu:22.04

# Avoid interactive prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# Install LibreOffice
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

# Create shared temp directory for fixture files
RUN mkdir -p /tmp/duke-sheets-urp
