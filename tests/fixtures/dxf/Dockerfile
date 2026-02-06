# Dockerfile for generating DXF test fixtures using LibreOffice PyUNO
#
# Build:
#   docker build -t dxf-fixture-gen tests/fixtures/dxf/
#
# Run:
#   docker run --rm -v $(pwd)/tests/fixtures/dxf:/output dxf-fixture-gen
#   docker run --rm -v $(pwd)/tests/fixtures/dxf:/output dxf-fixture-gen --list
#   docker run --rm -v $(pwd)/tests/fixtures/dxf:/output dxf-fixture-gen pyuno_dxf_font.xlsx
#
FROM debian:bookworm-slim

# Install LibreOffice with Python UNO support
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-script-provider-python \
    python3-uno \
    python3 \
    procps \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy framework and scripts
COPY pyuno_framework.py /app/
COPY run_pyuno.sh /app/
RUN chmod +x /app/run_pyuno.sh

# Output directory
VOLUME /output
ENV OUTPUT_DIR=/output

# Run the generator (pass any args to the script)
ENTRYPOINT ["/app/run_pyuno.sh"]
