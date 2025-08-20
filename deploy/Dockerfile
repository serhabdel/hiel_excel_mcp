# Multi-stage build for optimized production image
FROM python:3.11-slim as builder

# Install build dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    make \
    libffi-dev \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Create virtual environment
RUN python -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"

# Copy requirements
COPY pyproject.toml /app/
WORKDIR /app

# Install Python dependencies
RUN pip install --upgrade pip && \
    pip install build && \
    python -m build . && \
    pip install dist/*.whl

# Production stage
FROM python:3.11-slim

# Install runtime dependencies
RUN apt-get update && apt-get install -y \
    curl \
    tesseract-ocr \
    tesseract-ocr-eng \
    libgomp1 \
    && rm -rf /var/lib/apt/lists/*

# Copy virtual environment from builder
COPY --from=builder /opt/venv /opt/venv

# Set environment variables
ENV PATH="/opt/venv/bin:$PATH" \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    EXCEL_MCP_FILES_PATH=/data \
    EXCEL_MCP_CACHE_PATH=/cache

# Create non-root user
RUN useradd -m -u 1000 mcp && \
    mkdir -p /data /cache /logs && \
    chown -R mcp:mcp /data /cache /logs

# Copy application code
COPY --chown=mcp:mcp . /app
WORKDIR /app

# Switch to non-root user
USER mcp

# Health check endpoint
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

# Default command
CMD ["python", "-m", "hiel_excel_mcp", "serve", "--host", "0.0.0.0", "--port", "8000"]
