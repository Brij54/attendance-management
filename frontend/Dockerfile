# =====================================================================
# STAGE 1: Builder (install dependencies and prune unnecessary files)
# =====================================================================
FROM python:3.12-slim AS builder

WORKDIR /app

# Upgrade pip
RUN pip install --no-cache-dir --upgrade pip

# Install dependencies in two layers to avoid oversized image layers
# 1. Core Streamlit framework and HTTP client
RUN pip install --no-cache-dir \
    streamlit==1.48.0 \
    requests==2.32.4 \
    python-dotenv==1.1.1

# 2. Data processing and visualization packages
RUN pip install --no-cache-dir \
    openpyxl \
    msoffcrypto-tool \
    plotly

# 3. Clean up tests, pycache, and bytecode to minimize size
RUN find /usr/local/lib/python3.12 -type d -name "tests" -prune -exec rm -rf {} + \
    && find /usr/local/lib/python3.12 -type d -name "__pycache__" -prune -exec rm -rf {} + \
    && find /usr/local/lib/python3.12 -name "*.pyc" -delete

# =====================================================================
# STAGE 2: Minimal Runtime Image
# =====================================================================
FROM python:3.12-slim

WORKDIR /app

# Install minimal runtime tools: curl for container health checks, libcap2-bin for port 80 binding
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    libcap2-bin \
    && rm -rf /var/lib/apt/lists/*

# Copy pre-installed python packages and binaries from builder stage
COPY --from=builder /usr/local/lib/python3.12/site-packages /usr/local/lib/python3.12/site-packages
COPY --from=builder /usr/local/bin /usr/local/bin

# Allow python3 binary to bind to port 80 as a non-privileged user
RUN setcap 'cap_net_bind_service=+ep' $(readlink -f $(which python3))

# Create non-root user for security
RUN useradd -u 1000 -m appuser && chown -R appuser:appuser /app

# Copy application assets and code
COPY .streamlit ./.streamlit
COPY assets ./assets
COPY app.py .

USER appuser

EXPOSE 80

# Configure Streamlit production settings via environment variables
ENV STREAMLIT_SERVER_PORT=80 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_SERVER_ENABLE_CORS=false \
    STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION=false \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

CMD ["streamlit", "run", "app.py"]
