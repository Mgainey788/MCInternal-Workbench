FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=8000

# Install required system packages
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        build-essential \
        curl \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies first for better Docker build caching
COPY requirements.txt ./

RUN python -m pip install --upgrade pip \
    && python -m pip install --no-cache-dir -r requirements.txt

# Copy the full application
COPY . .

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=10s --start-period=60s --retries=3 \
    CMD curl --fail http://localhost:${PORT}/_stcore/health || exit 1

CMD ["sh", "-c", "python -m streamlit run streamlit_app.py --server.port=${PORT} --server.address=0.0.0.0 --server.headless=true"]
