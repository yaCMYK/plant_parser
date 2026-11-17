# syntax=docker/dockerfile:1
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DATA_DIR=/data \
    PARSER_BROWSER=requests
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
# Optional: docker build --build-arg INSTALL_PLAYWRIGHT=true ...
ARG INSTALL_PLAYWRIGHT=false
RUN if [ "$INSTALL_PLAYWRIGHT" = "true" ]; then \
      pip install --no-cache-dir 'playwright>=1.49,<2' && \
      python -m playwright install --with-deps chromium; \
    fi
COPY plant_parser_production.py web.py ./
COPY templates ./templates
RUN mkdir -p /data
CMD ["sh", "-c", "exec gunicorn --bind 0.0.0.0:${PORT:-8080} --workers 1 --threads 4 --timeout 60 --access-logfile - --error-logfile - 'web:create_app()'"]
