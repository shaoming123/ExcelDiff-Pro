FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    HOME=/home/excelapp

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && groupadd --system excelapp \
    && useradd --system --create-home --gid excelapp --home /home/excelapp excelapp

COPY --chown=excelapp:excelapp app.py ./
COPY --chown=excelapp:excelapp exceldiff ./exceldiff
COPY --chown=excelapp:excelapp assets ./assets
COPY --chown=excelapp:excelapp .streamlit ./.streamlit
USER excelapp

EXPOSE 8501
HEALTHCHECK --interval=30s --timeout=5s --start-period=25s \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8501/_stcore/health', timeout=3)" || exit 1
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=8501"]
