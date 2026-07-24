FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py generate_lighting_dashboard.py test.xlsx ./
COPY templates ./templates

RUN python generate_lighting_dashboard.py \
      --excel test.xlsx \
      --output lighting_dashboard.html \
    && useradd --create-home --uid 10001 lighting \
    && chown -R lighting:lighting /app

USER lighting

EXPOSE 8000

CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "2", "--threads", "4", "--timeout", "120", "--access-logfile", "-", "--error-logfile", "-", "app:app"]
