FROM python:3.12-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /opt/cabinet/Serveur
COPY Serveur/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY Serveur/cabinet ./cabinet
COPY Build/schemas.json ../Build/schemas.json
RUN useradd --uid 1026 --create-home cabinet
USER 1026:100
CMD ["python", "-m", "uvicorn", "cabinet.api:app", "--host", "0.0.0.0", "--port", "8765", "--no-access-log", "--limit-concurrency", "32", "--timeout-keep-alive", "5"]
