FROM python:3.12-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /opt/cabinet/Serveur
COPY Serveur/requirements-runtime.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY Serveur/cabinet ./cabinet
COPY Build/schemas.json ../Build/schemas.json
# COPY conserve les droits des sources locales, parfois restrictifs sur le NAS.
# Seul le code de l'image est concerne, jamais les donnees ni les secrets montes.
RUN find /opt/cabinet -type d -exec chmod 755 {} + \
    && find /opt/cabinet -type f -exec chmod 644 {} +
RUN useradd --uid 1026 --create-home cabinet
USER 1026:100
CMD ["python", "-m", "uvicorn", "cabinet.api:app", "--host", "0.0.0.0", "--port", "8765", "--no-access-log", "--limit-concurrency", "32", "--timeout-keep-alive", "5"]
