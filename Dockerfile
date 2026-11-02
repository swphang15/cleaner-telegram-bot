FROM python:3.13-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    BOT_DATA_DIR=/data \
    BOT_TOKEN_FILE=/run/secrets/telegram_token

WORKDIR /app

COPY requirements.txt ./
RUN apt-get update \
    && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends tzdata \
    && rm -rf /var/lib/apt/lists/* \
    && pip install --no-cache-dir -r requirements.txt \
    && groupadd --gid 10001 bot \
    && useradd --uid 10001 --gid bot --no-create-home bot \
    && mkdir /data \
    && chown bot:bot /data

COPY --chown=bot:bot bot.py ./

USER 10001:10001
CMD ["python", "bot.py"]
