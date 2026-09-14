FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

COPY telegram_india_mvp/requirements.txt /app/telegram_india_mvp/requirements.txt
RUN python -m pip install --no-cache-dir -r /app/telegram_india_mvp/requirements.txt

COPY . /app

RUN mkdir -p /data

WORKDIR /app/telegram_india_mvp

CMD ["python", "bot.py"]
