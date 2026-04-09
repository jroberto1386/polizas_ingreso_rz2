FROM python:3.11-slim

RUN apt-get update && apt-get install -y libreoffice --no-install-recommends && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads outputs

CMD gunicorn app:app --bind 0.0.0.0:$PORT --timeout 300 --workers 2
