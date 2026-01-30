FROM python:3.12-slim

WORKDIR /app

RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice fonts-dejavu-core && \
    rm -rf /var/lib/apt/lists/*

COPY package /app/package
RUN pip install --no-cache-dir /app/package

COPY worker/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY worker/src ./src
COPY worker/server.py .

EXPOSE 8002
CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8002"]
