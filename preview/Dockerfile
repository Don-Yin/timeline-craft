# Preview service with persistent LibreOffice
# Using Python 3.12 + LibreOffice installed via package manager
FROM python:3.12-slim-bookworm

# Install LibreOffice and fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-impress \
    fonts-liberation \
    fonts-dejavu \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install timeline package first
COPY package/ /app/package/
RUN pip install --no-cache-dir /app/package/

# Install Python dependencies (without unoserver since we can't use it with pip python)
COPY preview/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY preview/ .

# Expose port 8004 for FastAPI
EXPOSE 8004

# Make start script executable
RUN chmod +x /app/start.sh

CMD ["/app/start.sh"]

