FROM python:3.11-slim

WORKDIR /app
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1

COPY requirements.txt ./

# Install system deps needed by some Python packages
RUN apt-get update && \
    apt-get install -y --no-install-recommends build-essential libjpeg-dev zlib1g-dev && \
    pip install --no-cache-dir -r requirements.txt && \
    apt-get remove -y build-essential && apt-get autoremove -y && rm -rf /var/lib/apt/lists/*

COPY . ./
RUN chmod +x ./start.sh

EXPOSE 5000
CMD ["./start.sh"]
