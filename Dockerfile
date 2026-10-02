# Use official Microsoft Playwright Python Jammy image with Chromium pre-installed
FROM mcr.microsoft.com/playwright/python:v1.48.0-jammy

# Set working directory
WORKDIR /app

# Set environment variables
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=10000

# Install python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Ensure playwright chromium browser is ready
RUN playwright install chromium

# Copy application files
COPY . .

# Expose port (Render sets $PORT dynamically)
EXPOSE 10000

# Run with Gunicorn production server (1 worker to stay within 512MB RAM on free tier)
CMD ["sh", "-c", "gunicorn app:app --bind 0.0.0.0:${PORT:-10000} --workers 1 --threads 4 --timeout 180"]
