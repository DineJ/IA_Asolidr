
# Use the official Python image
FROM python:3.12-slim

# Configure Python
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set the working directory
WORKDIR /app

# Copy Python dependencies
COPY requirement.txt .

# Install Python dependencies
RUN python -m pip install --no-cache-dir --upgrade pip \
    && python -m pip install --no-cache-dir -r requirement.txt

# Copy application source code
COPY . .

# Default command
CMD ["python", "main.py"]
