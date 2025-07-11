# Use the official Python image.
FROM python:3.11-slim

# Install system dependencies for lightgbm and others
RUN apt-get update && apt-get install -y libgomp1

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set work directory
WORKDIR /app

# Install dependencies
COPY requirements.txt /app/
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# Copy project
COPY . /app/

# Expose the port
EXPOSE 8080

# Run the app with Gunicorn
CMD ["gunicorn", "App:app", "--bind", "0.0.0.0:8080"]