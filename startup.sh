#!/bin/bash

# Ensure all dependencies are installed
echo "Installing dependencies..."
pip install --no-cache-dir -r requirements.txt

# Start the Flask app
echo "Starting Flask app..."
gunicorn --bind=0.0.0.0:8000 wsgi:app