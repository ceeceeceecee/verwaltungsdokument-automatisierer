#!/bin/bash
set -e
echo "=== Installing verwaltungsdokument-automatisierer ==="

if ! command -v docker &> /dev/null; then
    echo "Docker is required."
    exit 1
fi

echo "Starting Ollama service..."
docker compose up -d ollama
sleep 5

echo "Pulling llama3.1 model..."
docker compose exec ollama ollama pull llama3.1 2>/dev/null || echo "Model pull initiated"

echo "Starting all services..."
docker compose up -d

echo "=== verwaltungsdokument-automatisierer is running ==="
echo "Open http://localhost:8503"
