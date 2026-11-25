
#!/usr/bin/env bash

set -euo pipefail

# Move to the project directory
cd "$(dirname "$0")"

# Validate Docker Compose configuration
docker compose config --quiet

# Start persistent services
docker compose up -d

# Build the Python image
docker compose build ia

# Run tests
docker compose run --rm ia python -m pytest -v \
    -o cache_dir=/tmp/.pytest_cache

# Run the processing pipeline
docker compose run --rm ia python main.py "$@"

echo "Pipeline completed successfully."

