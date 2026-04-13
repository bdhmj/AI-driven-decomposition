#!/bin/bash
# Archive current project input/output into projects/<name>/
# Usage: ./scripts/archive.sh "Project Name"

set -e

PROJECT_NAME="${1:?Usage: ./scripts/archive.sh \"Project Name\"}"
SAFE_NAME=$(echo "$PROJECT_NAME" | tr ' ' '_' | tr -cd '[:alnum:]_-')
DATE=$(date +%Y-%m-%d)
DIR="projects/${DATE}_${SAFE_NAME}"

mkdir -p "$DIR"

# Move input and output
# Move files (excluding .gitkeep)
find input -maxdepth 1 -not -name '.gitkeep' -not -path input -exec mv {} "$DIR/" \; 2>/dev/null || true
find output -maxdepth 1 -not -name '.gitkeep' -not -path output -exec mv {} "$DIR/" \; 2>/dev/null || true

# Ensure .gitkeep remains
touch input/.gitkeep output/.gitkeep

echo "Archived to $DIR"
