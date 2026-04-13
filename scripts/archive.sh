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
[ "$(ls -A input/ 2>/dev/null)" ] && mv input/* "$DIR/" 2>/dev/null || true
[ "$(ls -A output/ 2>/dev/null)" ] && mv output/* "$DIR/" 2>/dev/null || true

echo "Archived to $DIR"
