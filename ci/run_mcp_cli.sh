#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

# Load .env from project root
ENV_FILE="$PROJECT_DIR/.env"
if [ -f "$ENV_FILE" ]; then
  set -a
  # shellcheck disable=SC1090
  source "$ENV_FILE"
  set +a
fi

: "${ORACLE_USER:?ORACLE_USER not set in .env}"
: "${ORACLE_PASSWORD:?ORACLE_PASSWORD not set in .env}"
: "${ORACLE_CONNECTION_STRING:?ORACLE_CONNECTION_STRING not set in .env}"

cd "$PROJECT_DIR"

echo "Compiling doc-parser packages via sqlplus..."
sqlplus -s "${ORACLE_USER}/${ORACLE_PASSWORD}@${ORACLE_CONNECTION_STRING}" @compile_docs_parser.sql
