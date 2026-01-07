#!/usr/bin/env bash
set -euo pipefail

# Wrapper to invoke an MCP/CLI (SQLcl-like) binary to run run_compile.sql
# Environment variables:
#   MCP_CLI   - path to the mcp/sqlcl binary (optional)
#   DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE

LOGFILE=ci/mcp_compile_output.log
mkdir -p ci
: > ${LOGFILE}

if [ -z "${DB_USER:-}" ] || [ -z "${DB_PASSWORD:-}" ] || [ -z "${DB_HOST:-}" ] || [ -z "${DB_PORT:-}" ] || [ -z "${DB_SERVICE:-}" ]; then
  echo "Missing DB connection environment variables. Set DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE." | tee -a ${LOGFILE}
  exit 1
fi

MCP_BIN=${MCP_CLI:-}
if [ -z "${MCP_BIN}" ]; then
  # try common binary names
  for cmd in mcp sdcli sql; do
    if command -v "${cmd}" >/dev/null 2>&1; then
      MCP_BIN=$(command -v "${cmd}")
      break
    fi
  done
fi

if [ -z "${MCP_BIN}" ]; then
  echo "MCP/CLI binary not found. Set MCP_CLI env var or install 'mcp'/'sdcli'/'sql' on PATH." | tee -a ${LOGFILE}
  exit 1
fi

CONNECT_STRING="${DB_USER}/${DB_PASSWORD}@//${DB_HOST}:${DB_PORT}/${DB_SERVICE}"

echo "Using MCP/CLI: ${MCP_BIN}" | tee -a ${LOGFILE}

echo "Running: ${MCP_BIN} -L \"${CONNECT_STRING}\" @run_compile.sql" | tee -a ${LOGFILE}

"${MCP_BIN}" -L "${CONNECT_STRING}" @run_compile.sql 2>&1 | tee -a ${LOGFILE}

EXIT_CODE=${PIPESTATUS[0]:-0}
if [ ${EXIT_CODE} -ne 0 ]; then
  echo "MCP/CLI exited with ${EXIT_CODE}; see ${LOGFILE}" | tee -a ${LOGFILE}
  exit ${EXIT_CODE}
fi

echo "MCP/CLI compile finished; see ${LOGFILE}" | tee -a ${LOGFILE}
