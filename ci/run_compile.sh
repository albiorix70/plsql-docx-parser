#!/usr/bin/env bash
set -euo pipefail

# CI helper: run run_compile.sql using sqlcl
# Requires these secrets/env vars in CI:
#  - SQLCL_DIR (set by workflow after download)
#  - DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE

LOGFILE=ci/compile_output.log
mkdir -p ci
: > ${LOGFILE}

if [ -z "${DB_USER:-}" ] || [ -z "${DB_PASSWORD:-}" ] || [ -z "${DB_HOST:-}" ] || [ -z "${DB_PORT:-}" ] || [ -z "${DB_SERVICE:-}" ]; then
  echo "Missing DB connection environment variables. Set DB_USER, DB_PASSWORD, DB_HOST, DB_PORT, DB_SERVICE." | tee -a ${LOGFILE}
  exit 1
fi

if [ -z "${SQLCL_DIR:-}" ]; then
  # try common install location
  if [ -d "/opt/sqlcl/sqlcl" ]; then
    SQLCL_DIR=/opt/sqlcl/sqlcl
  else
    echo "SQLCL_DIR not set and /opt/sqlcl/sqlcl not found." | tee -a ${LOGFILE}
    exit 1
  fi
fi

SQL_BIN="${SQLCL_DIR}/bin/sql"
if [ ! -x "${SQL_BIN}" ]; then
  echo "sql executable not found at ${SQL_BIN}" | tee -a ${LOGFILE}
  ls -la "${SQLCL_DIR}" >> ${LOGFILE} 2>&1 || true
  exit 1
fi

CONNECT_STRING="${DB_USER}/${DB_PASSWORD}@//${DB_HOST}:${DB_PORT}/${DB_SERVICE}"

echo "Running compile script with: ${CONNECT_STRING}" | tee -a ${LOGFILE}
# Run the run_compile.sql script; capture output
"${SQL_BIN}" -L "${CONNECT_STRING}" @run_compile.sql | tee -a ${LOGFILE}

EXIT_CODE=${PIPESTATUS[0]:-0}
if [ ${EXIT_CODE} -ne 0 ]; then
  echo "sqlcl exited with ${EXIT_CODE}; see ${LOGFILE}" | tee -a ${LOGFILE}
  exit ${EXIT_CODE}
fi

echo "Compile finished; see ${LOGFILE}" | tee -a ${LOGFILE}
