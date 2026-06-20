#!/usr/bin/env bash
set -euo pipefail

# If SERVICE_ACCOUNT_JSON env var is set with the JSON content, write it to a file
if [ -n "${SERVICE_ACCOUNT_JSON-}" ]; then
  echo "$SERVICE_ACCOUNT_JSON" > /tmp/service_account.json
  export GOOGLE_SERVICE_ACCOUNT_JSON=/tmp/service_account.json
fi

# Run the app with gunicorn
exec gunicorn web:app --bind 0.0.0.0:${PORT:-5000}
