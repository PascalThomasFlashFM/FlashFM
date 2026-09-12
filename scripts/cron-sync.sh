#!/usr/bin/env bash
# Appelle la synchro BOCIR -> Google Sheet. Ne declenche jamais d'envoi de mail :
# la validation des mails de confirmation reste manuelle (CLI `send`/`send-all`
# ou interface web).
set -euo pipefail
cd "$(dirname "$0")/.."
node dist/cli.js sync >> "${LOG_PATH:-./data/cron.log}" 2>&1
