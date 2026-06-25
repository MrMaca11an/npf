#!/usr/bin/env bash
# Удобный запуск системы сверки БУ-ПУ (Linux/macOS).
# Открывает интерактивное меню. Передайте любые флаги, чтобы пробросить их в CLI.
set -e
cd "$(dirname "$0")"

PY=python3
command -v "$PY" >/dev/null 2>&1 || PY=python

if [ "$#" -eq 0 ]; then
  "$PY" -m npf_recon --menu
else
  "$PY" -m npf_recon "$@"
fi
