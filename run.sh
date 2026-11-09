#!/usr/bin/env bash
# Ativa o venv, abre o frontend e sobe o servidor backend
set -e

source venv/bin/activate

# Abre o frontend no navegador apos 2 segundos (tempo pro servidor subir)
(sleep 2 && \
  if command -v xdg-open &>/dev/null; then
    xdg-open frontend/index.html        # Linux
  elif command -v open &>/dev/null; then
    open frontend/index.html            # Mac
  fi
) &

python backend/main.py
