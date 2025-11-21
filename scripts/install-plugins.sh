#!/usr/bin/env bash
set -euo pipefail

SITE_DIR="${1:-/var/www/scottnandrew-dev}"
MANIFEST="${2:-.grav-plugins.txt}"

cd "$SITE_DIR"

if [[ ! -f "$MANIFEST" ]]; then
  echo "Manifest $MANIFEST not found in $SITE_DIR"; exit 1
fi

# Ensure GPM works; install Problems first (handy for diagnostics)
php bin/gpm install problems -y || true

# Install each plugin if missing (idempotent)
while IFS= read -r name; do
  [[ -z "$name" || "$name" =~ ^# ]] && continue
  if [[ -d "user/plugins/$name" ]]; then
    echo "✓ $name already present"
  else
    echo "→ Installing $name"
    php bin/gpm install "$name" -y || true
  fi
done < "$MANIFEST"

# Clear cache as web user (adjust if needed)
sudo -u www-data php bin/grav cache --all -n || true
