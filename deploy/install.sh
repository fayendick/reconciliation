#!/usr/bin/env bash
# Installation prod — Reconciliation + nginx + systemd
# Usage (en root) :
#   sudo bash deploy/install.sh [/chemin/app] [server_name]
#
# Exemple :
#   sudo bash deploy/install.sh /opt/reconciliation reconc.banque.local

set -euo pipefail

APP_DIR="${1:-/opt/reconciliation}"
SERVER_NAME="${2:-_}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
SERVICE_USER="reconc"

if [[ "$(id -u)" -ne 0 ]]; then
  echo "À lancer en root : sudo bash deploy/install.sh"
  exit 1
fi

echo "==> Paquets système"
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq python3 python3-venv python3-pip nginx curl

echo "==> Utilisateur système $SERVICE_USER"
if ! id "$SERVICE_USER" &>/dev/null; then
  useradd --system --home-dir "$APP_DIR" --shell /usr/sbin/nologin "$SERVICE_USER"
fi

echo "==> Copie de l'application vers $APP_DIR"
mkdir -p "$APP_DIR"
rsync -a --delete \
  --exclude '.git' \
  --exclude '.venv' \
  --exclude '.venv_windows' \
  --exclude '__pycache__' \
  --exclude '_archive' \
  --exclude '*.pyc' \
  "$REPO_DIR/" "$APP_DIR/"

if [[ ! -f "$APP_DIR/.env" ]]; then
  if [[ -f "$APP_DIR/.env.example" ]]; then
    cp "$APP_DIR/.env.example" "$APP_DIR/.env"
    echo "    .env créé depuis .env.example — À RENSEIGNER (Oracle, etc.)"
  fi
fi

echo "==> Environnement virtuel Python"
python3 -m venv "$APP_DIR/.venv"
"$APP_DIR/.venv/bin/pip" install --upgrade pip -q
"$APP_DIR/.venv/bin/pip" install -r "$APP_DIR/requirements.txt" -q

mkdir -p "$APP_DIR/data/mappings"
chown -R "$SERVICE_USER:$SERVICE_USER" "$APP_DIR"

echo "==> systemd"
sed "s|/opt/reconciliation|$APP_DIR|g" "$SCRIPT_DIR/reconciliation.service" \
  > /etc/systemd/system/reconciliation.service
systemctl daemon-reload
systemctl enable reconciliation.service

echo "==> nginx"
sed "s|server_name _;|server_name $SERVER_NAME;|" \
  "$SCRIPT_DIR/nginx-reconciliation.conf" \
  > /etc/nginx/sites-available/reconciliation
ln -sfn /etc/nginx/sites-available/reconciliation /etc/nginx/sites-enabled/reconciliation
# Désactiver le site default s'il existe (évite le conflit sur :80)
if [[ -L /etc/nginx/sites-enabled/default ]]; then
  rm -f /etc/nginx/sites-enabled/default
fi
nginx -t
systemctl enable nginx
systemctl reload nginx

echo "==> Démarrage de l'application"
systemctl restart reconciliation.service

echo
echo "Installation terminée."
echo "  UI  : http://$SERVER_NAME/  (ou IP du serveur)"
echo "  API : http://$SERVER_NAME/health"
echo "  Logs: journalctl -u reconciliation -f"
echo
echo "Pensez à éditer $APP_DIR/.env puis : systemctl restart reconciliation"
