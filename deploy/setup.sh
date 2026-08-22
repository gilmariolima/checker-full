#!/usr/bin/env bash
# Roda isso DENTRO da VM Oracle (via SSH), como usuário ubuntu.
# Uso: bash setup.sh
set -euo pipefail

REPO_URL="https://github.com/gilmariolima/checker-full.git"
APP_DIR="$HOME/checker-full"

echo "==> Atualizando o sistema e instalando dependências..."
sudo apt-get update -y
sudo apt-get install -y python3-venv python3-pip git

if [ -d "$APP_DIR" ]; then
  echo "==> Repositório já existe, atualizando..."
  cd "$APP_DIR"
  git pull origin main
else
  echo "==> Clonando o repositório..."
  git clone "$REPO_URL" "$APP_DIR"
  cd "$APP_DIR"
fi

echo "==> Criando ambiente virtual e instalando dependências Python..."
python3 -m venv .venv
./.venv/bin/pip install --upgrade pip
./.venv/bin/pip install -r requirements.txt

echo "==> Instalando o serviço systemd..."
sudo cp deploy/checker-full.service /etc/systemd/system/checker-full.service
sudo systemctl daemon-reload
sudo systemctl enable checker-full
sudo systemctl restart checker-full

echo "==> Abrindo a porta 8000 no firewall local (iptables/ufw, se ativo)..."
if command -v ufw >/dev/null 2>&1; then
  sudo ufw allow 8000/tcp || true
fi

echo ""
echo "==> Pronto! Status do serviço:"
sudo systemctl status checker-full --no-pager

echo ""
echo "IMPORTANTE: além disso, abra a porta 8000 na Oracle Cloud Console em"
echo "  Networking -> Virtual Cloud Networks -> (sua VCN) -> Security Lists"
echo "  -> Add Ingress Rule: Source 0.0.0.0/0, Destination Port 8000, TCP"
echo ""
echo "Depois disso, acesse: http://<IP-PUBLICO-DA-VM>:8000"
