#!/usr/bin/env bash
# ==============================================================================
# Gola PDF - Automated 1-Click Production Deployment Script
# Optimized for $14/mo Linux VPS (Ubuntu 22.04 / 24.04 LTS or Debian 11/12)
# ==============================================================================

set -euo pipefail

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}================================================================${NC}"
echo -e "${GREEN}       Gola PDF - Automated Production Deployment Engine        ${NC}"
echo -e "${BLUE}================================================================${NC}"

# 1. Check Root Privileges
if [ "$EUID" -ne 0 ]; then
  echo -e "${RED}[ERROR] Please run this script with sudo or as root.${NC}"
  echo "Usage: sudo bash deploy.sh [your_domain.com] [admin_email@example.com]"
  exit 1
fi

DOMAIN="${1:-}"
EMAIL="${2:-}"

if [ -z "$DOMAIN" ]; then
  read -rp "Enter your primary domain (e.g., golapdf.com): " DOMAIN
fi

if [ -z "$EMAIL" ]; then
  read -rp "Enter your admin email for free SSL certificate: " EMAIL
fi

APP_DIR="/var/www/golapdf"
VENV_DIR="/opt/golapdf_env"

echo -e "\n${YELLOW}[1/7] Updating system packages and installing dependencies...${NC}"
apt-get update -y
apt-get install -y --no-install-recommends \
  python3 \
  python3-pip \
  python3-venv \
  python3-dev \
  build-essential \
  tesseract-ocr \
  tesseract-ocr-eng \
  ghostscript \
  nginx \
  certbot \
  python3-certbot-nginx \
  curl \
  git \
  ufw

echo -e "\n${YELLOW}[2/7] Setting up application directory at ${APP_DIR}...${NC}"
mkdir -p "$APP_DIR"
cp -r ./* "$APP_DIR/" || true
mkdir -p "$APP_DIR/ephemeral_gcs_storage"

# Create application system user if not exists
if ! id -u golapdf >/dev/null 2>&1; then
  useradd -r -s /bin/false -d "$APP_DIR" golapdf
fi
chown -R golapdf:www-data "$APP_DIR"
chmod -R 775 "$APP_DIR/ephemeral_gcs_storage"

echo -e "\n${YELLOW}[3/7] Creating Python Virtual Environment at ${VENV_DIR}...${NC}"
python3 -m venv "$VENV_DIR"
"$VENV_DIR/bin/pip" install --upgrade pip setuptools wheel
"$VENV_DIR/bin/pip" install -r "$APP_DIR/requirements.txt"

echo -e "\n${YELLOW}[4/7] Configuring Systemd Service (/etc/systemd/system/golapdf.service)...${NC}"
cat > /etc/systemd/system/golapdf.service <<EOF
[Unit]
Description=Gola PDF Production Application Server
After=network.target

[Service]
Type=simple
User=golapdf
Group=www-data
WorkingDirectory=${APP_DIR}
Environment="PATH=${VENV_DIR}/bin:/usr/local/bin:/usr/bin"
Environment="PORT=8080"
ExecStart=${VENV_DIR}/bin/uvicorn app:app --host 127.0.0.1 --port 8080 --workers 2 --log-level warning
Restart=always
RestartSec=5s
KillSignal=SIGQUIT
StandardOutput=journal
StandardError=journal
SyslogIdentifier=golapdf

# Security Hardening
PrivateTmp=true
ProtectSystem=full
ProtectHome=true

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable golapdf
systemctl restart golapdf

echo -e "\n${YELLOW}[5/7] Configuring Nginx Reverse Proxy with Gzip & Security...${NC}"
cat > "/etc/nginx/sites-available/${DOMAIN}" <<EOF
server {
    listen 80;
    listen [::]:80;
    server_name ${DOMAIN} www.${DOMAIN};

    client_max_body_size 100M;

    # Gzip Compression
    gzip on;
    gzip_vary on;
    gzip_min_length 1024;
    gzip_proxied any;
    gzip_comp_level 6;
    gzip_types
        text/plain
        text/css
        text/xml
        text/javascript
        application/json
        application/javascript
        application/x-javascript
        application/xml
        application/xml+rss
        image/svg+xml;

    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_http_version 1.1;
        proxy_set_header Upgrade \$http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;

        proxy_connect_timeout 180s;
        proxy_send_timeout 180s;
        proxy_read_timeout 180s;
        proxy_buffering off;
    }
}
EOF

ln -sf "/etc/nginx/sites-available/${DOMAIN}" "/etc/nginx/sites-enabled/"
rm -f /etc/nginx/sites-enabled/default || true

nginx -t
systemctl reload nginx

echo -e "\n${YELLOW}[6/7] Configuring Firewall (UFW)...${NC}"
ufw allow OpenSSH || true
ufw allow 'Nginx Full' || true
ufw --force enable || true

echo -e "\n${YELLOW}[7/7] Requesting Let's Encrypt SSL Certificate...${NC}"
if certbot --nginx --non-interactive --agree-tos --email "$EMAIL" -d "$DOMAIN" -d "www.$DOMAIN" --redirect; then
  echo -e "${GREEN}[SUCCESS] SSL Certificate installed and HTTPS redirect enabled!${NC}"
else
  echo -e "${YELLOW}[NOTICE] Let's Encrypt could not verify DNS yet. Ensure your domain's A-record points to this server's IP, then run: sudo certbot --nginx -d ${DOMAIN} -d www.${DOMAIN}${NC}"
fi

# Verify local application health
sleep 2
echo -e "\n${BLUE}Verifying local deployment health...${NC}"
if curl -fs "http://127.0.0.1:8080/api/health" > /dev/null; then
  echo -e "${GREEN}[PASS] Gola PDF is running healthy on 127.0.0.1:8080!${NC}"
else
  echo -e "${RED}[WARN] Application is starting up. Check status with: systemctl status golapdf${NC}"
fi

echo -e "\n${GREEN}================================================================${NC}"
echo -e "${GREEN}   🎉 Gola PDF is Successfully Deployed and LIVE!               ${NC}"
echo -e "${GREEN}   URL: https://${DOMAIN}                                       ${NC}"
echo -e "${GREEN}================================================================${NC}"
echo -e "Useful management commands:"
echo -e "  - View service status:  ${BLUE}systemctl status golapdf${NC}"
echo -e "  - View live logs:       ${BLUE}journalctl -u golapdf -f${NC}"
echo -e "  - Restart server:       ${BLUE}systemctl restart golapdf${NC}"
echo -e "  - Reload Nginx:         ${BLUE}systemctl reload nginx${NC}"
