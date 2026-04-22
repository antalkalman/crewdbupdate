#!/bin/bash
# CrewDB Updater — SSL setup script
# Run after setup.sh, once DNS is pointing to this server.
#   bash deploy/setup_ssl.sh

set -e
DOMAIN="honinbo.net"
REPO_DIR="/home/ubuntu/crewdbupdate"

echo "==> Installing certbot..."
sudo apt install -y certbot python3-certbot-nginx

echo "==> Obtaining SSL certificate for $DOMAIN..."
sudo certbot certonly --nginx \
    -d "$DOMAIN" -d "www.$DOMAIN" \
    --non-interactive --agree-tos \
    --email "admin@$DOMAIN"

echo "==> Swapping in SSL nginx config..."
sudo cp "$REPO_DIR/deploy/nginx.crewdb.ssl" /etc/nginx/sites-available/crewdb
sudo nginx -t
sudo systemctl restart nginx

echo ""
echo "==> Done! App is live at https://$DOMAIN"
echo "    Auto-renewal is handled by certbot's systemd timer (runs twice daily)."
