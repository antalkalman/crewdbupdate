#!/bin/bash
# CrewDB Updater — Basic Auth setup script
# Run after setup_ssl.sh. Requires HTTPS to be working first.
#   bash deploy/setup_auth.sh

set -e
REPO_DIR="/home/ubuntu/crewdbupdate"

echo "==> Installing apache2-utils (for htpasswd)..."
sudo apt install -y apache2-utils

echo ""
echo "==> Create login credentials for CrewDB"
read -p "Username: " USERNAME </dev/tty
sudo htpasswd -c /etc/nginx/.htpasswd "$USERNAME" </dev/tty

echo ""
echo "==> Installing auth nginx config..."
sudo cp "$REPO_DIR/deploy/nginx.crewdb.auth" /etc/nginx/sites-available/crewdb
sudo nginx -t
sudo systemctl restart nginx

echo ""
echo "==> Done! https://honinbo.net is now password protected."
echo "    To add more users: sudo htpasswd /etc/nginx/.htpasswd <username>"
echo "    To change password: sudo htpasswd /etc/nginx/.htpasswd <username>"
echo "    To remove a user:   sudo htpasswd -D /etc/nginx/.htpasswd <username>"
