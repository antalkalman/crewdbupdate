#!/bin/bash
# CrewDB Updater — server setup script
# Run once on a fresh Ubuntu 22.04 EC2 instance from ~/crewdbupdate:
#   bash deploy/setup.sh

set -e
REPO_DIR="/home/ubuntu/crewdbupdate"

echo "==> Installing system packages..."
sudo apt update && sudo apt install -y python3-pip python3-venv nginx git

echo "==> Setting up Python environment..."
cd "$REPO_DIR"
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

echo "==> Installing systemd service..."
sudo cp "$REPO_DIR/deploy/crewdb.service" /etc/systemd/system/crewdb.service
sudo systemctl daemon-reload
sudo systemctl enable crewdb
sudo systemctl start crewdb

echo "==> Installing nginx config..."
sudo cp "$REPO_DIR/deploy/nginx.crewdb" /etc/nginx/sites-available/crewdb
sudo ln -sf /etc/nginx/sites-available/crewdb /etc/nginx/sites-enabled/crewdb
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl restart nginx

echo ""
echo "==> Done! App should be live at http://$(curl -s ifconfig.me)"
echo ""
echo "    NOTE: Upload your data files and .env.title_mapper before starting:"
echo "    scp -i your-key.pem -r New_Master_Database ubuntu@<IP>:~/crewdbupdate/"
echo "    scp -i your-key.pem .env.title_mapper ubuntu@<IP>:~/crewdbupdate/"
