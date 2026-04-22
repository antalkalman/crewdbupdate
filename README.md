# CrewDB Updater

Internal crew database management tool for Pioneer Pictures. Built with FastAPI + AG Grid.

---

## Pages

| URL | Description |
|---|---|
| `/` | Redirect to Match |
| `/match` | Match SFlist crew against CrewIndex |
| `/crew_explorer` | Browse and filter CrewIndex |
| `/registry` | Edit CrewRegistry |
| `/sf_issues` | Validate SFlist for missing fields, fees, signatures |
| `/titles` | Manage title/department mappings |

---

## Local Development

### Start
```bash
export $(grep -v '^#' .env.title_mapper | xargs) && python3 -m uvicorn backend.app:app --host 127.0.0.1 --port 8000
```
Or double-click `Start crew_db_updater.command`.

Open: http://127.0.0.1:8000

### Stop
```bash
lsof -ti:8000 | xargs kill -9
```
Or double-click `Stop crew_db_updater.command`.

---

## Production (AWS EC2)

Live at: https://honinbo.net  
Instance: EC2 t3.small, Ubuntu 22.04, eu-north-1  
Key pair: stored locally at `~/Downloads/your-key.pem`

### SSH in
```bash
ssh -i ~/Downloads/your-key.pem ubuntu@13.49.75.170
```

### Start / Stop / Restart app
```bash
sudo systemctl start crewdb
sudo systemctl stop crewdb
sudo systemctl restart crewdb
```

### Check app status and logs
```bash
sudo systemctl status crewdb
sudo journalctl -u crewdb -n 50 --no-pager
```

### Deploy code update
```bash
cd ~/crewdbupdate
git pull
sudo systemctl restart crewdb
```

---

## Fresh Server Deployment

Use this when setting up a new EC2 instance from scratch.

### Prerequisites
- Ubuntu 22.04 LTS EC2 instance (t3.small recommended)
- Security group: ports 22, 80, 443 open
- DNS A record pointing your domain to the instance IP

### Step 1 — Upload data files (from your Mac)
```bash
scp -i ~/Downloads/your-key.pem .env.title_mapper ubuntu@<IP>:~/
scp -i ~/Downloads/your-key.pem -r New_Master_Database ubuntu@<IP>:~/crewdbupdate/
```

### Step 2 — Clone and set up (on the instance)
```bash
git clone -b aws https://github.com/antalkalman/crewdbupdate.git
cd crewdbupdate
mv ~/.env.title_mapper .
bash deploy/setup.sh
```

App will be live at `http://<IP>`.

### Step 3 — Enable HTTPS (once DNS resolves)
```bash
cd ~/crewdbupdate
bash deploy/setup_ssl.sh
```

App will be live at `https://honinbo.net`.

---

## Data Files (not in repo)

These live on the server only and must be uploaded manually after a fresh deploy:

| File/Folder | Description |
|---|---|
| `.env.title_mapper` | API credentials and config |
| `New_Master_Database/*.xlsx` | CrewIndex, CrewRegistry, TitleMap, GCMID_Map, projects.json |
| `New_Master_Database/Historical/` | 35 historical SFlist files |

---

## Deploy Configs (in repo under `deploy/`)

| File | Description |
|---|---|
| `setup.sh` | Full server setup: packages, venv, nginx, systemd |
| `setup_ssl.sh` | Let's Encrypt SSL cert + swap in SSL nginx config |
| `crewdb.service` | systemd unit file |
| `nginx.crewdb` | nginx config (HTTP) |
| `nginx.crewdb.ssl` | nginx config (HTTPS, used after SSL setup) |
