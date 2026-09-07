# Co-Centre Interface

Flask web app for the **Food Co-Centre Sustainability Compass**: interactive visualization, searchable data tables, researcher overlays, and manager workflows.

---

## Two ways to run

| Method | Best for | Needs |
|---|---|---|
| **A. Docker Compose** | Local demo / server with Docker permission | Docker Desktop or Docker Engine |
| **B. Python (`app.py`)** | School server without Docker sudo | Python 3.10+, optional MongoDB |

If MongoDB is unavailable, the app still starts and falls back to local JSON under `exports/local_20260622_161257/` for browsing, filters, and researcher views. Persistent writes work best with MongoDB.

---

## Method A — Docker Compose

### 1. Install Docker

- macOS/Windows: [Docker Desktop](https://www.docker.com/products/docker-desktop/)
- Linux: Docker Engine + Compose plugin

### 2. Start

```bash
cd /path/to/interface

# free port 5001 if needed
# lsof -t -iTCP:5001 -sTCP:LISTEN | xargs kill 2>/dev/null

docker compose up --build -d
docker compose ps
docker compose logs -f web
```

Open: **http://localhost:5001**

Compass UI: **http://localhost:5001/d3viz**

### 3. Stop

```bash
docker compose down
```

### Docker notes

- `web` + `mongo` start together; web waits for Mongo, then can import `exports/local_20260622_161257`.
- Default admin (first start, if `INIT_ADMIN=true`): `admin` / `admin123`
- After the first successful import, set in `docker-compose.yml`:
  - `DATASET_IMPORT_REPLACE: "false"` (or clear `DATASET_IMPORT_PATH`)
  - so restarts do not wipe edited data

### Docker permission denied on a shared server

```text
permission denied ... /var/run/docker.sock
```

Means your user is not in the `docker` group and has no `sudo`. Ask an admin:

```bash
sudo usermod -aG docker YOUR_USERNAME
```

Then log out/in and retry `docker compose up --build -d`.  
If you cannot get Docker access, use **Method B** below.

---

## Method B — Python (`python app.py`)

### 1. Create a virtualenv and install deps

```bash
cd /path/to/interface

python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Use `python3 -m pip` / `python -m pip`. Avoid bare `pip` if the system `command-not-found` helper is broken.

### 2. (Optional) MongoDB

If MongoDB is running locally:

```bash
export MONGO_URI=mongodb://127.0.0.1:27017/
```

If Mongo is missing, you will see a short “MongoDB connection failed …” message, then Flask still starts and uses JSON fallbacks.

### 3. Run

```bash
export FLASK_HOST=0.0.0.0
export FLASK_PORT=5001
export FLASK_DEBUG=false

python app.py
```

Open:

- Local machine: **http://127.0.0.1:5001**
- Same server LAN/public IP (if firewall allows): **http://SERVER_IP:5001**

### 4. Keep it running after SSH disconnect (server)

```bash
cd ~/path/to/interface
source .venv/bin/activate
export FLASK_HOST=0.0.0.0 FLASK_PORT=5001 FLASK_DEBUG=false
nohup python app.py > app.log 2>&1 &

curl -I http://127.0.0.1:5001/
tail -n 50 app.log
```

If `Address already in use`, something is already on that port:

```bash
lsof -iTCP:5001 -sTCP:LISTEN    # or: ss -lntp | grep 5001
kill PID
```

Or use another port:

```bash
export FLASK_PORT=5002
python app.py
```

### 5. Access from your laptop when the firewall blocks the server IP

On your laptop:

```bash
ssh -N -L 5002:127.0.0.1:5001 yunan@aginsight.ucd.ie
```

Keep that window open, then browse **http://127.0.0.1:5002**  
(If the app listens on 5002 on the server, forward `5002:127.0.0.1:5002` instead.)

`http://127.0.0.1:5001` in your laptop browser only works if the app runs on the laptop, or an SSH tunnel is active.

---

## Deploy on a school server (summary)

1. Upload code (from your laptop):

```bash
rsync -avz --exclude '.git' --exclude '__pycache__' --exclude '.venv' --exclude '*.bak' \
  /path/to/interface/ \
  USER@SERVER:~/co-centre-interface/
```

2. Prefer **Method A** if you have Docker permission; otherwise **Method B**.
3. Ask IT to reverse-proxy (recommended for others):

```text
https://aginsight.ucd.ie  →  http://127.0.0.1:5001
```

4. Share the public URL (or `http://SERVER_IP:PORT` if the firewall allows it).

---

## Main URLs

| Path | Purpose |
|---|---|
| `/` | Home / dataset entry |
| `/d3viz` or `/d3viz2` | Sustainability Compass (current UI) |
| `/intro` | Compass introduction |
| `/table/__ALL__` | Full data table + filters |
| `/table/<module>?field=Factor` | Table scoped to a module |
| `/manager` | Manager tools |
| `/map` | Location map |

---

## What the important files are

### Application

| File / folder | Role |
|---|---|
| `app.py` | Main Flask app, APIs, Mongo/JSON fallbacks |
| `templates/` | HTML pages (`d3viz2.html`, `table.html`, …) |
| `static/` | CSS/JS/assets |
| `requirements.txt` | Python dependencies |
| `dataset_io.py` | Export/import Mongo collections ↔ JSON |
| `init_admin.py` | Create default admin user |
| `scripts/build_compass_from_v2.py` | Build Compass tree JSON from flat data items |

### Docker

| File | Role |
|---|---|
| `Dockerfile` | Image build (Python 3.11, deps, port 5001) |
| `docker-compose.yml` | `mongo` + `web`, env, import path, ports |
| `docker-entrypoint.sh` | Wait for Mongo → import → init admin → `app.py` |
| `.dockerignore` | Files excluded from the image |

### Data (Mongo collections ↔ JSON)

Imported from `exports/local_20260622_161257/` when using Docker import:

| File | Collection / use |
|---|---|
| `Compass.json` | Compass labels / structure (flat Q/S/F rows in current pack) |
| `rolNLDraft.json` | Main table dataset |
| `Researcher.json` | Researcher one-hot Factor columns + filters |
| `description.json` | Hover / definition text for Compass modules |
| `_manifest.json` | Export metadata / counts |
| `Manager.json`, `HigherManager.json`, … | Auth / workflow collections |

Root helpers / fallbacks:

| File | Role |
|---|---|
| `compass.json` | Tree-shaped Compass fallback |
| `description.json` / `food_system_full.json` | Description fallbacks |
| `local.Researcher.json` | Alternate researcher dump |
| `exports/compass_dataitems_v2_final.json` | Source flat data items used to rebuild Compass |

### Environment variables

| Variable | Default | Meaning |
|---|---|---|
| `MONGO_URI` | `mongodb://127.0.0.1:27017/` | Mongo connection |
| `FLASK_HOST` | `127.0.0.1` (Docker sets `0.0.0.0`) | Bind address |
| `FLASK_PORT` | `5001` | HTTP port |
| `FLASK_DEBUG` | `true` locally / `false` in Docker | Debug mode |
| `DATASET_IMPORT_PATH` | (Docker) export folder | Auto-import on container start |
| `DATASET_IMPORT_REPLACE` | `true` in compose | Wipe target collections before import |
| `INIT_ADMIN` | `true` in compose | Ensure default admin exists |

---

## Data model (how pieces link)

- **Compass / table rows**: `Quadrant` → `Segment` → `Factor` (+ `Location`, `Title`, `Url`, …)
- **Descriptions**: names in `description.json` must match Compass labels (`&` / `and` aliases supported in code)
- **Researchers**: one-hot columns named after Factors; Platform Alignment / Institution / Career Stage drive the Research Outcomes filters

Rebuild Compass tree from flat v2 items (optional):

```bash
python3 scripts/build_compass_from_v2.py
# optional: python3 scripts/build_compass_from_v2.py --mongo
```

Import/export with Mongo running:

```bash
python dataset_io.py export --folder
python dataset_io.py import exports/local_20260622_161257 --folder --replace
```

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `MongoDB connection failed` then app starts | No local Mongo | OK for read-only JSON mode; install Mongo for full writes |
| `Address already in use` / port busy | Old `app.py` still running | `lsof`/`ss` + `kill`, or change `FLASK_PORT` |
| `permission denied ... docker.sock` | Not in `docker` group | Ask admin, or use Method B |
| `yunan is not in the sudoers file` | No admin rights | Cannot self-install Docker; use Python mode or ask IT |
| Browser `127.0.0.1:5001` fails on your laptop | App runs on the server, not your PC | Use `http://SERVER_IP:PORT` or SSH tunnel |
| Page spins forever from off-campus | Firewall blocks the port | Use SSH tunnel or ask IT for reverse proxy / VPN |

---

## Development notes

- Primary Compass UI is `templates/d3viz2.html` (served by `/d3viz` and `/d3viz2`).
- Table filters use `/api/compass/structure` and `/api/dataset/filter` (Mongo or JSON).
- Researcher bubbles/filters use `/api/researchers/*` (Mongo or `Researcher.json`).

---

## License / project

Internal Co-Centre tooling for Ireland / UK food-system sustainability outcomes visualization and data search.
