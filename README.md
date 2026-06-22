# Co-Centre Interface

A Flask web application for the Food Co-Centre Sustainability Compass. It provides data tables, visualizations, manager workflows, and MongoDB-backed dataset management.

## Features

- Sustainability Compass visualization (`/d3viz`, `/d3viz2`)
- Searchable data tables with Excel export
- Manager and higher-manager dashboards
- MongoDB dataset storage (`local` database)
- Dataset import/export utilities

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)
- MongoDB Community Edition (runs locally on `127.0.0.1:27017`)

---

## Install Python

### Windows

1. Download the latest Python 3 installer from [python.org/downloads](https://www.python.org/downloads/windows/).
2. Run the installer.
3. On the first screen, check **Add python.exe to PATH**.
4. Click **Install Now**.
5. Open **Command Prompt** or **PowerShell** and verify:

```bash
python --version
pip --version
```

If `python` is not found, try:

```bash
py --version
py -m pip --version
```

### macOS

**Option A — Official installer (recommended for beginners)**

1. Download Python 3 from [python.org/downloads/macos](https://www.python.org/downloads/macos/).
2. Open the `.pkg` file and follow the installer.
3. Open **Terminal** and verify:

```bash
python3 --version
pip3 --version
```

**Option B — Homebrew**

```bash
brew install python
python3 --version
pip3 --version
```

On macOS, use `python3` and `pip3` in the commands below unless your system maps `python` to Python 3.

---

## Install MongoDB

The app connects to MongoDB at `mongodb://127.0.0.1:27017/` and uses the `local` database.

### Windows

1. Download **MongoDB Community Server** from [mongodb.com/try/download/community](https://www.mongodb.com/try/download/community).
2. Choose:
   - Version: latest stable (e.g. 7.0 or 8.0)
   - Platform: Windows
   - Package: MSI
3. Run the installer and choose **Complete** setup.
4. When prompted, install **MongoDB as a Service** so it starts automatically.
5. Optionally install **MongoDB Compass** (GUI) when offered.
6. Verify MongoDB is running:

```bash
mongosh
```

If `mongosh` is not recognized, MongoDB may still be running as a Windows service. Open Compass or check **Services** for `MongoDB`.

### macOS

**Option A — Homebrew (recommended)**

```bash
brew tap mongodb/brew
brew install mongodb-community
brew services start mongodb-community
mongosh
```

**Option B — Official installer**

1. Download MongoDB Community Server from [mongodb.com/try/download/community](https://www.mongodb.com/try/download/community).
2. Choose macOS and follow the installation guide in the MongoDB documentation.

Verify the server is running:

```bash
mongosh
```

---

## Project Setup

### 1. Clone or download this repository

```bash
cd interface
```

### 2. Create a virtual environment (recommended)

**Windows:**

```bash
python -m venv venv
venv\Scripts\activate
```

**macOS:**

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Python dependencies

```bash
pip install -r requirements.txt
```

On macOS, you may need `pip3` instead of `pip`.

---

## Import the Dataset

Dataset files are stored as JSON and imported into MongoDB using `dataset_io.py`.

### Import from a folder (recommended)

If you have a folder with one JSON file per collection (for example `exports/local_20260622_161228/`):

```bash
python dataset_io.py import exports/local_20260622_161228 --folder --replace
```

On macOS:

```bash
python3 dataset_io.py import exports/local_20260622_161228 --folder --replace
```

This imports all collections:

| Collection     | Description                          |
|----------------|--------------------------------------|
| `rolNLDraft`   | Main compass dataset                 |
| `Researcher`   | Researcher profiles                  |
| `Compass`      | Compass structure (Quadrant/Segment) |
| `Manager`      | Manager login accounts               |
| `HigherManager`| Higher manager accounts              |
| `PendingItems` | Items awaiting approval              |
| `description`  | Module descriptions                  |

- `--replace` clears each collection before import (use for a fresh setup).
- Omit `--replace` to merge data (update existing documents by `_id`, insert new ones).

### Import a single collection

```bash
python dataset_io.py import local.rolNLDraft.json
python dataset_io.py import local.rolNLDraft.json --collection rolNLDraft --replace
```

### Initialize admin user (optional)

If the `Manager` collection is empty and you need a default login:

```bash
python init_admin.py
```

Default credentials:

- Username: `admin`
- Password: `admin123`

Change this password before any production use.

---

## Run the Application

1. Make sure MongoDB is running.
2. Start the Flask server:

```bash
python app.py
```

3. Open your browser:

- [http://127.0.0.1:5000](http://127.0.0.1:5000)
- Compass visualization: [http://127.0.0.1:5000/d3viz2](http://127.0.0.1:5000/d3viz2)
- Manager login: [http://127.0.0.1:5000/manager](http://127.0.0.1:5000/manager)

Press `Ctrl+C` in the terminal to stop the server.

---

## Export the Dataset

Export the entire `local` database to a folder:

```bash
python dataset_io.py export --folder
```

Output is saved under `exports/local_YYYYMMDD_HHMMSS/` with one `.json` file per collection.

Other useful commands:

```bash
# Export a single collection
python dataset_io.py export --collection rolNLDraft

# Export all collections into one JSON file
python dataset_io.py export --all -o exports/full_backup.json
```

---

## Project Structure

```
interface/
├── app.py                 # Flask application
├── dataset_io.py          # MongoDB import/export utility
├── init_admin.py          # Create default manager account
├── requirements.txt       # Python dependencies
├── exports/               # Exported dataset folders
├── templates/             # HTML templates
├── static/                # CSS and assets
└── README.md
```

## Troubleshooting

| Problem | What to check |
|---------|----------------|
| `MongoDB not connected` | Is MongoDB running? Try `mongosh` |
| `python` not found (Windows) | Reinstall Python with **Add to PATH**, or use `py` |
| `pip` not found (macOS) | Use `pip3` and `python3` |
| Import fails | Confirm the JSON folder path exists and MongoDB is running |
| Port 5000 in use | Stop other apps on port 5000 or change the port in `app.py` |

## Development

The Flask application runs in debug mode by default:

- Automatic reload when code changes
- Detailed error messages in the browser
