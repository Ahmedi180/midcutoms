# Local Run Instructions

Quick steps to run the app locally on Windows (and an optional Linux command).

Prerequisites
- Python 3.10+ installed and available as `python` on PATH.
- Git (optional) if you cloned the repo.

Windows (PowerShell)

1. Create a virtual environment:

```powershell
python -m venv venv
```

2. Activate the venv (PowerShell):

```powershell
# If activation is blocked, allow the script for this session:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\n+venv\Scripts\Activate.ps1
# Or (from PowerShell) simply:
venv\Scripts\Activate
```

3. Install dependencies:

```powershell
pip install -r requirements.txt
```

4. Run the development server:

```powershell
python app.py
```

Open your browser at: http://127.0.0.1:5000

Windows (Command Prompt)

```cmd
python -m venv venv
venv\Scripts\activate.bat
pip install -r requirements.txt
python app.py
```

Linux / macOS (optional, for production testing)

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
# Run dev server
python app.py
# Or run with gunicorn for a more production-like server (Linux):
gunicorn app:app --bind 0.0.0.0:5000 --workers 2 --timeout 120
```

Notes & Troubleshooting
- If you see errors about activation or execution policies on PowerShell, run the `Set-ExecutionPolicy` line shown above in the same PowerShell session.
- Ensure `requirements.txt` is present in the project root. If `pip install` fails, check your internet connection and Python version.
- The default web port is `5000`. If the port is in use change the port in `app.py` or set environment variable accordingly.

If you'd like, I can also add a one-click `run.bat` for Windows or a `Procfile`/`start.sh` for easy hosting. Tell me which you prefer.
