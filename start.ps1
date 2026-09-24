# One-command local start: .\start.ps1
if (-not (Test-Path .\venv)) { python -m venv venv }
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:FLASK_DEBUG = "1"
python app.py
