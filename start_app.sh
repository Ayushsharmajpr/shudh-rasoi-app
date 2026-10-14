#!/usr/bin/env bash
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt
echo; echo "Starting Shudh Rasoi... first sign-in is admin / admin. Open http://127.0.0.1:5000 in your browser."
python3 app.py
