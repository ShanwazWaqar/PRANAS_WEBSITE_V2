@echo off
set DATABASE_URL=postgresql://postgres:root@localhost:5432/pranas_db
set FLASK_APP=App.py
set FLASK_ENV=development
.venv\Scripts\python.exe App.py 