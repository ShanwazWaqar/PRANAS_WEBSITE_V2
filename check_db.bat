@echo off
set DATABASE_URL=postgresql://postgres:root@localhost:5432/pranas_db
.venv\Scripts\python.exe check_db.py 