@echo off

cd /d "C:\Users\Bruno\Desktop\Utopia"

for /f "tokens=5" %%a in ('netstat -aon ^| find ":8503" ^| find "LISTENING"') do (
    taskkill /PID %%a /F >nul 2>&1
)

start "" http://localhost:8503

python -m streamlit run app_utopia.py --server.port 8503

pause