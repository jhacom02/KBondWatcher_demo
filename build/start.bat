@echo off
cd /d "%~dp0"
REM Edit Admin URL / public key before shipping. Keep this window open while using the app.
set KBOND_ADMIN_URL= https://cemetery-walking-characterized-considerable.trycloudflare.com
set KBOND_SIGNING_PUBLIC_KEY=O-4f7ECa2gLzqs-dUWlu52BibojbD5ZiubV7xfGjZbU=
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://127.0.0.1:8765/"
main.exe --serve
