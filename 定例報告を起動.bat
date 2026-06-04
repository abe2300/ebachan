@echo off
chcp 932 > /dev/null
cd /d "%~dp0"

rem 既にローカルサーバーが起動済みか確認（127.0.0.1:8767 LISTENING）
netstat -ano -p TCP 2>/dev/null | findstr "127.0.0.1:8767" | findstr "LISTENING" >/dev/null
if errorlevel 1 (
  rem サーバー未起動 → pythonw で server.py を完全バックグラウンド起動
  start "" pythonw "%~dp0server.py"
  rem 立ち上がりを少し待つ
  ping -n 2 127.0.0.1 >/dev/null
)

rem Microsoft Edge をアプリモード（URLバーなし）で起動
set "EDGE=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" set "EDGE=C:\Program Files\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE%" (
  echo Microsoft Edge が見つかりません。
  pause
  exit /b 1
)

start "" "%EDGE%" --app=http://localhost:8767/index.html
exit /b 0
