@echo off
chcp 932 > /dev/null
echo 定例報告アプリのローカルサーバーを停止します...
for /f "tokens=5" %%a in ('netstat -ano -p TCP ^| findstr "127.0.0.1:8767" ^| findstr "LISTENING"') do (
  taskkill /PID %%a /F >/dev/null 2>&1
)
echo 停止しました。
timeout /t 2 >/dev/null
