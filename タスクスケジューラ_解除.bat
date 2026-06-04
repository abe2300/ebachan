@echo off
chcp 932 > /dev/null

set TASK_NAME=定例報告_朝の自動起動

echo ===============================================
echo  Windows タスクスケジューラ 解除ツール
echo ===============================================
echo.
echo 「%TASK_NAME%」を解除しますか?
choice /M "解除する"
if errorlevel 2 (
  echo キャンセルしました
  pause
  exit /b
)

schtasks /delete /tn "%TASK_NAME%" /f

if errorlevel 0 (
  echo.
  echo タスクを解除しました
) else (
  echo.
  echo タスクが見つからないか、解除に失敗しました
)
echo.
pause
