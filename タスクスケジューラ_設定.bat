@echo off
chcp 932 > /dev/null
cd /d "%~dp0"

echo ===============================================
echo  Windows タスクスケジューラ 設定ツール
echo  定例報告アプリの毎朝7時自動起動を設定します
echo ===============================================
echo.

set TASK_NAME=定例報告_朝の自動起動
set BAT_PATH=%~dp0朝の自動更新.bat

echo 設定内容:
echo   - タスク名: %TASK_NAME%
echo   - 実行ファイル: %BAT_PATH%
echo   - 実行時刻: 毎日 7:00
echo.

choice /M "この内容で設定しますか?"
if errorlevel 2 (
  echo キャンセルしました
  pause
  exit /b
)

rem 既存タスクがあれば削除
schtasks /query /tn "%TASK_NAME%" >/dev/null 2>&1
if not errorlevel 1 (
  echo [情報] 既存のタスクを削除します...
  schtasks /delete /tn "%TASK_NAME%" /f
)

rem タスクを作成 (毎日 7:00 実行)
schtasks /create /tn "%TASK_NAME%" /tr ""%BAT_PATH%"" /sc daily /st 07:00 /f

if errorlevel 0 (
  echo.
  echo ===============================================
  echo  設定完了!
  echo  毎朝 7:00 に「朝の自動更新.bat」が起動します
  echo ===============================================
  echo.
  echo 確認方法:
  echo   1. Windowsスタートメニュー > タスクスケジューラ
  echo   2. 「タスクスケジューラ ライブラリ」をクリック
  echo   3. 「%TASK_NAME%」を探す
  echo.
  echo 解除する場合:
  echo   このフォルダの「タスクスケジューラ_解除.bat」を実行
) else (
  echo.
  echo [エラー] タスク作成に失敗しました
  echo 管理者権限が必要な可能性があります
  echo BATファイルを右クリック → 管理者として実行 で再試行してください
)

echo.
pause
