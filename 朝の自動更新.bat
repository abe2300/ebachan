@echo off
chcp 932 > /dev/null
cd /d "%~dp0"

rem ===== デスクトップ通知 =====
powershell -ExecutionPolicy Bypass -Command " = New-BurntToastNotification -Text '定例報告アプリ', '朝の自動更新時間です。Claude Codeで「定例報告」と入力してください' 2>"

rem 通知が失敗してもメッセージボックスで案内
if errorlevel 1 (
  powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')|Out-Null; [System.Windows.Forms.MessageBox]::Show('Claude Codeで「定例報告」と入力してHTMLを更新してください。完了後、このウィンドウでEnterキーを押すとGitHubにpushします。','定例報告 朝の自動更新','OK','Information')"
)

echo ===============================================
echo  定例報告 朝の自動更新ルーティン
echo ===============================================
echo.
echo 【作業手順】
echo   1. Claude Codeを起動 (本BATが自動起動)
echo   2. 「定例報告」と入力してHTMLを更新してもらう
echo   3. 完了したらこのウィンドウでEnterキーを押す
echo   4. 自動でGitHubにpushしてGitHub Pagesに公開
echo.
echo ===============================================
echo.

rem Claude Codeを開く
echo [情報] Claude Codeを起動します...
start "" "claude" 2>/dev/null
if errorlevel 1 (
  echo [情報] claudeコマンドが見つからないため、ブラウザで起動します
  start "" "https://claude.ai/new"
)

echo.
echo ===============================================
echo  作業完了後、このウィンドウでEnterキーを押してください
echo  -> 自動でGitHubにpushしてGitHub Pagesに公開します
echo ===============================================
pause >/dev/null

rem 変更を確認
echo.
echo [情報] 変更を確認中...
git status --short

rem 変更がなければスキップ
git diff --quiet
set DIFF_RESULT=%errorlevel%
git diff --cached --quiet
set CACHED_RESULT=%errorlevel%

if %DIFF_RESULT% equ 0 if %CACHED_RESULT% equ 0 (
  echo.
  echo [情報] 変更がありません。終了します。
  pause
  exit /b 0
)

rem 変更があればpush
echo.
echo [情報] GitHubにpushします...
git add -A

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
  set DATE_STR=%%a/%%b/%%c
)
git commit -m "Auto morning update: %DATE_STR%"
git push

if errorlevel 0 (
  echo.
  echo ===============================================
  echo  Push成功!
  echo  数秒後: https://abe2300.github.io/ebachan/
  echo ===============================================
) else (
  echo.
  echo [エラー] pushに失敗しました
)

echo.
pause
