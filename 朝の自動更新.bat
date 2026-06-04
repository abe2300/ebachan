@echo off
chcp 932 > /dev/null
cd /d "%~dp0"

echo ===============================================
echo  定例報告 朝の自動更新ルーティン
echo ===============================================
echo.
echo このBATは以下を実行します:
echo   1. Claude Codeを起動
echo   2. (手動) Claudeに「定例報告」を依頼
echo   3. (手動) Claudeが index.html を更新
echo   4. Enterキー押下後、自動でGitHubにpush
echo   5. Netlifyが自動デプロイ
echo.
echo ===============================================
echo.

rem Claude Codeを開く (インストール先により変更)
echo [情報] Claude Codeを起動します...
start "" "claude" 2>/dev/null
if errorlevel 1 (
  echo [情報] claudeコマンドが見つからないため、ブラウザで起動します
  start "" "https://claude.ai/"
)

echo.
echo ===============================================
echo  作業完了後、このウィンドウでEnterキーを押してください
echo  → 自動でGitHubにpushしてNetlifyに公開します
echo ===============================================
pause >/dev/null

rem 変更を確認してgit pushを実行
echo.
echo [情報] 変更を確認中...
git status --short

echo.
echo [情報] 変更があればGitHubにpushします...
git add -A

git diff --cached --quiet
if errorlevel 1 (
  for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
    set DATE_STR=%%a/%%b/%%c
  )
  git commit -m "Auto morning update: %DATE_STR%"
  git push
  if errorlevel 0 (
    echo.
    echo ===============================================
    echo  Push成功! Netlifyで自動公開中...
    echo  約30秒後: https://teirei-houkoku.netlify.app
    echo ===============================================
  ) else (
    echo [エラー] pushに失敗しました
  )
) else (
  echo [情報] 変更がないためpushしません
)

echo.
pause
