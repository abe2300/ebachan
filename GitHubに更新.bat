@echo off
chcp 932 > /dev/null
cd /d "%~dp0"
echo ===============================================
echo  GitHubに更新をPushしてGitHub Pagesに自動公開
echo ===============================================
echo.

rem 変更があるか確認
git diff --quiet
if errorlevel 1 (
  echo [情報] 変更を検知しました。Commitします...
) else (
  git diff --cached --quiet
  if errorlevel 1 (
    echo [情報] ステージング済みの変更があります...
  ) else (
    echo [情報] 変更はありません。終了します。
    pause
    exit /b 0
  )
)

rem 全変更をadd
git add -A

rem コミットメッセージに日時を含める
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do (
  set DATE_STR=%%a/%%b/%%c
)
for /f "tokens=1-2 delims=:" %%a in ("%time%") do (
  set TIME_STR=%%a:%%b
)
git commit -m "Update: %DATE_STR% %TIME_STR%"

rem GitHubにpush
echo.
echo [情報] GitHubにpushします...
git push

if errorlevel 0 (
  echo.
  echo ===============================================
  echo  Push成功! GitHub Pagesが自動デプロイ中...
  echo  約30秒で https://abe2300.github.io/ebachan/ に反映されます
  echo ===============================================
) else (
  echo.
  echo [エラー] Pushに失敗しました。
)

echo.
pause
