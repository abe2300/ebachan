@echo off
chcp 932 > /dev/null
cd /d "%~dp0"

if not exist "公開用" mkdir "公開用"

echo 公開用フォルダにPWA必須ファイルをコピーします...
copy /Y "index.html"    "公開用" >/dev/null
copy /Y "manifest.json" "公開用" >/dev/null
copy /Y "sw.js"         "公開用" >/dev/null
copy /Y "icon-192.png"  "公開用" >/dev/null
copy /Y "icon-512.png"  "公開用" >/dev/null

echo.
echo === 公開用フォルダの中身 ===
dir /B "公開用"

echo.
echo 完了しました。
echo Netlifyダッシュボードの「Production deploys」エリアに
echo 「公開用」フォルダをドラッグしてください。
echo.
pause
