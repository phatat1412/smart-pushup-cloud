@echo off
title Push Code Len GitHub - Smart Push Up
color 0A
cd /d  D:\IOT\Smart Push Up\cloud_server
echo ==============================================================
echo   DANG DAY CODE LEN GITHUB: phatat1412/smart-pushup-cloud
echo ==============================================================
echo.
git push -u origin main
echo.
if %ERRORLEVEL% EQU 0 (
    echo ==============================================================
    echo [THANH CONG] Da day toan bo code len GitHub!
    echo Bay gio quay lai Render bam nut: Manual Deploy -> Deploy latest commit
    echo ==============================================================
) else (
    echo [LUU Y] Neu co cua so GitHub hien len hoi dang nhap, ban hay bam Sign in with your browser nhe.
)
echo.
pause
