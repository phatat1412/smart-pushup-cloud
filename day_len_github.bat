@echo off
title Day Code Len GitHub - Smart Push-Up Cloud
color 0A

echo ==============================================================================
echo        TU DONG DAY MA NGUON LEN GITHUB DE DEPLOY RENDER 24/7
echo ==============================================================================
echo.
cd /d  D:\IOT\Smart Push Up\cloud_server

echo [BUOC 1] Nhap link GitHub repository cua ban.
echo (Vi du: https://github.com/nguyentanphat/smart-pushup-cloud.git)
echo.
set /p REPO_URL=Dan link GitHub vao day roi bam Enter: 

if  %REPO_URL%== (
    echo.
    echo [LOI] Ban chua nhap link GitHub! Vui long chay lai file.
    pause
    exit /b
)

echo.
echo Dang ket noi toi GitHub: %REPO_URL% ...
git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%
git branch -M main

echo.
echo Dang day ma nguon len GitHub (git push -u origin main)...
git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ==============================================================================
    echo    CHUC MUNG! DA DAY CODE LEN GITHUB THANH CONG!
    echo    Bay gio ban hay vao https://render.com va chon New Web Service de deploy nhe!
    echo ==============================================================================
) else (
    echo.
    echo [LUU Y] Neu GitHub yeu cau dang nhap, hay nhap Username va Token/Password cua ban tren cua so hien len nhe.
)

echo.
pause
