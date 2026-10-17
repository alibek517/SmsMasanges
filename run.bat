@echo off
title SMS Xabarnoma va Boshqaruv Tizimi
chcp 65001 >nul
cls
echo =======================================================
echo    SMS Xabarnoma va Boshqaruv Tizimi
echo =======================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [XATO] Kompyuterda Python topilmadi!
    echo Iltimos, https://www.python.org saytidan Python-ni o'rnating.
    echo O'rnatish paytida "Add Python to PATH" katakchasiga belgi qo'ying!
    echo.
    pause
    exit /b
)

echo [1/2] Kutubxonalar tekshirilmoqda...
pip install -r requirements.txt --quiet --disable-pip-version-check

echo [2/2] Tizim ishga tushirilmoqda...
echo.
echo Brauzerda ochilmoqda: http://localhost:8501
echo Dasturni to'xtatish uchun ushbu oynani yoping.
echo.
streamlit run app.py
pause
