@echo off
chcp 65001 > nul
title Muratpaşa İlçe MEM - İSG Yönetim Portalı
cls
echo ======================================================================
echo    MURATPAŞA İLÇE MİLLİ EĞİTİM MÜDÜRLÜĞÜ - İSG YÖNETİM PORTALI
echo ======================================================================
echo.
echo  Portal Başlatılıyor, lütfen bekleyin...
echo.
echo  Bu Bilgisayardan Erişim  : http://localhost:8000
echo  Aynı Ağdaki Diğer Cihazlar: http://192.168.1.102:8000
echo.
echo  Koordinatör Girişi       : ramazan.eryildiz  /  isg2026!
echo.
echo  [Bu siyah pencereyi açık tutunuz. Kapatırsanız portal durur.]
echo ======================================================================
echo.

start "" http://localhost:8000

"C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe" -m uvicorn app:app --host 0.0.0.0 --port 8000 --reload
pause

