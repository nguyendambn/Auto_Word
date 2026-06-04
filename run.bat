@echo off
chcp 65001 >nul 2>nul
title Tool Word - Desktop Offline By Nguyen Van Dam

color 0B
echo.
echo  ================================================================
echo  #                                                              #
echo  #       A U T O  -  W O R D    F O R M A T T E R              #
echo  #                                                              #
echo  ================================================================
echo  #                                                              #
echo  #   Tool Word - Desktop Offline              By Nguyen Van Dam #
echo  #   Dai hoc Cong nghiep Ha Noi (HaUI)                         #
echo  #                                                              #
echo  #   Donate: 2562207069 - BIDV (Nguyen Van Dam)                 #
echo  #                                                              #
echo  ================================================================
echo.

:: Kiem tra Python
where python >nul 2>nul
if errorlevel 1 (
    echo  [LOI] Khong tim thay Python tren he thong!
    echo  Vui long cai dat Python 3.8+ va them vao PATH.
    echo.
    pause
    exit /b 1
)

:: Tao moi truong ao neu chua co
if not exist ".venv" (
    echo  [INFO] Dang tao moi truong ao Python...
    python -m venv .venv
    if errorlevel 1 (
        echo  [LOI] Khong the tao moi truong ao!
        pause
        exit /b 1
    )
    echo  [OK] Da tao moi truong ao thanh cong.
)

:: Kich hoat moi truong ao
call .venv\Scripts\activate.bat

:: Cai dat thu vien
echo  [INFO] Dang kiem tra thu vien...
pip install -q -r requirements.txt >nul 2>nul
if errorlevel 1 (
    echo  [INFO] Dang cai dat thu vien - co the mat 1-2 phut lan dau...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo  [LOI] Khong the cai dat thu vien!
        pause
        exit /b 1
    )
)

:: Chay ung dung
echo.
echo  [OK] Dang khoi chay ung dung...
echo  (Cua so ung dung se tu dong mo. Dong cua so nay khi muon tat.)
echo.
python app.py

if errorlevel 1 (
    echo.
    echo  [LOI] Co loi xay ra khi chay ung dung.
    pause
)
