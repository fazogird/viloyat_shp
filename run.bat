@echo off
rem ===================================================================
rem  Conda muhitida skript ishga tushirish (conda activate shart emas).
rem    .\run.bat atribut_otkazish.py
rem    .\run.bat qoshnidan_toldirish.py
rem  Avval viloyat_shp muhiti, bo'lmasa fenologiya muhiti ishlatiladi.
rem  Muhit yaratish: conda env create -f environment.yml
rem ===================================================================
set "ENV=%LOCALAPPDATA%\miniconda3\envs\viloyat_shp"
if not exist "%ENV%\python.exe" set "ENV=%LOCALAPPDATA%\miniconda3\envs\fenologiya"
if not exist "%ENV%\python.exe" (
  echo Conda muhiti topilmadi: viloyat_shp yoki fenologiya
  echo Yaratish: conda env create -f environment.yml
  exit /b 1
)
"%ENV%\python.exe" %*
