@echo off
echo Criando executavel...

pyinstaller --onefile --windowed --icon=BKP.ico --add-data "BKP.ico:." --distpath=build BKP.py
echo.

echo Executavel criado em: build\
pause
