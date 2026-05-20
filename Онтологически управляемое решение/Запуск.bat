@echo off
chcp 65001 > nul

cd /d "%~dp0"

python "Онтологически управляемое решение Sv16.py"

pause