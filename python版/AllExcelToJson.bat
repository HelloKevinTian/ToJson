@echo off
set PATH=%CD%\Python27;%PATH%
del %CD%\Json\*.json
del %CD%\Json\Localization\*.json
python AllExcelToJson.py

copy %CD%\Json\*.json %CD%\..\..\code\gameServer\table

pause