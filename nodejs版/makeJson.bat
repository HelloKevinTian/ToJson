SET NODE_EXE=.\nodejs\node

SET OUTPUT_FILE_PATH=json

del config.json
rd /s /q %OUTPUT_FILE_PATH%

%NODE_EXE% excelTools.js ./table ./json

pause