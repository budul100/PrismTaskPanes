@echo off

PUSHD "%~dp0..\..\"

echo.
echo Clean empty dirs, obj dirs and bin dirs in '%cd%'.
echo.

PAUSE

for /f "usebackq delims=" %%d in (`"dir /ad-h /b /s | sort /R"`) do rd "%%d"

FOR /F "tokens=*" %%G IN ('DIR /B /ad-h /S bin') DO RMDIR /S /Q "%%G"
FOR /F "tokens=*" %%G IN ('DIR /B /ad-h /S obj') DO RMDIR /S /Q "%%G"

POPD