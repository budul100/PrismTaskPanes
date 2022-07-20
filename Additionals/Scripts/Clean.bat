ECHO off

powershell -command "Stop-Process -Name "excel" -Force"
powershell -command "Stop-Process -Name "powerpnt" -Force"

pushd ..\..

for /f "usebackq delims=" %%d in (`"dir /ad/b/s | sort /R"`) do rd "%%d"

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S _Out') DO RMDIR /S /Q "%%G"
FOR /F "tokens=*" %%G IN ('DIR /B /AD /S bin') DO RMDIR /S /Q "%%G"
FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"