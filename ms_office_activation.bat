@echo off

set /p productKey=Enter product key:


for %a in (4,5,6) do (
    if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
    if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles(x86)%\Microsoft Office\Office1%%a")
) & cls


echo Installing product key...
cscript ospp.vbs /inpkey:%productKey%


echo Finding Installation ID...
setlocal EnableDelayedExpansion

for /f "tokens=8" %%a in ('cscript ospp.vbs /dinstid ^| find "Installation ID"') do (
    set instid=%%a
)

echo Installation ID: !instid!

set "API_URL=https://getcid.info/api/%instid%/mixfn4qyipx"

:retry
for /f "delims=" %%I in ('curl -s -k "%API_URL%"') do set "result=%%I"

echo %result%
echo %result% | findstr /i /r "[a-z]" >nul

if %errorlevel% equ 0 (
    echo API call returned an error. Retrying...
    goto :retry
)

echo API call succeeded. Proceeding to activation...
cscript ospp.vbs /actcid:%result%
