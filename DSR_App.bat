TITLE Automated DSR Run In Server
ECHO Please wait... Gathering Information

hostname | findstr /C:"PPBWDLM0MA00G" >nul
if %errorlevel% equ 0 (
    call C:\Users\I024605\AppData\Local\anaconda3\condabin\activate.bat MegaReport
) else (
    hostname | findstr /C:"PPBWDLC0SG7A1" >nul
    if %errorlevel% equ 0 (
        call C:\ProgramData\anaconda3\condabin\activate.bat MegaReport
    )
)

timeout 5
cd C:\Users\i024605\OneDrive - AIA Group Ltd\Documents\PythonCodes\Python_Pyfiles\DSR_Automate
timeout 5
python -m DSR_Auto
timeout 5

ECHO Finish Data Gathering...
timeout 5
exit