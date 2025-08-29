@echo off
title ACTIVATOR CMD by @ariphx
color a
setlocal ENABLEDELAYEDEXPANSION

:: Check if the script is run as Administrator
whoami /groups | find "S-1-16-12288" >nul
if not !errorlevel! == 0 (
    echo This script requires Administrator privileges.
    echo Restarting with Administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)

:main_menu
cls
echo ============================================
echo          ACTIVATOR CMD by @ariphx
echo ============================================
echo.
echo Select a category:
echo.
echo [1] Microsoft
echo [2] Windows
echo [3] Enable/Disable Windows Script Host
echo [0] Exit
echo.
choice /c 1230 /n /m "Choose an option [1, 2, 3, or 0]: "

if !errorlevel! == 1 goto microsoft_menu
if !errorlevel! == 2 goto windows_menu
if !errorlevel! == 3 goto wsh_menu
if !errorlevel! == 4 goto exit

:microsoft_menu
cls
echo ============================================
echo        Microsoft Activation Menu
echo ============================================
echo.
echo Select the Microsoft program to activate:
echo.
echo [1] MICROSOFT 365
echo [2] MICROSOFT OFFICE 2010
echo [3] MICROSOFT OFFICE 2013
echo [4] MICROSOFT OFFICE 2016
echo [5] MICROSOFT OFFICE 2019
echo [6] MICROSOFT OFFICE 2021
echo [7] MICROSOFT OFFICE (All-in-one)
echo [0] Back to Main Menu
echo.
choice /c 12345670 /n /m "Choose an option [1-7 or 0]: "

if !errorlevel! == 1 goto ao365
if !errorlevel! == 2 goto o10
if !errorlevel! == 3 goto o13
if !errorlevel! == 4 goto o16
if !errorlevel! == 5 goto ao2019
if !errorlevel! == 6 goto ao2021
if !errorlevel! == 7 goto aso
if !errorlevel! == 8 goto main_menu

:windows_menu
cls
echo ============================================
echo         Windows Activation Menu
echo ============================================
echo.
echo Select the Windows version to activate:
echo.
echo [1] WINDOWS 11
echo [2] WINDOWS 10/11 ENTERPRISE LTSC
echo [3] WINDOWS 10 v2
echo [4] WINDOWS 10
echo [5] WINDOWS 8
echo [6] WINDOWS 7
echo [0] Back to Main Menu
echo.
choice /c 1234560 /n /m "Choose an option [1-6 or 0]: "

if !errorlevel! == 1 goto aw11
if !errorlevel! == 2 goto ltsc
if !errorlevel! == 3 goto w10v2
if !errorlevel! == 4 goto aw10
if !errorlevel! == 5 goto w8
if !errorlevel! == 6 goto w7
if !errorlevel! == 7 goto main_menu

:ao2019
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/ao2019.cmd -o ao2019.cmd && start /w ao2019.cmd
del ao2019.cmd
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto microsoft_menu

:ao365
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/ao365.cmd -o ao365.cmd && start /w ao365.cmd
del ao365.cmd
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto microsoft_menu

:ao2021
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/ao2021.cmd -o ao2021.cmd && start /w ao2021.cmd
del ao2021.cmd
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto microsoft_menu

:aso
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/aso.cmd -o aso.cmd && start /w aso.cmd
del aso.cmd
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto microsoft_menu

:o10
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/o10.bat -o o10.bat && start /w o10.bat
del o10.bat
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto microsoft_menu

:o13
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/o13.bat -o o13.bat && start /w o13.bat
del o13.bat
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto microsoft_menu

:o16
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/o16.bat -o o16.bat && start /w o16.bat
del o16.bat
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto microsoft_menu

:aw10
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/aw10.cmd -o aw10.cmd && start /w aw10.cmd
del aw10.cmd
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto windows_menu

:aw11
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/aw11.cmd -o aw11.cmd && start /w aw11.cmd
del aw11.cmd
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto windows_menu

:w10v2
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/w10.bat -o w10.bat && start /w w10.bat
del w10.bat
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto windows_menu

:w7
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/w7.bat -o w7.bat && start /w w7.bat
del w7.bat
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto windows_menu

:w8
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/w8.bat -o w8.bat && start /w w8.bat
del w8.bat
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto windows_menu

:ltsc
call :msgbox "Please wait!"
curl -L https://raw.githubusercontent.com/ariphx/activator-repo/main/ltsc.cmd -o ltsc.cmd && start /w ltsc.cmd
del ltsc.cmd
call :msgbox "Thank you for using ACTIVATOR CMD by @ariphx."
goto windows_menu

:wsh_menu
cls
echo ============================================
echo   Enable/Disable Windows Script Host
echo ============================================
echo.
echo Select an option:
echo.
echo [1] Enable Windows Script Host
echo [2] Disable Windows Script Host
echo [0] Back to Main Menu
echo.
choice /c 120 /n /m "Choose an option [1, 2, or 0]: "

if !errorlevel! == 1 goto enable_wsh
if !errorlevel! == 2 goto disable_wsh
if !errorlevel! == 3 goto main_menu

:enable_wsh
echo [PROSES] Enabling Windows Script Host...
REG ADD "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f >nul
call :msgbox "Windows Script Host has been enabled. Please restart your computer for the changes to take effect."
goto wsh_menu

:disable_wsh
echo [PROSES] Disabling Windows Script Host...
REG ADD "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 0 /f >nul
call :msgbox "Windows Script Host has been disabled. Please restart your computer for the changes to take effect."
goto wsh_menu

:exit
echo Thank you for using ACTIVATOR CMD by @ariphx.
pause
exit

:msgbox
setlocal
echo Set WshShell = CreateObject("WScript.Shell") > msgbox.vbs
echo WshShell.Popup "%~1", 5, "Activation", 64 >> msgbox.vbs
cscript //nologo msgbox.vbs
del msgbox.vbs
endlocal
goto :eof
