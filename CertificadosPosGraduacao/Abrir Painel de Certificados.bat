@echo off
cd /d "%~dp0"
if not exist "%~dp0Lib" mkdir "%~dp0Lib"

:: Copia UIA.ahk se nao existir
if not exist "%~dp0Lib\UIA.ahk" (
    if exist "%USERPROFILE%\Desktop\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA.ahk" (
        copy /Y "%USERPROFILE%\Desktop\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA.ahk" "%~dp0Lib\UIA.ahk" >nul
    )
)

:: Copia UIA_Browser.ahk se nao existir
if not exist "%~dp0Lib\UIA_Browser.ahk" (
    if exist "%USERPROFILE%\Desktop\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA_Browser.ahk" (
        copy /Y "%USERPROFILE%\Desktop\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA_Browser.ahk" "%~dp0Lib\UIA_Browser.ahk" >nul
    )
)

start "" "%~dp0PainelControle_Certificados.ahk"

