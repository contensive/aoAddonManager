@echo off
setlocal

@echo Build project and deploy to veronica

rem run the build
call "%~dp0build.cmd" /nopause
if not "%errorlevel%"=="0" (
    echo Build failed, skipping deploy.
    pause
    exit /b 1
)

rem upload to contensive application
cc -a veronica --installFile "%~dp0..\collections\Add-on Manager\Add-on Manager.zip"

pause
exit /b %errorlevel%
