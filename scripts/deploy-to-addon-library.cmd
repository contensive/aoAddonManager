@echo off
setlocal

rem Uploads the built collection to the Addon Collection Library on contensive.com.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy-to-addon-library.ps1" ^
    -CollectionName "Add-on Manager" ^
    -CollectionPath "%~dp0..\collections\Add-on Manager" ^
    -DeploymentPath "C:\Deployments\aoAddonManager" ^
    -UiPath "%~dp0..\ui"

set exitCode=%errorlevel%
pause
exit /b %exitCode%
