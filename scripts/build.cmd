
rem @echo off

rem Must be run from the projects git\project\scripts folder - everything is relative
rem run >build [versionNumber]
rem versionNumber is YY.MM.DD.build-number, like 20.5.8.1
rem

c:
cd \Git\aoAddonManager\scripts

set appName=app200509
set collectionName=Add-on Manager
set collectionPath=..\collections\Add-on Manager\
set binPath=..\source\addonManager51\bin\debug\
set deploymentFolderRoot=C:\Deployments\aoAddonManager\Dev\
set msbuildLocation=C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\
set NuGetLocalPackagesFolder=C:\NuGetLocalPackages\

rem @echo off
rem Setup deployment folder

set year=%date:~12,4%
set month=%date:~4,2%
if %month% GEQ 10 goto monthOk
set month=%date:~5,1%
:monthOk
set day=%date:~7,2%
if %day% GEQ 10 goto dayOk
set day=%date:~8,1%
:dayOk
set versionMajor=%year%
set versionMinor=%month%
set versionBuild=%day%
set versionRevision=1
rem
rem if deployment folder exists, delete it and make directory
rem
:tryagain
set versionNumber=%versionMajor%.%versionMinor%.%versionBuild%.%versionRevision%
if not exist "%deploymentFolderRoot%%versionNumber%" goto :makefolder
set /a versionRevision=%versionRevision%+1
goto tryagain
:makefolder
md "%deploymentFolderRoot%%versionNumber%"

rem ==============================================================
rem
rem clean build folders
rem
rem rd /S /Q "..\source\addonManager51\bin"
rem rd /S /Q "..\source\addonManager51\obj"
rem 
rem rd /S /Q "..\source\addonManager51\bin"
rem rd /S /Q "..\source\addonManager51\obj"
rem 
rem ==============================================================
rem
echo build addonManager 
rem
cd ..\source

dotnet clean %solutionName%

dotnet build addonManager51/addonManager51.vbproj --no-restore --configuration Debug --no-dependencies /property:Version=%versionNumber% /property:AssemblyVersion=%versionNumber% /property:FileVersion=%versionNumber%
if errorlevel 1 (
   echo failure building addonManager51
   pause
   exit /b %errorlevel%
)

cd ..\scripts

rem ==============================================================
rem
echo Build addon collection
rem

rem remove old DLL files from the collection folder
del "%collectionPath%"\*.DLL
del "%collectionPath%"\*.dll.config

rem copy bin folder assemblies to collection folder
copy "%binPath%*.dll" "%collectionPath%"
copy "%binPath%*.dep" "%collectionPath%"
copy "%binPath%*.dll.config" "%collectionPath%"

rem create new collection zip file
c:
cd %collectionPath%
del "%collectionName%.zip" /Q
"c:\program files\7-zip\7z.exe" a "%collectionName%.zip"
xcopy "%collectionName%.zip" "%deploymentFolderRoot%%versionNumber%" /Y
cd ..\..\scripts

