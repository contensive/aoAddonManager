
rem all paths are relative to the git scripts folder

set appName=app200509
set collectionName=Add-on Manager
set collectionPath=..\collections\Add-on Manager\
set binPath=..\source\addonManager51\bin\debug\

rem copy bin folder assemblies to collection folder
copy "%binPath%*.dll" "%collectionPath%"

rem create new collection zip file
c:
cd %collectionPath%
del "%collectionName%.zip" /Q
"c:\program files\7-zip\7z.exe" a "%collectionName%.zip"
cd ..\..\scripts
