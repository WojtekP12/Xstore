echo environment shortcut has been added. SIS STORE
@echo off

set SCRIPT="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"

echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT%
echo sLinkFile = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\environment.lnk" >> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.TargetPath = "C:\environment\environment.bat" >> %SCRIPT%
echo oLink.Save >> %SCRIPT% 
cscript /nologo %SCRIPT%
del %SCRIPT%
PAUSE 
