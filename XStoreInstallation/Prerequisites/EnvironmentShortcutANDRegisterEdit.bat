echo Environment shortcut and register entry has been added! DOS STORE
@echo off

set SCRIPT="%TEMP%\%RANDOM%-%RANDOM%-%RANDOM%-%RANDOM%.vbs"

echo Set oWS = WScript.CreateObject("WScript.Shell") >> %SCRIPT%
echo sLinkFile = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\environment.lnk" >> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> %SCRIPT%
echo oLink.TargetPath = "C:\environment\environment.bat" >> %SCRIPT%
echo oLink.Save >> %SCRIPT% 
cscript /nologo %SCRIPT%
del %SCRIPT%
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v Shell /t REG_SZ /d "c:\environment\environment.bat" /f
PAUSE 
