echo Only the register entry has been added.
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v Shell /t REG_SZ /d "c:\environment\environment.bat" /f
PAUSE 
