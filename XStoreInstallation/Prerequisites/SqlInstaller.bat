echo Copying Installer
xcopy /s C:\POS-Software-Sources\DTVINST\SQLEXPRWT_x64_ENU.exe "C:\Z_pulpitu\XStoreInstallation\Installation\XStoreInstallation\Prerequisites\SQLScript\" /y
cd C:\Z_pulpitu\XStoreInstallation\Installation\XStoreInstallation\Prerequisites\SQLScript
START /W SQLEXPRWT_x64_ENU.exe /x
START /W SETUP.EXE /ConfigurationFile=ConfigurationFile.ini
net stop MSSQLSERVER
net start MSSQLSERVER






