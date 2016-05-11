
1. Run and extract to C:\sqlserver2012_install
SQLEXPRWT_x64_ENU /x 
2. copy ConfigurationFile.ini to C:\sqlserver2012_install
3. cd C:\sqlserver2012_install
4. run 
SETUP.EXE /ConfigurationFile=ConfigurationFile.ini
5. restart windows service
net stop MSSQLSERVER
net start MSSQLSERVER
