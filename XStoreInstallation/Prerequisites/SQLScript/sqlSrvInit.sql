CREATE LOGIN micros WITH PASSWORD = 'M1crosReta!l';
ALTER LOGIN micros WITH PASSWORD = 'M1crosReta!l' UNLOCK,CHECK_POLICY=OFF;
EXEC sp_addsrvrolemember 'micros', 'sysadmin';
EXEC sp_addsrvrolemember 'micros', 'public';
