@ECHO OFF & CLS

REM Set script filename:
SET "_MyScript=WMF51-FullInstall.ps1"

REM OS Run architecture check and redirect if needed:
If "%PROCESSOR_ARCHITEW6432%"=="" (GOTO :_STANDARD) ELSE (GOTO :_SYSNATIVE)

:_SYSNATIVE
echo The Operating System is 64bit version
net stop PSEXESVC
sc delete PSEXESVC
del c:\windows\system32\psexesvc.exe /F /Q
del C:\Windows\SysWOW64\psexesvc.exe /F /Q
%WINDIR%\sysnative\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy bypass -file "%~dp0%_MyScript%"
IF /I "%ERRORLEVEL%" EQU "0" (
goto :end
) ELSE (
exit /b %errorlevel%
)

:_STANDARD
echo The Operating System is 32bit version
net stop PSEXESVC
sc delete PSEXESVC
del c:\windows\system32\psexesvc.exe /F /Q
powershell.exe -NoProfile -ExecutionPolicy bypass -File "%~dp0%_MyScript%"
IF /I "%ERRORLEVEL%" EQU "0" (
goto :end
) ELSE (
exit /b %errorlevel%
)

:end
%~dp0PSExec.exe -s -accepteula %~dp0ServiceUI.exe %~dp0ShutdownTool.exe /t:600 /m:0 /r /f /c
IF /I "%errorlevel%" EQU "0" (
exit /b 1641
) ELSE (
exit /b %errorlevel%
)
