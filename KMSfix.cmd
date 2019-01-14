@ECHO OFF & CLS

REM Set script filename:
SET "_MyScript=KMS_Error_Fix.ps1"

REM OS Run architecture check and redirect if needed:
If "%PROCESSOR_ARCHITEW6432%"=="" (GOTO :_STANDARD) ELSE (GOTO :_SYSNATIVE)

:_SYSNATIVE
echo The Operating System is 64bit version
%WINDIR%\sysnative\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy bypass -file "%~dp0%_MyScript%"
goto :end

:_STANDARD
powershell.exe -NoProfile -ExecutionPolicy bypass -File "%~dp0%_MyScript%"
goto :end

:end
exit /b %errorlevel%
