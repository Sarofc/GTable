@echo off

echo copy tables folder to unity project

set TABLETOOL_NAME=tables
set TABLETOOL_PATH=..\%TABLETOOL_NAME%
set UNITY_PROJECT_PATH=..\..\DNF-Unity\%TABLETOOL_NAME%

REM echo %TABLETOOL_PATH%
REM echo %UNITY_PROJECT_PATH%
REM echo %UNITY_PROJECT_PATH%\%TABLETOOL_NAME%

xcopy %TABLETOOL_PATH% %UNITY_PROJECT_PATH% /s /exclude:uncopy~.txt

pause