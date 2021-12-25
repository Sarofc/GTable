@echo off

echo copy tables folder to unity project

set UNITY_PROJECT_PATH=..\..\souls

set TABLETOOL_NAME=tables
set TABLETOOL_PATH=..\%TABLETOOL_NAME%
set UNITY_PROJECT_TABLE_PATH=%UNITY_PROJECT_PATH%\%TABLETOOL_NAME%

set TABLETOOL_SCRIPT_PATH=..\tabtool\src\loader
set UNITY_PROJECT_SCRIPT_PATH=%UNITY_PROJECT_PATH%\Assets\Scripts\Generate\TableLoader

REM echo %TABLETOOL_PATH%
REM echo %UNITY_PROJECT_TABLE_PATH%
REM echo %UNITY_PROJECT_TABLE_PATH%\%TABLETOOL_NAME%

xcopy %TABLETOOL_PATH% %UNITY_PROJECT_TABLE_PATH% /s /exclude:uncopy~.txt

REM echo copy %TABLETOOL_SCRIPT_PATH% to %UNITY_PROJECT_SCRIPT_PATH%

xcopy %TABLETOOL_SCRIPT_PATH% %UNITY_PROJECT_SCRIPT_PATH% /s

pause