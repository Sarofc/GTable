@echo off

REM https://www.cnblogs.com/lixiaobin/p/MsbuildSetting.html

REM MSBuild.exe ·��
set MSBUILD_PATH=MSBuild.exe
REM solution ·��
set SOLUTION_FILE=..\GTable\GTable.csproj
REM Release / Debug
set BUILD_CONFIG=Release
REM build / rebuild / clean
set BUILD_TYPE=clean;build
REM AnyCPU(���ܴ��ո�) / x64 / x86
set BUILD_PLATFORM="AnyCPU"
set OUTPUT_PATH=..\tables\bin\

"%MSBUILD_PATH%" %SOLUTION_FILE% /t:%BUILD_TYPE% /p:Configuration=%BUILD_CONFIG%;Platform=%BUILD_PLATFORM%;OutDir=%OUTPUT_PATH%;AllowedReferenceRelatedFileExtensions=none

pause