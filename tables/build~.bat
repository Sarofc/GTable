@echo off

REM https://www.cnblogs.com/lixiaobin/p/MsbuildSetting.html

REM MSBuild.exe 路径
set MSBUILD_PATH=MSBuild.exe
REM solution 路径
set SOLUTION_FILE=..\tabtool\tabtool.csproj
REM Release / Debug
set BUILD_CONFIG=Release
REM build / rebuild / clean
set BUILD_TYPE=clean;build
REM AnyCPU(不能带空格) / x64 / x86
set BUILD_PLATFORM="AnyCPU"
set OUTPUT_PATH=..\tables\bin\

"%MSBUILD_PATH%" %SOLUTION_FILE% /t:%BUILD_TYPE% /p:Configuration=%BUILD_CONFIG%;Platform=%BUILD_PLATFORM%;OutDir=%OUTPUT_PATH%;AllowedReferenceRelatedFileExtensions=none

pause