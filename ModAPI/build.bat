@echo off
setlocal

set "VSDIR=C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Tools\MSVC\14.44.35207"
set "SDKDIR=C:\Program Files (x86)\Windows Kits\10"
set "SDKVER=10.0.26100.0"

set "PATH=%VSDIR%\bin\Hostx64\x64;%PATH%"
set "INCLUDE=%VSDIR%\include;%SDKDIR%\Include\%SDKVER%\ucrt;%SDKDIR%\Include\%SDKVER%\um;%SDKDIR%\Include\%SDKVER%\shared"
set "LIB=%VSDIR%\lib\x64;%SDKDIR%\Lib\%SDKVER%\ucrt\x64;%SDKDIR%\Lib\%SDKVER%\um\x64"

cd /d "C:\Users\zzies\Documents\Scala\ModAPI"

echo Compiling scala_cursor_mod.c...
cl -LD -O2 -W3 scala_cursor_mod.c -Fe:scala_cursor_mod.dll -link user32.lib

if exist scala_cursor_mod.dll (
    echo.
    echo Build successful!
    dir scala_cursor_mod.dll
) else (
    echo.
    echo Build failed!
)
