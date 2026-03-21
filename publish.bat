@echo off
echo [1/3] remove old files ...
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj

echo [2/3] Publishing (Release, Single file, including Runtime)...
:: TRIM excluded Compresion enabled
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishReadyToRun=true -p:EnableCompressionInSingleFile=true

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [3/3] Sucessfully！
    echo Path: bin\Release\net8.0-windows\win-x64\publish\
    explorer "bin\Release\net8.0-windows\win-x64\publish\"
) else (
    echo.
    echo [!] Failed，please check above information。
    pause
)