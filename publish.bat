@echo off
echo [1/3] 正在清理舊的發布檔案...
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj

echo [2/3] 正在發布專案 (Release, 單一檔案, 包含 Runtime)...
:: 這裡使用了不包含修剪 (Trim) 但開啟壓縮的設定
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishReadyToRun=true -p:EnableCompressionInSingleFile=true

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [3/3] 發布成功！
    echo 檔案路徑: bin\Release\net8.0-windows\win-x64\publish\
    explorer "bin\Release\net8.0-windows\win-x64\publish\"
) else (
    echo.
    echo [!] 發布失敗，請檢查上方錯誤訊息。
    pause
)