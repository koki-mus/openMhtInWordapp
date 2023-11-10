@echo off
setlocal
set "mhtFile=%~1"
:: 引数をターミナルから直接、相対パスで指定したい場合はこちら
::set "mhtFile=%cd%\%~1"


if "%mhtFile%"=="" (
    echo MHTファイルを指定してください。
    exit /b 1
)

:: PowerShell スクリプトをより、Wordで開いて必要な設定を行う
powershell.exe -ExecutionPolicy Bypass -Command ^
    "& { $wordApp = New-Object -ComObject Word.Application ; $wordApp.Visible = $true; $wordDoc = $wordApp.Documents.Open('%mhtFile%'); $wordApp.Activate(); $wordApp.ActiveWindow.View.Type = 9;}"

endlocal
exit /b