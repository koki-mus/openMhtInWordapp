@echo off
setlocal
set "mhtFile=%~1"
:: �������^�[�~�i�����璼�ځA���΃p�X�Ŏw�肵�����ꍇ�͂�����
::set "mhtFile=%cd%\%~1"


if "%mhtFile%"=="" (
    echo MHT�t�@�C�����w�肵�Ă��������B
    exit /b 1
)

:: PowerShell �X�N���v�g�����AWord�ŊJ���ĕK�v�Ȑݒ���s��
powershell.exe -ExecutionPolicy Bypass -Command ^
    "& { $wordApp = New-Object -ComObject Word.Application ; $wordApp.Visible = $true; $wordDoc = $wordApp.Documents.Open('%mhtFile%'); $wordApp.Activate(); $wordApp.ActiveWindow.View.Type = 9;}"

endlocal
exit /b