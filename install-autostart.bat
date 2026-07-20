@echo off
chcp 65001 > nul

set "SERVER_VBS_SRC=%~dp0start-server.vbs"
set "LEDGER_VBS_SRC=%~dp0start-ledger-sync.vbs"
set "AUTOSYNC_VBS_SRC=%~dp0start-autosync.vbs"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "SERVER_VBS_DST=%STARTUP%\RnD-Dashboard-Server.vbs"
set "LEDGER_VBS_DST=%STARTUP%\RnD-Dashboard-LedgerSync.vbs"
set "AUTOSYNC_VBS_DST=%STARTUP%\RnD-Dashboard-AutoSync.vbs"

echo [1/2] Windows 시작 프로그램에 서버 + 원장 동기화 + Git 자동 동기화를 등록합니다...
copy /y "%SERVER_VBS_SRC%" "%SERVER_VBS_DST%" > nul
copy /y "%LEDGER_VBS_SRC%" "%LEDGER_VBS_DST%" > nul
copy /y "%AUTOSYNC_VBS_SRC%" "%AUTOSYNC_VBS_DST%" > nul

if exist "%SERVER_VBS_DST%" if exist "%LEDGER_VBS_DST%" if exist "%AUTOSYNC_VBS_DST%" (
    echo       완료: 다음 로그인부터 서버, 원장 동기화, Git 자동 동기화가 모두 자동으로 시작됩니다.
) else (
    echo       [오류] 등록에 실패했습니다. 폴더 권한을 확인해주세요.
    pause
    exit /b 1
)

echo [2/2] 지금 바로 서버, 원장 동기화, Git 자동 동기화를 시작합니다...
wscript.exe "%SERVER_VBS_SRC%"
wscript.exe "%LEDGER_VBS_SRC%"
wscript.exe "%AUTOSYNC_VBS_SRC%"
echo       완료.

echo.
echo ======================================================
echo  설치 완료!
echo  이제 dashboard.html 을 더블클릭하면 바로 열립니다.
echo ======================================================
echo.
pause
