@echo off
chcp 65001 > nul

set "VBS_SRC=%~dp0start-server.vbs"
set "STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_DST=%STARTUP%\RnD-Dashboard-Server.vbs"

echo [1/2] Windows 시작 프로그램에 서버를 등록합니다...
copy /y "%VBS_SRC%" "%VBS_DST%" > nul

if exist "%VBS_DST%" (
    echo       완료: 다음 로그인부터 서버가 자동으로 시작됩니다.
) else (
    echo       [오류] 등록에 실패했습니다. 폴더 권한을 확인해주세요.
    pause
    exit /b 1
)

echo [2/2] 지금 바로 서버를 시작합니다...
wscript.exe "%VBS_SRC%"
echo       완료.

echo.
echo ======================================================
echo  설치 완료!
echo  이제 dashboard.html 을 더블클릭하면 바로 열립니다.
echo ======================================================
echo.
pause
