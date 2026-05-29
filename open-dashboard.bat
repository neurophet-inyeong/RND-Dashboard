@echo off
chcp 65001 > nul
cd /d "%~dp0"

:: 포트 8080이 이미 사용 중인지 확인
netstat -an 2>nul | find "8080" | find "LISTENING" > nul
if %errorlevel% == 0 (
    echo [RnD Dashboard] 포트 8080이 이미 열려 있습니다. 브라우저를 엽니다.
    start "" "http://localhost:8080/dashboard.html"
    goto :end
)

:: 새 창에서 HTTP 서버 실행
start "RnD Dashboard Server" /min py -m http.server 8080
timeout /t 1 /nobreak > nul

:: 브라우저 열기
start "" "http://localhost:8080/dashboard.html"

echo.
echo [RnD Dashboard] 서버가 실행 중입니다.
echo 이 창(RnD Dashboard Server)을 닫으면 서버가 종료됩니다.

:end
