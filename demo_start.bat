@echo off
chcp 65001 >nul
setlocal
rem ================================================================
rem  데모 게시판 켜기: 더블클릭하면 http://localhost:3456 으로 열린다.
rem  가짜·예시 데이터(data\fake, data\sample)만 쓴다. 실제 데이터는 run_merge.bat.
rem  이미 만든 결과(out\run_*)가 있으면 바로 띄우고, 없으면 통합부터 한다(몇 분).
rem  끄려면 이 창에서 Ctrl+C를 누르거나 창을 닫거나 demo_stop.bat을 더블클릭한다.
rem  통합을 새로 하려면 out 폴더를 지우고 다시 실행한다.
rem ================================================================
cd /d "%~dp0"

dir /b /ad "out\run_*" >nul 2>nul
if errorlevel 1 goto make_new

echo 이미 만든 결과로 게시판을 띄웁니다.
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\serve_board.ps1"
goto end

:make_new
echo 처음 실행이라 예시 데이터를 통합하고 게시판을 만듭니다. 몇 분 걸립니다.
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\run_demo.ps1" -Serve

:end
echo.
echo 게시판을 껐습니다.
pause
