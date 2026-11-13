@echo off
chcp 65001 >nul
setlocal
rem ================================================================
rem  ?곕え 寃뚯떆??耳쒓린: ?붾툝?대┃?섎㈃ http://localhost:3456 ?쇰줈 ?대┛??
rem  媛吏쑣룹삁???곗씠??data\fake, data\sample)留??대떎. ?ㅼ젣 ?곗씠?곕뒗 run_merge.bat.
rem  ?대? 留뚮뱺 寃곌낵(out\run_*)媛 ?덉쑝硫?諛붾줈 ?꾩슦怨? ?놁쑝硫??듯빀遺???쒕떎(紐?遺?.
rem  ?꾨젮硫???李쎌뿉??Ctrl+C瑜??꾨Ⅴ嫄곕굹 李쎌쓣 ?リ굅??demo_stop.bat???붾툝?대┃?쒕떎.
rem  ?듯빀???덈줈 ?섎젮硫?out ?대뜑瑜?吏?곌퀬 ?ㅼ떆 ?ㅽ뻾?쒕떎.
rem ================================================================
cd /d "%~dp0"

dir /b /ad "out\run_*" >nul 2>nul
if errorlevel 1 goto make_new

echo ?대? 留뚮뱺 寃곌낵濡?寃뚯떆?먯쓣 ?꾩썎?덈떎.
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\serve_board.ps1"
goto end

:make_new
echo 泥섏쓬 ?ㅽ뻾?대씪 ?덉떆 ?곗씠?곕? ?듯빀?섍퀬 寃뚯떆?먯쓣 留뚮벊?덈떎. 紐?遺?嫄몃┰?덈떎.
powershell -NoProfile -ExecutionPolicy Bypass -File "tools\run_demo.ps1" -Serve

:end
echo.
echo 寃뚯떆?먯쓣 猿먯뒿?덈떎.
pause
