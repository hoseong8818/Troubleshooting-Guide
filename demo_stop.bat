@echo off
chcp 65001 >nul
rem ================================================================
rem  ?곕え 寃뚯떆???꾧린: 3456 ?ы듃瑜??곕뒗 python(寃뚯떆???쒕쾭)留??덈떎.
rem ================================================================
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c = Get-NetTCPConnection -LocalPort 3456 -State Listen -ErrorAction SilentlyContinue; if (-not $c) { Write-Host '耳쒖졇 ?덈뒗 寃뚯떆?먯씠 ?놁뒿?덈떎.'; exit 0 }; foreach ($x in $c) { $p = Get-Process -Id $x.OwningProcess -ErrorAction SilentlyContinue; if ($p -and $p.ProcessName -like 'python*') { Stop-Process -Id $p.Id -Force; Write-Host '寃뚯떆?먯쓣 猿먯뒿?덈떎.' } else { Write-Host '3456 ?ы듃瑜?python???꾨땶 ?꾨줈洹몃옩???곌퀬 ?덉뼱 ?꾩? ?딆븯?듬땲??' } }"
echo.
pause
