@echo off
chcp 65001 >nul
rem ================================================================
rem  데모 게시판 끄기: 3456 포트를 쓰는 python(게시판 서버)만 끈다.
rem ================================================================
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c = Get-NetTCPConnection -LocalPort 3456 -State Listen -ErrorAction SilentlyContinue; if (-not $c) { Write-Host '켜져 있는 게시판이 없습니다.'; exit 0 }; foreach ($x in $c) { $p = Get-Process -Id $x.OwningProcess -ErrorAction SilentlyContinue; if ($p -and $p.ProcessName -like 'python*') { Stop-Process -Id $p.Id -Force; Write-Host '게시판을 껐습니다.' } else { Write-Host '3456 포트를 python이 아닌 프로그램이 쓰고 있어 끄지 않았습니다.' } }"
echo.
pause
