# 게시판을 http://localhost:3456 으로 띄운다(이 PC에서만 접속, 외부 공개 안 함).
#
# 사용(저장소 루트 어디서든):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\serve_board.ps1            # 가장 최근 데모 결과
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\serve_board.ps1 -Dir <폴더> # board.html이 있는 폴더 지정
#
# 끝낼 때는 이 창에서 Ctrl+C. 표준 라이브러리 http.server만 쓴다(외부 패키지 없음).
param(
  [string]$Dir = "",
  [int]$Port = 3456
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not $Dir) {
  $latest = Get-ChildItem (Join-Path $root "out") -Directory -Filter "run_*" -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending | Select-Object -First 1
  if (-not $latest) { throw "out\run_* 폴더가 없습니다. 먼저 tools\run_demo.ps1을 실행하세요." }
  $Dir = Join-Path $latest.FullName "sample"
}
if (-not (Test-Path (Join-Path $Dir "board.html"))) { throw "board.html이 없습니다: $Dir" }

# 첫 화면이 바로 게시판이 되도록 같은 내용의 index.html을 둔다(board.html은 그대로)
Copy-Item (Join-Path $Dir "board.html") (Join-Path $Dir "index.html") -Force

$url = "http://localhost:$Port/"
Write-Host "게시판: $url"
Write-Host "폴더: $Dir"
Write-Host "끝내려면 Ctrl+C"
Start-Process $url
python -m http.server $Port --bind 127.0.0.1 --directory $Dir
