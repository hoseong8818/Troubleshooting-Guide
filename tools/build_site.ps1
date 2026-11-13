# 배포용 정적 사이트 만들기: 예시 데이터 게시판을 public\index.html로 복사한다(Vercel 배포용).
#
# 사용(저장소 루트 어디서든):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\build_site.ps1
#
# - tools\run_demo.ps1 -NoOpen을 실행해 새 결과 폴더 out\run_<시각>을 만들고(브라우저·서버는 띄우지 않음),
#   이번 실행에서 만들어진 가장 최근 out\run_*\sample\board.html을 public\index.html로 복사한다.
# - 입력은 가상 예시 데이터(data\sample)뿐이다. -Sample 같은 입력 위치 옵션이 없고, 실제 데이터 폴더
#   (real_data, real_out 등)를 읽거나 복사하는 경로가 없다. 실제 데이터로 만든 board.html은 public에 넣지 않는다.
# - 복사 전에 외부 요청 흔적(https?://, fetch(, XMLHttpRequest, WebSocket, <script src, <link)을 검사하고,
#   하나라도 있으면 복사하지 않고 실패한다(어느 검사에 걸렸는지만 출력하고 내용은 출력하지 않는다).
# - 게시판을 고친 뒤에는 이 스크립트를 다시 실행하고 public\index.html을 커밋한다.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
$started = Get-Date

# 1. 통합과 게시판 생성(예시 데이터만, 브라우저·서버 없음)
# 일부 확인 필요가 있으면 python이 종료 코드 1을 내므로 종료 코드는 보지 않고, 아래에서 결과 파일로 판단한다.
& (Join-Path $PSScriptRoot "run_demo.ps1") -NoOpen
$ErrorActionPreference = "Stop"

# 2. 이번 실행에서 만들어진 가장 최근 예시 데이터 게시판 찾기
$board = Get-ChildItem (Join-Path $root "out") -Directory -Filter "run_*" -ErrorAction SilentlyContinue |
  Sort-Object Name -Descending |
  ForEach-Object { Join-Path $_.FullName "sample\board.html" } |
  Where-Object { Test-Path -LiteralPath $_ } |
  Select-Object -First 1
if (-not $board) { throw "예시 데이터 게시판(out\run_*\sample\board.html)을 찾지 못했습니다." }
if ((Get-Item -LiteralPath $board).LastWriteTime -lt $started.AddSeconds(-1)) {
  throw "이번 실행에서 만든 게시판이 아닙니다. 통합이나 게시판 생성이 실패했는지 확인하세요: $board"
}
Write-Host ("복사할 게시판: " + $board)

# 3. 복사 전 검사: 외부 요청 흔적이 하나라도 있으면 실패
$html = [IO.File]::ReadAllText($board, [Text.Encoding]::UTF8)
$checks = [ordered]@{
  "https?://"        = "https?://"
  "fetch("           = "fetch\("
  "XMLHttpRequest"   = "XMLHttpRequest"
  "WebSocket"        = "WebSocket"
  "<script src"      = "<script\s[^>]*\bsrc\s*="
  "<link"            = "<link[\s>/]"
}
$bad = @()
foreach ($name in $checks.Keys) {
  if ($html -match ("(?i)" + $checks[$name])) { $bad += $name }
}
if ($bad.Count -gt 0) {
  throw ("배포용 검사 실패: 외부 요청 흔적이 있어 복사하지 않습니다. 걸린 검사: " + ($bad -join ", "))
}
Write-Host "배포용 검사 통과: 외부 주소, 외부 스크립트, 외부 요청 코드가 없습니다."

# 4. public\index.html로 복사(같은 파일을 새 결과로 바꾼다)
$pub = Join-Path $root "public"
New-Item -ItemType Directory -Force $pub | Out-Null
Copy-Item -LiteralPath $board -Destination (Join-Path $pub "index.html") -Force
$size = (Get-Item (Join-Path $pub "index.html")).Length
Write-Host ("완료: public\index.html ({0:N0} 바이트). 확인 후 git add public\index.html 하고 커밋하세요." -f $size)
