# tasks.md에 Phase 11(배포 준비)을 추가하고 커밋한다. 사람이 직접 실행한다.
# 사용(저장소 루트에서):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\add_phase11.ps1
# - '## 의존 관계' 바로 앞에 Phase 11 절을 넣는다. 이미 있으면 아무것도 하지 않는다.
# - BOM 없이 CRLF로 쓴다(기존 tasks.md와 같은 형식).
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
$path = Join-Path $root "specs\001-trouble-merge\tasks.md"
$enc = New-Object System.Text.UTF8Encoding $false
$text = [IO.File]::ReadAllText($path, $enc)

if ($text.Contains("## Phase 11:")) {
  Write-Host "이미 Phase 11이 있어 건너뜁니다."
  exit 0
}
$marker = "## 의존 관계"
$idx = $text.IndexOf($marker)
if ($idx -lt 0) { throw "'## 의존 관계' 줄을 찾지 못했습니다." }

$bt = [string][char]96
$lines = @(
  "## Phase 11: 배포 준비 (P3, 데모) GitHub + Vercel 정적 배포",
  "",
  "- [x] T069 [US5] ${bt}vercel.json${bt}: 정적 사이트(framework·buildCommand·installCommand는 null, outputDirectory는 public, cleanUrls), 모든 경로에 CSP(외부 연결·외부 스크립트·폼 전송 차단)·nosniff·no-referrer·noindex·Cache-Control 헤더. 게시판이 이 CSP에서 깨지지 않는지 확인",
  "- [x] T070 [US5] ${bt}tools/run_demo.ps1${bt}에 ${bt}-NoOpen${bt}(브라우저도 서버도 띄우지 않음), ${bt}tools/build_site.ps1${bt}: 예시 데이터(data\sample)로 통합한 가장 최근 ${bt}out\run_*\sample\board.html${bt}을 ${bt}public\index.html${bt}로 복사. 복사 전에 외부 주소·${bt}fetch(${bt}·XMLHttpRequest·WebSocket·${bt}<script src${bt}·${bt}<link${bt}가 있으면 실패. 실제 데이터 경로 없음(헌법 원칙 1)",
  "- [x] T071 [US5] ${bt}public/index.html${bt} 생성·커밋, README에 'GitHub + Vercel 배포' 절 추가",
  "- [ ] T072 [사람] 비공개 GitHub 저장소 만들기, ${bt}git remote add${bt}와 ${bt}git push${bt}, Vercel에서 Import Project(Framework Preset Other, Build Command 비움, Output Directory public), 시연 링크를 비공개로 하려면 Deployment Protection 켜기",
  "- [ ] T073 [사람] 배포 뒤 확인: 검색·글 보기·오류 현황 동작, 응답 헤더에 CSP와 ${bt}X-Robots-Tag${bt}가 있는지, 실제 데이터로 만든 board.html이 저장소·public에 없는지",
  "",
  ""
)
$block = ($lines -join "`r`n")
$text = $text.Substring(0, $idx) + $block + $text.Substring($idx)
[IO.File]::WriteAllText($path, $text, $enc)
Write-Host "tasks.md에 Phase 11을 추가했습니다."

git add specs/001-trouble-merge/tasks.md
git commit -m "문서: tasks.md에 Phase 11(배포 준비) 추가" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
git status --short
