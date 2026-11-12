# NFR-007 상한 구현 뒤 낡아진 문구를 고친다 (CLAUDE.md, spec.md, tasks.md 체크박스).
# 사용: powershell -ExecutionPolicy Bypass -File C:\mywork\tools\apply_doc_fixes.ps1
# 이미 고쳐져 있으면 그 부분은 건너뛴다.
$enc = New-Object System.Text.UTF8Encoding($false)
$dot = [string][char]0x00B7

# --- CLAUDE.md ---
$p = "C:\mywork\CLAUDE.md"
$t = [System.IO.File]::ReadAllText($p, $enc)
$old = "데모 범위에서는 입력 크기" + $dot + "행 수 상한(NFR-007), 운영체제 권한, 회사 PC 확인을 하지 않았고 실제 업무 이관 시 T050에서 한다(README 참고)."
$new = "입력 크기와 행 수 상한(NFR-007)은 명령줄 옵션으로 구현했고(기본값은 데모용), 운영체제 권한과 회사 PC 확인은 데모 범위에서 하지 않았으며 실제 업무 이관 시 T050에서 한다(README 참고)."
if ($t.Contains($old)) {
    $t = $t.Replace($old, $new).Replace("테스트 74개가 통과한다", "테스트 81개가 통과한다")
    [System.IO.File]::WriteAllText($p, $t, $enc)
    "CLAUDE.md 고침"
} else { "CLAUDE.md: 고칠 문구가 없음(이미 고쳤거나 문구가 다름)" }

# --- spec.md ---
$p = "C:\mywork\specs\001-trouble-merge\spec.md"
$t = [System.IO.File]::ReadAllText($p, $enc)
$old = "현재 구현은 ``.xlsm``과 ``.xls`` 거부까지만 하며 크기와 행 수 상한은 데모 범위에서 구현하지 않았다. 실제 업무 이관 전에 숫자를 정해 구현한다."
$old = $old.Replace("``", "``")
$bt = [string][char]96
$old = "현재 구현은 " + $bt + ".xlsm" + $bt + "과 " + $bt + ".xls" + $bt + " 거부까지만 하며 크기와 행 수 상한은 데모 범위에서 구현하지 않았다. 실제 업무 이관 전에 숫자를 정해 구현한다."
$new = "크기와 행 수 상한은 명령줄 옵션 " + $bt + "--max-file-mb" + $bt + ", " + $bt + "--max-unzipped-mb" + $bt + ", " + $bt + "--max-rows" + $bt + "로 받고 기본값은 50, 200, 100000(데모용)이다. 실제 업무 이관 전에 정보보호 부서 기준으로 숫자를 정한다. 행 수는 파일을 읽은 뒤에 검사한다."
$fr24old = "실패한 파일(없으면 " + $bt + "없음" + $bt + "), "
$fr24new = $fr24old + "상한 초과 파일(파일 표시명과 사유 종류), "
$changed = $false
if ($t.Contains($old)) { $t = $t.Replace($old, $new); $changed = $true }
if ($t.Contains($fr24old) -and -not $t.Contains("상한 초과 파일(")) { $t = $t.Replace($fr24old, $fr24new); $changed = $true }
if ($changed) { [System.IO.File]::WriteAllText($p, $t, $enc); "spec.md 고침" } else { "spec.md: 고칠 문구가 없음" }

# --- tasks.md 체크박스 ---
$p = "C:\mywork\specs\001-trouble-merge\tasks.md"
$lines = [System.IO.File]::ReadAllLines($p, $enc)
$n = 0
$out = foreach ($l in $lines) {
    $m = $l
    foreach ($id in "T017e", "T017f") {
        $pre = "- [ ] $id "
        if ($l.StartsWith($pre)) { $m = "- [x] $id " + $l.Substring($pre.Length); $n++ }
    }
    $m
}
[System.IO.File]::WriteAllLines($p, $out, $enc)
"tasks.md 체크한 줄 수: $n"

git -C C:\mywork diff --stat
