# 시연용 설정 폴더를 만든다: config 의 사본에 샘플 엑셀 80개에서 나온 설비 코드와 분류를 변환표로 더한다.
# 저장소의 config 는 건드리지 않는다(기존 테스트와 정답 파일이 그 값에 의존한다).
# 사용: powershell -ExecutionPolicy Bypass -File C:\mywork\tools\make_demo_config.ps1 [-WithVariants]
#   -WithVariants: 하이픈이 빠진 설비 코드 5개(RFL003, RFL004, PLT017, GRD009, FLT009)를 하이픈 있는 코드로 연결하는 줄도 더한다.
#                  사람이 각 값이 맞는지 확인한 뒤에만 쓴다.
#   -Force: C:\temp\demo_config 가 이미 있으면 지우고 다시 만든다(이 폴더만 지운다).
param([switch]$WithVariants, [switch]$Force)

$source = "C:\mywork\config"
$dest = "C:\temp\demo_config"
if (Test-Path $dest) {
    if ($Force) { Remove-Item -Recurse -Force $dest } else { throw "이미 있음: $dest (-Force 로 다시 만들 수 있습니다)" }
}

# 줄 계산이 맞는지 확인하기 전에 설정 폴더를 만들지 않도록, 변환표를 임시 파일에서 먼저 만든다
$path = "$dest\conversion_map.csv"
$stage = "$env:TEMP\demo_conversion_stage.csv"
Copy-Item "$source\conversion_map.csv" $stage -Force
$origPath = $path
$path = $stage
$enc = New-Object System.Text.UTF8Encoding($false)
$text = [System.IO.File]::ReadAllText($path, $enc)
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
if (-not $text.EndsWith("`n")) { $text += $nl }

# 이미 있는 키(종류,원래 값을 소문자로)는 건너뛴다
$existing = @{}
foreach ($line in ($text -split "`r?`n")) {
    $f = $line.TrimStart([char]0xFEFF).Split(",")
    if ($f.Length -ge 2) { $existing[($f[0].Trim() + "," + $f[1].Trim().ToLower())] = $true }
}

$new = New-Object System.Collections.Generic.List[string]
function Add-Line($kind, $original, $standard) {
    $key = $kind + "," + $original.ToLower()
    if (-not $existing.ContainsKey($key)) {
        $existing[$key] = $true
        $new.Add("$kind,$original,$standard")
    }
}

# 설비 코드: 샘플에서 나온 접두사별 번호 범위(HTF-007 은 샘플에 없음)
# 주의: PowerShell 에서 쉼표가 범위 연산자(..)보다 먼저 계산되므로 범위는 반드시 괄호로 묶는다
$ranges = [ordered]@{
    "CLT" = (1..10); "FLT" = (1..10); "GRD" = (1..10); "HTF" = ((1..6) + (8..20))
    "INJ" = (1..10); "PLT" = (1..20); "RFL" = (1..20); "VIS" = (1..10)
}
foreach ($prefix in $ranges.Keys) {
    foreach ($n in $ranges[$prefix]) {
        $code = "{0}-{1:D3}" -f $prefix, [int]$n
        Add-Line "설비" $code $code
    }
}
$equipmentAdded = ($new | Where-Object { $_.StartsWith("설비,") }).Count
if ($equipmentAdded -ne 109) { throw "설비 코드 줄이 109개가 아님: $equipmentAdded (스크립트 오류, 파일을 만들지 않았습니다)" }

# 분류: 샘플에서 나온 값 중 분류로 보이는 9개. 사람 이름처럼 보이는 값과 기타는 일부러 뺐다.
foreach ($c in "구동", "기구", "유공압", "소모품", "PC", "온도", "제어", "통신", "전원") { Add-Line "분류" $c $c }

if ($WithVariants) {
    Add-Line "설비" "RFL003" "RFL-003"
    Add-Line "설비" "RFL004" "RFL-004"
    Add-Line "설비" "PLT017" "PLT-017"
    Add-Line "설비" "GRD009" "GRD-009"
    Add-Line "설비" "FLT009" "FLT-009"
}

$text += ($new -join $nl) + $nl
[System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($true)))

# 검증을 통과했으니 설정 폴더를 만든다
New-Item -ItemType Directory -Path $dest | Out-Null
Copy-Item "$source\column_map.csv" $dest
Move-Item $path $origPath
"추가한 줄: $($new.Count)개 (설비 $equipmentAdded개) -> $origPath"
"저장소 config 는 바뀌지 않았습니다."
