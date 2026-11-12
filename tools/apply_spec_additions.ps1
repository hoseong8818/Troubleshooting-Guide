# spec.md에 구현에서 확정된 규칙을 반영한다 (docs/proposals/spec-additions.md 기준).
# 사용: powershell -ExecutionPolicy Bypass -File C:\mywork\tools\apply_spec_additions.ps1 [-Optional]
# -Optional: NFR-007 상태 기록과 US3 수용 기준 8, 9를 함께 반영한다.
param([switch]$Optional)

$p = "C:\mywork\specs\001-trouble-merge\spec.md"
$enc = New-Object System.Text.UTF8Encoding($false)
$lines = [System.IO.File]::ReadAllLines($p, $enc)

if ($lines -match "FR-024") {
    "이미 반영되어 있음(FR-024 있음). 아무것도 바꾸지 않았습니다."
} else {
    $fr21 = '- **FR-021 (값 정리)** 출력의 문자열 값은 앞뒤 공백을 뺀 값이다. 가운데 공백과 대소문자는 바꾸지 않는다. 설비와 분류는 FR-009, FR-004의 변환 규칙을 적용한 값이다.'
    $fr22 = '- **FR-022 (확인 필요 셀)** ''확인 필요''인 컬럼의 셀에는 문자열 `확인 필요`를 쓰고, 원래 값은 ''원래 값''에 FR-015 형식으로 남긴다. 해당 행의 다른 컬럼 값은 그대로 둔다.'
    $fr23 = '- **FR-023 (출력 파일 이름)** `--out` 폴더에 `merged.xlsx`, `audit_<실행 시각>.txt`, `names_<실행 시각>.txt`를 만든다. `merged.xlsx`가 이미 있으면 `merged_<실행 시각>.xlsx`로 만든다. 어떤 기존 파일도 덮어쓰지 않는다. 읽은 파일이 하나도 없으면(전부 실패) 통합 파일은 만들지 않고 감사 결과와 이름 대응표는 만든 뒤 종료 코드 2로 멈춘다.'
    $fr24 = '- **FR-024 (감사 결과 형식)** 감사 결과는 줄마다 `이름: 값` 형식이다. 낼 줄은 입력 행 수, 출력 행 수, 행 수 일치(예 또는 아니오), 확인 필요 건수, 중복 의심 건수, 읽은 파일 수, 처리 결과, 실패한 파일(없으면 `없음`), 대응 규칙 없음 시트, 양식 모호 시트, 대응 중복 양식 시트, 대응 규칙 없는 컬럼, 필수값 빈 칸(파일, 시트, 행 번호, 컬럼명), 숨긴 시트 수와 파일별 숨긴 시트, 수식 값 없음 건수, 수식 방지 접두사 처리 건수다. `처리 결과: 처리된 파일 없음` 줄은 읽은 파일이 0개일 때만 내며, 이때 입력과 출력 행 수가 모두 0이라 `행 수 일치: 예`가 나오므로 정상 처리로 오해하지 않게 하는 줄이다. 감사 결과에는 셀 값이 나오지 않는다.'
    $nfr11 = '- **NFR-011 (통합 시트의 이름 노출)** 통합 시트의 ''원본 파일''과 ''원본 시트''에는 마스킹하지 않은 실제 이름이 들어간다. 통합 시트와 그 안의 셀 값은 입력과 같은 수준의 비밀로 다루고, 출력 폴더는 클라우드 동기화에서 제외한다. 로그와 감사 결과에 대한 NFR-009와 구분한다.'

    $out = New-Object System.Collections.Generic.List[string]
    foreach ($l in $lines) {
        if ($l.StartsWith('- **FR-011**')) { $l = $l + ' 출력 일자 셀은 날짜형이 아니라 `YYYY-MM-DD` 문자열이다.' }
        if ($l.StartsWith('- **FR-016**')) { $l = $l + ' 같은 설비 안에서 일자가 ''확인 필요''이거나 비어 있는 행은 일자가 있는 행 뒤에 둔다.' }
        if ($l.StartsWith('- **FR-019**')) {
            $l = $l.Replace('종료 코드는 0 정상, 1 일부 파일 실패, 2 전부 실패 또는 설정 오류다.', '종료 코드는 0 정상, 1 일부 실패(열 수 없는 파일, 대응 규칙 없음 시트, 양식 모호 시트가 하나라도 있음), 2 전부 실패 또는 설정 오류다. 예외의 메시지와 종류는 출력하지 않는다.')
        }
        if ($l.StartsWith('- **NFR-009')) {
            $l = $l.Replace('실제 이름을 내는 옵션은 사람이 켠다.', '파일명과 시트명을 그대로 내는 옵션은 `--show-names`이며 사람이 켠다.')
        }
        if ($Optional) {
            if ($l.StartsWith('- **NFR-007')) { $l = $l + ' 현재 구현은 `.xlsm`과 `.xls` 거부까지만 하며 크기와 행 수 상한은 데모 범위에서 구현하지 않았다. 실제 업무 이관 전에 숫자를 정해 구현한다.' }
            $old = '모든 파일이 실패하면 통합 파일을 만들지 않고 실행을 멈춘다.'
            if ($l.Contains($old)) { $l = $l.Replace($old, $old + ' 이때도 감사 결과는 남긴다.') }
        }
        $out.Add($l)
        if ($l.StartsWith('- **FR-020**')) { $out.Add($fr21); $out.Add($fr22); $out.Add($fr23); $out.Add($fr24) }
        if ($l.StartsWith('- **NFR-010')) { $out.Add($nfr11) }
        if ($Optional -and $l.StartsWith('7. 감사 결과와 로그에는')) {
            $out.Add('8. Given 대응 규칙 없음이나 양식 모호인 시트가 있을 때, Then 감사 결과에 그 시트가 파일 표시명과 시트 표시명으로 나온다.')
            $out.Add('9. Given 모든 파일이 실패했을 때, Then 통합 파일은 만들지 않고 감사 결과는 남기며 `처리 결과: 처리된 파일 없음`이 나온다.')
        }
    }
    [System.IO.File]::WriteAllLines($p, $out, $enc)
    "반영했습니다. 아래 diff를 확인하세요."
}
git -C C:\mywork diff --stat
