# Phase 7 사람 몫 작업 초안 (2026-11-12, 데모용)

제가 할 수 없는 작업(보호 파일, 보안 설정, 외부 문의, 실제 데이터)은 초안과 명령으로 만들었습니다. 사람이 확인하고 반영합니다.

## A. T049 CLAUDE.md 정리 제안

점검 결과
- 줄 수 49줄(200줄 미만). 통과.
- `git log -p --all`에서 비밀 값 패턴(api key, secret, token, password, 개인 키, AKIA·sk- 형태)을 검색했습니다. 걸린 것은 훅 코드의 정규식 변수 `TOKEN_RE` 4줄뿐이며 비밀 값이 아닙니다.
- git이 추적하는 데이터 파일은 `data/fake/*.xlsx`, `expected/*.xlsx`, `config/*.csv`, `docs/proposals/config/*.csv`뿐이고 모두 가짜 데이터·가짜 설정입니다. `.env`, 키 파일은 추적되지 않습니다.

고칠 곳
- **'현재 상태' 단락이 낡았습니다.** "아직 구현 코드가 없다"는 사실과 다릅니다. 교체 제안:

> 구현은 Phase 6까지 끝났고 테스트 69개가 통과한다(`src/trouble_merge/`, `tests/`). 남은 일은 Phase 7(T047~T051)과 사람 작업(T003a, T003b, T005b, 실제 데이터 시험)이다. 보호 훅(`protect_files.py`)은 구현·등록했고, 보호할 파일은 모두 커밋되어 실제로 보호된다. 문서와 출력 문구는 모두 한국어다.

- 계획된 구조 절의 "아직 존재하지 않을 수 있다"는 "현재 존재한다"로 바꾸고, 출력 파일 이름(`merged.xlsx`, `audit_*.txt`, `names_*.txt`)과 `--show-names`, 전부 실패 때도 감사 결과를 남긴다는 점을 한 줄씩 추가합니다.
- 보호 장치 절에 "`requirements.txt`는 해시 고정, `pip install --require-hashes`" 한 줄을 추가합니다.

## B. T050 실무 이관 체크리스트 (한 장)

**실행 전**
- [ ] 회사 PC에서 Python 3.13 사용 허용, pip 접근(또는 사내 미러), 백신·DLP 영향 확인
- [ ] 실제 폴더 3곳(`real_data`, `real_out`, `real_config`)을 저장소 밖에 만들고 운영체제 권한 설정(T003a)
- [ ] 정보보호 부서 답변 받기(T003b): 컬럼명 노출 가능 여부, 크기·행 수 상한
- [ ] 컬럼 설명 문서(`docs/columns-*.md`)를 실제 컬럼명으로 갱신(문서, 가짜 데이터 생성기, `column_map.csv` 세 곳)
- [ ] 실제 변환표·컬럼 대응표를 `real_config`에 두고 `--config`로 지정(NFR-010)
- [ ] `requirements.txt`로 새 환경에 설치해 해시 검증 시험(README '미검증' 줄 정리)

**실제 데이터로 직접 실행 (본인 터미널에서만)**
- [ ] 병합 셀이 있는지, 헤더 행 위치가 1행이 맞는지(다르면 `column_map.csv`의 헤더 행 열 사용)
- [ ] 수식 셀의 저장된 계산값이 있는지(`수식 값 없음 건수`가 0인지)
- [ ] 끝쪽 빈 서식 행 때문에 느려지거나 메모리가 크게 늘지 않는지
- [ ] 설비 코드 형식, 날짜 형식이 가짜 데이터와 다른지(`확인 필요` 건수가 지나치게 많지 않은지)
- [ ] 처리 시간 측정(NFR-005)과 기록
- [ ] 감사 결과의 입력 행 수와 원본 엑셀의 행 수를 사람이 눈으로 대조

**실행 후**
- [ ] 결과 파일은 `real_out`에서만 열고, 클라우드 동기화 대상이 아닌지 확인
- [ ] 결과·감사 결과·로그·`names_*.txt`를 AI에 붙여 넣지 않았는지
- [ ] 보관 30일 후 삭제 일정 기록(README), `names_*.txt`가 가장 먼저 삭제 대상

## C. 외부 요청문 초안

**정보보호 부서 (T003b)**
> 설비 트러블슈팅 엑셀 통합 도구를 만들고 있습니다. 개발에는 가짜 데이터만 쓰고 실제 엑셀은 AI가 읽지 않습니다. 다음을 확인 부탁드립니다. (1) 엑셀의 컬럼 이름(헤더 문구)을 AI 도구에 알려도 되는지, 안 되면 어느 수준까지 되는지. (2) 입력 파일 크기와 시트당 행 수의 상한을 몇으로 둘지(압축 해제 후 크기 포함). (3) 통합 결과 파일의 보관 기간(현재 임시로 30일)과 저장 위치 정책.

**IT (T003a, T005b)**
> 개발 PC(Windows 11)에서 AI 코딩 도구를 쓰면서 실제 데이터 폴더 접근을 막고 싶습니다. (1) 실제 폴더 3곳을 별도 계정 또는 NTFS 권한으로 AI 도구 실행 계정에서 읽을 수 없게 하는 방법. (2) 사용자 홈의 도구 설정에서 보호 훅을 끄지 못하게 하는 관리자 설정(managed settings) 적용 가능 여부. (3) Python 3.13과 openpyxl 3.1.5, et_xmlfile 2.0.0을 해시 검증으로 설치할 수 있는 승인된 경로(사내 미러 포함).

## D. T047 교차 의심 시험 계획 (사람, 다른 사람의 눈으로)

데이터 쪽
- 가짜 복사본에 빈 칸·단위가 다른 값·처음 보는 파일·중복 행을 넣고 `out/`에서 `확인 필요`와 `중복 의심`이 맞게 나오는지 본다.
- 입력 파일이 실행 전후 같은지(해시) 본다.

보호 쪽 (보호 훅이 막는지 시도해 본다. 막히면 성공)
- 보호 파일(`tests/…`, `expected/…`)에 Edit·Write·`Set-Content`·`copy`·`git rm --cached` 시도
- 새 파일을 `tests/`에 `tasks.md`에 없는 이름으로 만들기 시도
- `.claude/settings.json`·`.claude/hooks/` 수정, `.claude/settings.local.json` 새로 만들기
- 저장소 밖 실제 폴더 경로를 Read·Bash·PowerShell로 열기 시도
- `python .claude/hooks/hook_test.py`가 `FAIL 0`인지 확인
- 결과(막힌 것과 뚫린 것)를 날짜와 함께 기록하고 뚫린 것은 고치거나 README '남은 위험'에 적는다.

## E. 사람이 실행할 명령

**1) tasks.md 체크박스 갱신** (끝난 작업만 `[x]`로. T017e와 T003a·T003b, T047~T051은 제외)

```powershell
$p = "C:\mywork\specs\001-trouble-merge\tasks.md"
$done = "T005a","T006","T007","T008","T009","T010","T011","T012","T013","T014","T015","T016","T017","T017a","T017b","T017c","T017d","T018","T019","T020","T021","T022","T023","T024","T025","T026","T027","T028","T029","T030","T031","T032","T033","T034","T035","T036","T037","T038","T039","T040","T041","T042","T042a","T042b","T043","T043a","T043b","T043c","T044","T045","T046"
$lines = [System.IO.File]::ReadAllLines($p)
$out = foreach ($l in $lines) { $m = $l; foreach ($id in $done) { if ($l.StartsWith("- [ ] $id ")) { $m = "- [x] $id " + $l.Substring(("- [ ] $id ").Length) } }; $m }
[System.IO.File]::WriteAllLines($p, $out, (New-Object System.Text.UTF8Encoding($false)))
git -C C:\mywork diff --stat
```

**2) 보호 파일 SHA-256 목록 갱신** (경로는 기존 목록과 같게)

```powershell
Get-ChildItem C:\mywork\tests, C:\mywork\expected, C:\mywork\data\fake, C:\mywork\config, C:\mywork\.claude\hooks, C:\mywork\.claude\agents -Recurse -File | Where-Object { $_.FullName -notmatch "__pycache__" } | Get-FileHash -Algorithm SHA256 | Sort-Object Path | ForEach-Object { "$($_.Hash)  $($_.Path)" } | Set-Content D:\hash_backup\protected_2026-11-12_phase6.txt -Encoding utf8
```

**3) 정리** (검증용 폴더, 이미 쓰임이 끝난 제안서)

```powershell
Remove-Item -Recurse -Force C:\temp\wheels, C:\temp\venv_check
```
