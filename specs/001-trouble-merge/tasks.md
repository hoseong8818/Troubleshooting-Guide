# 작업 목록: 트러블슈팅 엑셀 통합 도구

기능 폴더: `specs/001-trouble-merge` · 명세: `spec.md` · 계획: `plan.md`

## 읽는 법

- `[P]`: 다른 파일을 건드리고, 끝나지 않은 다른 작업에 의존하지 않아 동시에 해도 된다.
- `[US1]`: 어느 사용자 스토리의 작업인지.
- `[사람]`: 사람이 직접 하거나 확인해야 한다. 에이전트는 초안만 만들고 확정하지 않는다.
- 스토리 안의 순서는 테스트 → 구현이다. 테스트는 먼저 실패해야 한다. 실패하지 않으면 테스트가 잘못된 것이다.
- Phase가 끝날 때마다 `python -m unittest` 결과를 보고, 감사관의 판정을 받고, 커밋한다.
- 같은 실패는 3회까지 고치고 그래도 안 되면 멈추고 보고한다. 테스트 파일은 고치지 않는다.
- 범위가 반이 되어야 하면 Phase 6과 7은 뒤로 미룬다. 중간 점검에서 통과하는 테스트가 하나도 없으면 범위를 줄인다.

## Phase 1: Setup

- [x] T001 `C:\mywork`에서 `git init`, `.gitignore`(`out/`, `__pycache__/`, `.env`, `data/real/`), 폴더 뼈대(`src/trouble_merge`, `tests`, `config`, `data/fake`, `expected`, `tools`, `out`) 만들기
- [x] T002 [P] `CLAUDE.md` 초안(실행·테스트 명령, 폴더 구조, 헌법 요약, 200줄 미만)
- [x] T003 [P] `.claude/settings.json`에 deny 규칙(저장소 밖 실제 폴더 3곳과 다운로드·바탕화면, `data/real`을 쓰면 그것도: Read, Grep, Glob, Bash, PowerShell)과 훅 등록 자리 만들기. 이 파일은 사람이 만든다
- [ ] T003a [사람] 실제 데이터 폴더 3곳(`real_data`, `real_out`, `real_config`)의 위치를 정하고, IT와 함께 운영체제 권한(NTFS 권한 또는 Claude를 별도 계정·WSL·VM에서 실행)을 설정한다. 설정 날짜를 기록한다(헌법 원칙 1)
- [ ] T003b [사람] 정보보호 부서에 확인한다: 컬럼명을 AI에 알려도 되는지, 설비 코드·분류명이 기밀인지, 입력 파일 크기·행 수 상한 숫자(NFR-007)

## Phase 2: Foundational (보호, 가짜 데이터, 첫 커밋)

완료 조건: 보호 훅이 시도 목록을 모두 막고, 가짜 데이터와 정답과 규칙 표 초안이 있고, 커밋으로 보호가 켜진 상태.

- [x] T004 `.claude/hooks/hook_test.py`: 보호 대상 수정 시도 목록(Edit, Write, PowerShell 복사, Python 저장, 경로를 변수에 담기, 입력이 깨진 경우, git 실행 실패, `..`·역슬래시·대소문자 변형 경로, `git add`만 한 파일, 보호 장치 자체 수정, 삭제·이름 변경(`rm`, `del`, `ren`, `Remove-Item`), `git rm --cached`·`reset`·`restore`·`stash`·`config`, 8.3 짧은 이름·정션·`\\?\`·대체 데이터 스트림 경로, `.claude/settings.local.json` 새로 만들기, `tasks.md`에 없는 보호 폴더 새 파일, `docs/constitution.md`·`CLAUDE.md`·`specs/` 수정, `python -c`로 실제 폴더 열기)을 먼저 쓴다. 지금은 FAIL이어야 한다.
- [ ] T005 `.claude/hooks/protect_files.py`: 보호 대상(`expected/`, `data/fake/`, `tests/`, `config/*.csv`, `.claude/settings.json`, `.claude/hooks/`, `.claude/agents/`) 중 git에 추적되는 파일의 Edit·Write를 막는다. 보호 경로가 든 쓰기 명령(Bash·PowerShell)도 막는다. 새 파일은 허용한다. 훅 전체를 `try/except`로 감싸 모든 실패에서 종료 코드 2로 끝낸다. `hook_test.py`가 FAIL 0이 될 때까지(T004 이후). 허용·차단 규칙의 세부(git 명령 허용 목록, `realpath` 해석, `tasks.md` 허용 파일 목록, 링크 생성 차단)는 plan 보호 설계를 따른다
- [ ] T005a 훅이 에이전트별 도구 호출을 구분할 수 있는지, `settings.local.json`이 훅·권한에 미치는 영향이 어떤지 가짜 폴더에서 시험하고 결과를 기록한다. 구분하지 못하면 T006·T007의 에이전트는 파일을 쓰지 않는 방식으로 정의한다(T005 이후)
- [ ] T006 [P] `.claude/agents/format-analyst.md` (Read, Grep, Glob, Write는 `docs/columns-*.md`·`docs/proposals/**`만). 이 파일은 사람이 만든다(`.claude/` 신규 파일은 에이전트에게 차단). 경로 제한은 T005a 결과에 따라 `.claude/settings.json` 권한 규칙이나 훅으로 건다
- [ ] T007 [P] `.claude/agents/merge-auditor.md` (Read, Grep, Glob, `Bash(python -m unittest*)` 한 가지만). Write·Edit 없음. 이 파일은 사람이 만든다. 명령 제한은 T005a 결과에 따라 `.claude/settings.json` 권한 규칙이나 훅으로 건다
- [ ] T008 [사람] `docs/columns-A.md`: 양식 A의 컬럼 설명을 사람이 쓴다(컬럼명은 실제와 같게, 값은 쓰지 않는다)
- [ ] T009 `tools/make_fake_data.py`로 `data/fake/`에 양식 A 가짜 엑셀 두 개를 만든다. 이미 있는 파일은 덮어쓰지 않고 중단한다(FR-020). (T008 이후)
  - `A_basic.xlsx`(Phase 3용): 정상 행, 날짜 형식(날짜형, 시각 붙은 값, 엑셀 일련번호, `2026.11.07`, `2026/11/7`), 변환 표에 있는 분류, 표준에 없는 컬럼, 설비·일자·증상이 같은 행, 병합 셀, 완전히 빈 행
  - `A_trap.xlsx`(Phase 4용): 읽을 수 없는 날짜, 설비 코드 빈 칸, 변환 표에 없는 설비 코드와 분류, 필수값 빈 행, 사유가 여러 개인 행, 설비가 '확인 필요'인 중복 후보 행, 숨긴 시트와 `veryHidden` 시트, 표식 값 `LEAK-`
- [ ] T010 [사람] `expected/A_basic_expected.xlsx`, `expected/A_trap_expected.xlsx`: 가짜 입력의 정답을 사람이 만들거나 에이전트 초안을 한 행씩 검토해 확정한다(T009 이후)
- [ ] T011 [사람] `config/column_map.csv`, `config/conversion_map.csv` 양식 A 초안을 검토·확정한다(T008, T010 이후)
- [ ] T012 `hook_test.py` FAIL 0 확인, 첫 커밋으로 보호를 켠다(T005, T010, T011 이후). 이후 `expected/`, `data/fake/`, `tests/`, `config/*.csv`의 기존 파일 수정은 막힌다. 커밋 직후 [사람] 보호 대상 파일의 SHA-256 목록을 저장소 밖에 보관한다. 각 Phase 끝에 사람이 자기 터미널에서 대조한다.

## Phase 3: 스토리 1 (P1) 엑셀 한 개를 표준 표로 바꾼다

완료 조건: `A_basic` 입력이 정답(`A_basic_expected.xlsx`)과 한 행씩 일치하고, 행 수가 같고, 기타 항목과 분류 변환이 맞는다. `A_trap`은 Phase 4에서 다룬다.

- [ ] T013 [P] [US1] `tests/test_compare.py`: `A_basic` 통합 결과를 `expected/A_basic_expected.xlsx`와 한 행씩 비교. 출력은 임시 폴더에 쓴다
- [ ] T014 [P] [US1] `tests/test_row_count.py`: 입력과 출력 행 수 일치(빈 행 제외 규칙 포함)
- [ ] T015 [P] [US1] `tests/test_other_columns.py`: 표준에 없는 컬럼이 `이름: 값`으로 보관됨
- [ ] T016 [P] [US1] `tests/test_dates.py`: 날짜 세 형식이 `YYYY-MM-DD`로 통일됨
- [ ] T017 [P] [US1] `tests/test_duplicates.py`: 설비·일자·증상이 같은 행에 '중복 의심', 행은 지우지 않음
- [ ] T017a [P] [US1] `tests/test_normalize.py`: 설비 코드와 변환 표 키의 공백·대소문자 정리(FR-009, FR-018)
- [ ] T017b [P] [US1] `tests/test_merged_cells.py`, `tests/test_blank_rows.py`: 병합 셀 채움과 빈 행 제외(FR-014)
- [ ] T017c [P] [US1] `tests/test_input_unchanged.py`: 실행 전후 입력 파일 해시가 같음(NFR-003)
- [ ] T017d [P] [US1] `tests/test_formula_injection.py`: `=`, `+`, `-`, `@`로 시작하는 입력 값이 출력에서 수식 셀이 되지 않고 `'`가 붙음(NFR-008, SC-009). 이를 위해 가짜 데이터에 해당 값 행이 필요하므로 T009의 입력에 포함한다
- [ ] T017e [P] [US1] `tests/test_input_limits.py`: `.xlsm`·`.xls`, 크기·행 수 상한 초과 파일은 건너뛰고 감사에 보고, 나머지는 처리(NFR-007, SC-010)
- [ ] T018 [US1] 위 테스트가 모두 실패하는 것을 확인하고 실패 상태로 커밋. 모듈이 없어서 나는 `ImportError`가 아니라 기대값 불일치로 실패하는지 확인한다. 커밋 메시지는 `test:`로 시작한다(T013~T017c 이후)
- [ ] T019 [P] [US1] `src/trouble_merge/mapping.py`: 컬럼 대응 표 읽기와 양식 판별
- [ ] T020 [P] [US1] `src/trouble_merge/dates.py`: 날짜 읽기와 `YYYY-MM-DD` 변환
- [ ] T021 [P] [US1] `src/trouble_merge/reader.py`: 보이는 시트 읽기, 숨긴 시트 건너뛰기
- [ ] T022 [US1] `src/trouble_merge/merge.py`: 표준 컬럼 옮기기, 기타 항목, 분류 변환, 중복 의심 표시, 정렬(T019~T021 이후)
- [ ] T023 [US1] `src/trouble_merge/writer.py`와 `cli.py`: 통합 엑셀을 `out/`에 쓰고, 이미 있으면 실행 시각을 붙인다(T022 이후)
- [ ] T024 [US1] `python -m unittest` 통과 확인, 감사관 판정, 커밋

## Phase 4: 스토리 2 (P1) 판단이 안 되면 '확인 필요'로 남긴다

완료 조건: 네 가지 경우가 모두 '확인 필요'와 원래 값과 이유로 남고, `A_trap`이 `A_trap_expected.xlsx`와 한 행씩 일치하며, 변환 표에 한 줄 추가하면 다음 실행에서 사라진다.

- [ ] T025 [P] [US2] `tests/test_needs_check.py`: `A_trap`의 설비 코드 빈 칸, 설비 코드 미등록, 분류 미등록, 필수값 빈 행, 날짜 해석 불가, 사유 여러 개 서식(FR-015), `expected/A_trap_expected.xlsx`와 한 행씩 비교
- [ ] T026 [P] [US2] `tests/test_conversion_extend.py`: 변환 표에 한 줄을 더하면 해당 행이 자동 분류됨(임시 사본으로 시험)
- [ ] T027 [US2] 테스트가 실패하는 것을 확인하고 실패 상태로 커밋(T025, T026 이후)
- [ ] T028 [US2] `src/trouble_merge/check.py`: 확인 필요 판정, 원래 값, 이유 문구(T027 이후)
- [ ] T029 [US2] `merge.py`에 통합. 설비 코드 정리(공백, 대소문자)와 변환 표 일치(T028 이후)
- [ ] T030 [US2] 통과 확인, 감사관 판정, 커밋

## Phase 5: 스토리 3 (P2) 감사 결과와 안전한 로그

완료 조건: 감사 결과가 규칙대로 나오고, 정상·오류 실행의 로그와 stderr에 표식 값이 없다.

- [ ] T031 [P] [US3] `tests/test_audit.py`: 행 수 일치, 대응 규칙 없는 컬럼명, 빈 필수값(행 번호·컬럼명), 확인 필요 건수, 숨긴 시트 수(`veryHidden` 포함), 실패 파일명, 수식 값 없음 건수
- [ ] T032 [P] [US3] `tests/test_log_safe.py`: 정상·오류 실행에서 `LEAK-` 표식이 로그, stderr, `warnings` 출력에 0건. openpyxl 예외 메시지가 그대로 나오지 않는지 포함. 등록되지 않은 컬럼명(헤더 행 오인 상황 포함)과 기본 설정의 파일명·시트명이 로그·감사 결과에 그대로 나오지 않는지도 본다(NFR-009, SC-011)
- [ ] T033 [US3] 테스트가 실패하는 것을 확인하고 실패 상태로 커밋(T031, T032 이후)
- [ ] T034 [P] [US3] `src/trouble_merge/safe_log.py`: 고정 필드(종류, 파일명, 시트명, 행 번호, 컬럼명, 건수)만 받는 로그
- [ ] T035 [P] [US3] `src/trouble_merge/audit.py`: 감사 결과 텍스트
- [ ] T036 [US3] 모든 로그와 예외를 `safe_log` 한 길로 통일(T034 이후)
- [ ] T037 [US3] 통과 확인, 감사관 판정, 커밋

## Phase 6: 스토리 4 (P2) 여러 파일과 새 양식

완료 조건: 양식 2~3종, 파일 5~8개가 한 시트로 모이고, 새 양식은 표 줄 추가만으로 처리된다.

- [ ] T038 [사람] `docs/columns-B.md`(, `columns-C.md`): 양식 B, C의 컬럼 설명
- [ ] T039 `tools/make_fake_data.py` 확장: 양식 B, C와 파일 5~8개(T038 이후)
- [ ] T040 [사람] `expected/` 정답에 양식 B, C와 통합 정답을 추가하고 `config/column_map.csv`에 줄을 추가한다(T039 이후)
- [ ] T041 [P] [US4] `tests/test_multi_file.py`: 모든 파일의 모든 행이 한 시트에 모이고 정답과 일치
- [ ] T042 [P] [US4] `tests/test_new_form.py`: 새 양식은 코드 수정 없이 표 줄 추가만으로 처리됨(임시 사본으로 시험). 규칙이 없는 양식은 감사에 '대응 규칙 없음'
- [ ] T042a [P] [US4] `tests/test_form_ambiguous.py`: 동점 양식은 '양식 모호'로 보고하고 건너뜀(FR-017)
- [ ] T042b [P] [US4] `tests/test_multi_sheet.py`: 보이는 시트 여러 개 읽기, 시트별 양식 판별, `veryHidden` 시트 제외
- [ ] T043 [P] [US4] `tests/test_partial_failure.py`: 열 수 없는 파일은 건너뛰고 나머지는 처리하며(종료 코드 1) 전부 실패하면 멈춤(종료 코드 2)
- [ ] T043a [P] [US4] `tests/test_output_name.py`: 출력 파일이 이미 있으면 덮어쓰지 않고 새 이름으로 만듦
- [ ] T044 [US4] 테스트가 실패하는 것을 확인하고 실패 상태로 커밋(T041~T043a 이후)
- [ ] T045 [US4] `cli.py`와 `merge.py`: 여러 파일 순회, 실패 파일 건너뛰기(T044 이후)
- [ ] T046 [US4] 통과 확인, 감사관 판정, 커밋

## Phase 7: Polish

- [ ] T047 [사람] 교차 의심: 다른 사람의 눈으로 빈 칸, 단위가 다른 값, 처음 보는 파일, 중복 행을 넣어 보고, 보호를 뚫어 본다. 결과를 기록하고 고치거나 이유를 적는다
- [ ] T048 `README.md`: 설치, 실행, 테스트 명령, 남은 위험(샌드박스 없음, 훅 우회로, 운영체제 권한 설정 여부, 회사 PC 확인, 가짜 데이터와 실제 엑셀의 특성 차이), `out/` 보관 기간·동기화 제외, 실제 결과·로그를 AI에 붙여 넣지 않는다는 규칙, 설치 파일 버전 고정과 해시 검증
- [ ] T049 `CLAUDE.md` 정리, 200줄 미만, `git log -p`에 비밀 값이 없는지 확인
- [ ] T050 [사람] 실무 이관 체크리스트 한 장: 회사 PC 도구 확인, 컬럼 설명 문서 작성, 정보보호 부서 확인, 실제 데이터로 직접 실행(병합 셀, 수식 계산값, 끝쪽 빈 서식 행, 헤더 행 위치가 가짜 데이터와 다른지 점검), 처리 시간 측정(NFR-005), `out/` 보관 기간과 동기화 제외, 실제 결과·로그를 AI에 붙여 넣지 않는다는 확인, 실제 변환표의 저장소 밖 위치 확인(NFR-010)
- [ ] T051 `/speckit-converge` 방식으로 명세와 구현을 대조해 빠진 것을 채운다(Spec Kit을 쓰면). 쓰지 않으면 수용 기준 목록과 테스트를 사람이 한 줄씩 대조한다

## 의존 관계

```
Phase 1 → Phase 2 → Phase 3 → Phase 4 → Phase 5
                                  └──────→ Phase 6 → Phase 7
```

- Phase 2의 T012(첫 커밋)가 끝나기 전에는 보호가 켜지지 않는다.
- Phase 3의 구현(T019 이후)은 T018의 실패 커밋 뒤에 한다.
- Phase 4와 5는 둘 다 Phase 3 뒤에 오지만 `merge.py`를 같이 건드려서 동시에 하지 않는다. Phase 6은 Phase 4 뒤에 시작할 수 있다.
- 정답 파일 `A_trap_expected.xlsx`가 확정되기 전에는 Phase 4의 구현을 시작하지 않는다.
- `[P]`가 붙은 작업은 서로 다른 파일만 건드리는지 시작 전에 한 번 확인한다.
