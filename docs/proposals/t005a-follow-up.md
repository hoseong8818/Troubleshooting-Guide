# T005a 결과 반영안 (사람이 적용)

`.claude/`, `tasks.md`, `plan.md`는 보호 훅이 에이전트의 수정을 막는다. 이 문서의 내용은 사람이 직접 적용한다.
선택: 에이전트는 **조합안**(2026-11-12).

## 1. T005a 시험 결과

| 질문 | 결과 |
|---|---|
| 서브에이전트 호출에도 훅이 적용되는가 | 적용된다(실측: 쓰기와 `python -c`가 막히고 읽기는 통과) |
| 훅 입력에 호출자 식별 필드가 있는가 | 문서상 `agent_id`, `agent_type`이 있다. 이 PC에서 실제로 오는지는 실측하지 않았다 |
| `settings.local.json`이 deny를 뒤집는가 | 못 뒤집는다(deny 우선, 문서) |
| `settings.local.json`이 프로젝트 훅을 지우는가 | 못 지운다(훅은 병합, 문서) |
| 훅을 통째로 끌 수 있는가 | `disableAllHooks: true`로 끌 수 있다. 사용자·프로젝트·로컬 설정 어디서나 쓸 수 있다(문서) |
| 에이전트 정의의 `tools`에 명령·경로 패턴을 쓸 수 있는가 | 못 쓴다. 도구 이름만 쓰고, `disallowedTools`도 도구 전체를 제거한다(문서) |

## 2. 에이전트 정의 복사 (사람이 터미널에서)

`C:\mywork`에서 PowerShell로 한 줄씩 실행한다.

```powershell
New-Item -ItemType Directory -Force .claude\agents
```

```powershell
Copy-Item docs\proposals\agents\format-analyst.md .claude\agents\format-analyst.md
```

```powershell
Copy-Item docs\proposals\agents\merge-auditor.md .claude\agents\merge-auditor.md
```

복사한 뒤 새 세션을 열거나 `/agents`에서 두 에이전트가 보이는지 확인한다.

| | format-analyst | merge-auditor |
|---|---|---|
| tools | Read, Grep, Glob | Read, Grep, Glob, Bash |
| disallowedTools | (없음, 쓰기·실행 도구 자체가 없음) | Edit, Write, NotebookEdit |
| Bash | 없음 | 공용 훅의 허용 목록(`python -m unittest`, 읽기용 git 등)만 통과 |
| 산출물 | 대화로 제안(사람이 저장) | 대화로 판정 |

## 3. tasks.md 교체문

**T005a** (현재 `- [ ] T005a ...` 줄 전체를 바꾼다)

```
- [x] T005a 훅이 에이전트별 도구 호출을 구분할 수 있는지, `settings.local.json` 영향을 시험했다. 결과: 훅은 서브에이전트 호출에도 적용된다(실측). 훅 입력의 `agent_id`·`agent_type`은 문서상 있으나 실측하지 않았다. `settings.local.json`은 deny를 못 뒤집고 훅을 못 지우지만 `disableAllHooks`로 전체를 끌 수 있다. 에이전트별 제한은 훅이 아니라 에이전트 정의의 `tools`·`disallowedTools`와 공용 훅의 허용 목록으로 건다(조합안).
```

**T006** (줄 전체)

```
- [ ] T006 [P] `.claude/agents/format-analyst.md` (`tools: Read, Grep, Glob`. 파일을 쓰지 않고 초안은 대화로 낸다). 초안은 `docs/proposals/agents/format-analyst.md`에 있고 사람이 `.claude/agents/`로 복사한다(`.claude/` 신규 파일은 에이전트에게 차단)
```

**T007** (줄 전체)

```
- [ ] T007 [P] `.claude/agents/merge-auditor.md` (`tools: Read, Grep, Glob, Bash`, `disallowedTools: Edit, Write, NotebookEdit`). Bash는 공용 훅의 허용 목록이 제한한다. 초안은 `docs/proposals/agents/merge-auditor.md`에 있고 사람이 복사한다
```

**T005b** (T005a 바로 아래에 새 줄로 더한다)

```
- [ ] T005b [사람] `~/.claude/settings.json`으로 훅을 끄는 길을 막는다. `hook_test.py`에 시험을 먼저 더하고(실패 확인), `protect_files.py`에 `~/.claude` 쓰기 차단을 더해 통과시키고, `settings.json` deny에 `Edit(~/.claude/**)`와 명령 패턴(`*.claude/settings*`)을 더한다. T003a에서 IT와 관리자(managed) 설정에 훅을 넣을 수 있는지도 확인한다
```

## 4. plan.md 교체문

**에이전트 표의 두 행**

```
| format-analyst (양식 분석가) | `docs/`의 컬럼 설명 문서와 명세를 읽고 컬럼 대응 규칙 초안과 '규칙으로 판단하기 어려운 입력' 표를 **대화로** 제안한다(파일은 쓰지 않고 사람이 저장) | Read, Grep, Glob | 코드를 고칠 때, 실제 데이터 근처 |
| merge-auditor (감사관) | 구현을 모르는 상태로 테스트 결과와 보호 대상 변경 여부를 읽고 통과 또는 반려를 판정한다. 지적은 요약으로 돌려준다 | Read, Grep, Glob, Bash(공용 훅의 허용 목록만), disallowedTools: Edit, Write, NotebookEdit | 코드를 고치려 할 때 |
```

**"두 에이전트의 경로·명령 제한은 ..." 문단 전체를 아래로 바꾼다**

```
두 에이전트의 도구 권한은 에이전트 정의 파일로 다르게 둔다. 정의의 `tools`에는 명령·경로 패턴을 쓸 수 없고 `disallowedTools`도 도구 전체를 제거하므로, 명령 제한은 모든 호출에 걸리는 공용 훅(`protect_files.py`)의 허용 목록에 맡긴다. 훅은 서브에이전트 호출에도 적용된다(T005a 실측). format-analyst는 쓰기·실행 도구가 없고, merge-auditor는 Edit·Write가 없으며 Bash는 훅이 허용한 명령(`python -m unittest`, 읽기용 git 등)만 쓴다. 다만 `unittest`는 `tests/`와 `src/`의 코드를 실행하므로 "읽기 전용"은 아니다. 에이전트 정의 파일은 사람이 만든다. 훅 입력의 `agent_type`으로 에이전트별 규칙을 나누는 방법은 문서상 가능하나 실측하지 않아 쓰지 않는다.
```

## 5. 빈틈과 후속 작업(T005b)

`~/.claude/settings.json`에 `disableAllHooks: true`를 넣으면 보호 훅이 전부 꺼진다. 지금 훅은 저장소 밖 경로를 검사하지 않고 deny에도 `~/.claude`가 없다. T005b로 막는다. 근본 대책은 보호 훅을 관리자(managed) 설정에 두는 것이다(managed 훅은 사용자·프로젝트 설정의 `disableAllHooks`로 꺼지지 않는다, 문서). IT 협조가 필요하므로 T003a와 함께 확인한다.

## 6. 적용 후 확인

1. `/agents`에 두 에이전트가 보인다.
2. merge-auditor에게 파일 쓰기를 시켜 보고(도구가 없어 못 해야 한다), format-analyst에게 `python -m unittest` 실행을 시켜 본다(도구가 없어 못 해야 한다). 제가 두 시험을 실행할 수 있다.
3. 코스의 "도구 권한이 서로 다름" 기준의 증거로 두 결과를 남긴다.
