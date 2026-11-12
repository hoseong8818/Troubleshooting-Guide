#!/usr/bin/env python3
"""보호 훅 `protect_files.py` 시험 (tasks T004).

실행: python .claude/hooks/hook_test.py [훅 스크립트 경로]
통과 기준: 마지막 줄이 `FAIL 0`. T005(훅 작성) 전에는 전부 FAIL이 정상이다.

시험 방식: 임시 git 저장소(픽스처)를 만들고, 훅을 PreToolUse 입력(JSON, 표준입력)과
환경 변수 CLAUDE_PROJECT_DIR=<픽스처>로 실행해 종료 코드와 표준오류를 본다.
실제 저장소의 파일은 읽거나 바꾸지 않는다.

훅이 지킬 계약 (T005가 이 계약을 구현한다)
- 차단: 종료 코드 2 + 표준오류에 `차단:` 으로 시작하는 이유 한 줄.
- 판단 불가(입력이 깨졌거나 git·환경을 모를 때): 종료 코드 2 + `차단(판단 불가):`.
- 허용: 종료 코드 0, 표준오류에 '차단' 없음.
- 저장소 위치는 CLAUDE_PROJECT_DIR 로 안다(훅 파일 위치로 추정하지 않는다).
- 보호 대상: expected/, data/fake/, tests/, config/*.csv, .claude/(설정·훅·에이전트),
  docs/constitution.md, CLAUDE.md, specs/**/spec.md·plan.md·tasks.md.
- Edit·Write: 보호 대상이 git에 추적(커밋 또는 `git add`)되어 있으면 차단.
  아직 없는 새 파일은 '경로 문자열이 specs/001-trouble-merge/tasks.md 본문에 적힌 것'만 허용.
  단 .claude/ 아래는 새 파일도 차단. 경로는 realpath로 해석해서 판단한다
  (.., 역슬래시, 대소문자, \\\\?\\ 접두사, :스트림, 후행 점, 정션, 8.3 짧은 이름).
- Bash·PowerShell: 보호 경로가 있는 쓰기·삭제·이동 명령은 차단. 쓰기 동사가 있는데 대상이
  변수·문자열 이어붙이기라 해석할 수 없으면 차단(판단 불가). .git/ 쓰기와 정션·링크 생성도 차단.
- 인터프리터는 허용 목록만: python -m unittest*, python .claude/hooks/hook_test.py,
  python -m trouble_merge*, python tools/make_fake_data.py*. python -c, 그 밖의 스크립트,
  powershell -c, cmd /c 는 차단.
- git은 status, diff, log, show, add, commit(--amend 제외), switch -c 만 허용.
  rm, reset, restore, checkout --, stash, config, rebase, filter-*, commit --amend 는 차단.
"""
import ctypes
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
HOOK = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else HERE / "protect_files.py"
IS_WIN = os.name == "nt"

GIT = ["git", "-c", "core.autocrlf=false", "-c", "user.name=fixture",
       "-c", "user.email=fixture@example.invalid", "-c", "commit.gpgsign=false"]

TASKS_TEXT = """# fixture tasks
허용 목록(이름이 적힌 새 파일): tests/test_new_listed.py, expected/A_new_listed.xlsx,
.claude/hooks/listed_hook.py
"""

TRACKED = {
    "expected/A_basic_expected.xlsx": b"x",
    "data/fake/A_basic.xlsx": b"x",
    "tests/test_compare.py": "pass\n",
    "config/conversion_map.csv": "a,b\n",
    "config/column_map.csv": "a,b\n",
    ".claude/settings.json": "{}\n",
    ".claude/hooks/protect_files.py": "pass\n",
    ".claude/agents/format-analyst.md": "x\n",
    "docs/constitution.md": "x\n",
    "docs/request.md": "x\n",
    "CLAUDE.md": "x\n",
    "specs/001-trouble-merge/spec.md": "x\n",
    "specs/001-trouble-merge/plan.md": "x\n",
    "specs/001-trouble-merge/tasks.md": TASKS_TEXT,
    "src/trouble_merge/merge.py": "pass\n",
}


def git(root, *args):
    subprocess.run(GIT + list(args), cwd=root, check=True, capture_output=True)


def make_fixture(root: Path) -> None:
    for rel, content in TRACKED.items():
        p = root / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        if isinstance(content, bytes):
            p.write_bytes(content)
        else:
            p.write_text(content, encoding="utf-8")
    git(root, "init", "-q", "-b", "main")
    git(root, "add", "-A")
    git(root, "commit", "-q", "-m", "fixture")
    (root / "tests" / "test_staged.py").write_text("pass\n", encoding="utf-8")
    git(root, "add", "tests/test_staged.py")  # 커밋하지 않고 add만 한 파일
    if IS_WIN:
        subprocess.run(["cmd", "/c", "mklink", "/J", "src_link", "tests"],
                       cwd=root, capture_output=True)


def short_path(p: Path):
    """8.3 짧은 이름. 이 PC에서 만들 수 없으면 None."""
    if not IS_WIN:
        return None
    buf = ctypes.create_unicode_buffer(1024)
    if ctypes.windll.kernel32.GetShortPathNameW(str(p), buf, 1024) == 0:
        return None
    return buf.value if buf.value.lower() != str(p).lower() else None


def build_cases(root: Path):
    def a(*parts):
        return str(root.joinpath(*parts))

    def E(path):
        return "Edit", {"file_path": path, "old_string": "a", "new_string": "b"}

    def W(path):
        return "Write", {"file_path": path, "content": "x"}

    def B(cmd):
        return "Bash", {"command": cmd}

    def P(cmd):
        return "PowerShell", {"command": cmd}

    cases = []

    def add(name, spec, expect, **kw):
        cases.append(dict(name=name, tool=spec[0], input=spec[1], expect=expect, **kw))

    # --- 1. Edit·Write: 추적되는 보호 대상은 차단 ---
    add("Edit 추적된 tests", E(a("tests", "test_compare.py")), "block")
    add("Write 추적된 expected", W(a("expected", "A_basic_expected.xlsx")), "block")
    add("Edit 추적된 config csv", E(a("config", "conversion_map.csv")), "block")
    add("Edit 추적된 data/fake", E(a("data", "fake", "A_basic.xlsx")), "block")
    add("Edit 추적된 settings.json", E(a(".claude", "settings.json")), "block")
    add("Write 추적된 훅 파일", W(a(".claude", "hooks", "protect_files.py")), "block")
    add("Edit 추적된 에이전트 정의", E(a(".claude", "agents", "format-analyst.md")), "block")
    add("Edit 헌법", E(a("docs", "constitution.md")), "block")
    add("Edit CLAUDE.md", E(a("CLAUDE.md")), "block")
    for n in ("spec.md", "plan.md", "tasks.md"):
        add(f"Edit specs/{n}", E(a("specs", "001-trouble-merge", n)), "block")
    add("Edit git add만 한 파일", E(a("tests", "test_staged.py")), "block")
    add("Edit 상대 경로(cwd 기준)", E("tests/test_compare.py"), "block")

    # --- 2. 새 파일: 허용 목록만 ---
    add("Write 새 파일 tasks에 적힌 tests", W(a("tests", "test_new_listed.py")), "allow")
    add("Write 새 파일 tasks에 적힌 expected", W(a("expected", "A_new_listed.xlsx")), "allow")
    add("Write 새 파일 tasks에 없는 tests", W(a("tests", "test_unlisted.py")), "block")
    add("Write 새 파일 tasks에 없는 expected", W(a("expected", "unlisted.xlsx")), "block")
    add("Write 새 config csv(목록 밖)", W(a("config", "extra.csv")), "block")
    add("Write 새 .claude/settings.local.json", W(a(".claude", "settings.local.json")), "block")
    add("Write 새 .claude/hooks 파일(tasks에 적혀도 차단)",
        W(a(".claude", "hooks", "listed_hook.py")), "block")
    add("Write 새 .claude/agents 파일", W(a(".claude", "agents", "x.md")), "block")

    # --- 3. 보호 대상이 아닌 곳은 허용 ---
    add("Edit src 코드", E(a("src", "trouble_merge", "merge.py")), "allow")
    add("Write 새 src 코드", W(a("src", "trouble_merge", "new.py")), "allow")
    add("Write 새 docs/columns-A.md", W(a("docs", "columns-A.md")), "allow")
    add("Write 새 docs/proposals", W(a("docs", "proposals", "x.md")), "allow")
    add("Edit docs/request.md(보호 밖)", E(a("docs", "request.md")), "allow")
    add("Write out 결과", W(a("out", "audit.txt")), "allow")
    add("Write 저장소 밖 경로", W(str(root.parent / "outside.txt")), "allow")
    add("Read 보호 파일(읽기는 훅 대상 아님)",
        ("Read", {"file_path": a("tests", "test_compare.py")}), "allow")

    # --- 4. 경로 변형 (대상은 추적된 tests/test_compare.py) ---
    target = root / "tests" / "test_compare.py"
    add(".. 로 돌아가는 경로", E(a("src", "..", "tests", "test_compare.py")), "block")
    add("슬래시 섞인 경로", E(str(root).replace("\\", "/") + "/tests\\test_compare.py"), "block")
    if IS_WIN:
        add("대문자 경로", E(str(target).upper()), "block")
        add("\\\\?\\ 접두사 경로", E("\\\\?\\" + str(target)), "block")
        add("대체 데이터 스트림 경로", E(str(target) + ":evil"), "block")
        add("후행 점 경로", E(str(target) + "."), "block")
        add("후행 공백 경로", E(str(target) + " "), "block")
        junction_ok = (root / "src_link").exists()
        add("정션을 거친 경로", E(a("src_link", "test_compare.py")), "block",
            skip=None if junction_ok else "정션을 만들 수 없음")
        sp = short_path(target)
        add("8.3 짧은 이름 경로", E(sp or str(target)), "block",
            skip=None if sp else "이 PC에서 8.3 짧은 이름을 쓸 수 없음")
    else:
        for n in ("대문자 경로", "\\\\?\\ 접두사 경로", "대체 데이터 스트림 경로", "후행 점 경로",
                  "후행 공백 경로", "정션을 거친 경로", "8.3 짧은 이름 경로"):
            add(n, E(str(target)), "block", skip="Windows 전용")

    # --- 5. Bash: 보호 경로를 쓰는 명령은 차단 ---
    add("Bash cp로 tests에 복사", B("cp x.txt tests/test_compare.py"), "block")
    add("Bash > 로 tests 덮어쓰기", B("echo hi > tests/test_compare.py"), "block")
    add("Bash >> 로 expected 덧붙이기", B("echo hi >> expected/A_basic_expected.xlsx"), "block")
    add("Bash rm expected", B("rm expected/A_basic_expected.xlsx"), "block")
    add("Bash mv tests", B("mv tests/test_compare.py tests/old.py"), "block")
    add("Bash sed -i config", B("sed -i s/a/b/ config/conversion_map.csv"), "block")
    add("Bash cmd /c del data\\fake", B("cmd /c del data\\fake\\A_basic.xlsx"), "block")
    add("Bash powershell -c 로 저장", B('powershell -c "Set-Content tests/a.py x"'), "block")
    add("Bash python -c 로 저장", B("python -c \"open('tests/x.py','w').write('')\""), "block")
    add("Bash 허용 목록 밖 python 스크립트", B("python scratch.py"), "block")
    add("Bash .git 안에 쓰기", B("echo x > .git/hooks/pre-commit"), "block")
    add("Bash mklink 정션 만들기", B("cmd /c mklink /J lnk tests"), "block")

    # --- 6. PowerShell: 같은 규칙 ---
    add("PS Copy-Item expected", P("Copy-Item x.txt expected\\A_basic_expected.xlsx"), "block")
    add("PS Set-Content config", P("Set-Content -Path config\\conversion_map.csv -Value x"), "block")
    add("PS Out-File tests", P("'x' | Out-File tests\\test_compare.py"), "block")
    add("PS Remove-Item data\\fake", P("Remove-Item -Recurse -Force data\\fake"), "block")
    add("PS Rename-Item tests", P("Rename-Item tests\\test_compare.py x.py"), "block")
    add("PS Move-Item expected", P("Move-Item expected\\A_basic_expected.xlsx out\\"), "block")
    add("PS Junction 만들기", P("New-Item -ItemType Junction -Path src\\lnk -Target tests"), "block")
    add("PS SymbolicLink 만들기",
        P("New-Item -ItemType SymbolicLink -Path src\\lnk -Target tests"), "block")
    add("PS 경로를 변수에 담아 저장", P("$p = 'expected'; Set-Content \"$p\\a.csv\" x"), "block")
    add("PS 문자열 이어붙여 삭제", P("$p = 'exp' + 'ected'; Remove-Item \"$p\\a.csv\""), "block")

    # --- 7. git: 추적 상태를 바꾸는 명령은 차단 ---
    for n, cmd in (("git rm --cached", "git rm --cached tests/test_compare.py"),
                   ("git reset --hard", "git reset --hard"),
                   ("git restore", "git restore tests/test_compare.py"),
                   ("git checkout --", "git checkout -- tests/test_compare.py"),
                   ("git stash", "git stash"),
                   ("git config", "git config core.hooksPath x"),
                   ("git rebase", "git rebase main"),
                   ("git commit --amend", "git commit --amend -m x"),
                   ("git filter-branch", "git filter-branch --force HEAD")):
        add(f"Bash {n}", B(cmd), "block")
    add("PS git reset", P("git reset --hard"), "block")

    # --- 8. 허용되어야 하는 명령 ---
    for n, cmd in (("python -m unittest", "python -m unittest"),
                   ("python -m unittest 모듈", "python -m unittest tests.test_compare"),
                   ("훅 시험 실행", "python .claude/hooks/hook_test.py"),
                   ("도구 실행", "python -m trouble_merge --input data/fake --config config --out out"),
                   ("가짜 데이터 생성기", "python tools/make_fake_data.py"),
                   ("git status", "git status"),
                   ("git diff", "git diff"),
                   ("git log", "git log --oneline"),
                   ("git add src", "git add src/trouble_merge/merge.py"),
                   ("git commit", 'git commit -m "기능: 예시"'),
                   ("git switch -c", "git switch -c feature-x"),
                   ("ls tests", "ls tests"),
                   ("cat 보호 파일(읽기)", "cat expected/A_basic_expected.xlsx"),
                   ("out에 리다이렉트", "echo ok > out/log.txt")):
        add(f"Bash 허용: {n}", B(cmd), "allow")
    add("PS 허용: Get-Content 보호 파일(읽기)", P("Get-Content config\\conversion_map.csv"), "allow")
    add("PS 허용: python -m unittest", P("python -m unittest"), "allow")

    # --- 9. 판단 불가: 입력이 깨졌거나 환경을 모를 때 ---
    add("입력 JSON이 깨짐", ("Edit", {}), "unsure", raw_stdin="{깨진 json")
    add("입력이 비어 있음", ("Edit", {}), "unsure", raw_stdin="")
    add("Edit에 file_path가 없음", ("Edit", {"old_string": "a"}), "unsure")
    add("Bash에 command가 없음", ("Bash", {}), "unsure")
    add("git 실행 실패(보호 파일 수정 시도)", E(a("tests", "test_compare.py")), "unsure",
        env_set={"PATH": str(root / "nogit")})
    add("CLAUDE_PROJECT_DIR 없음", W(a("src", "trouble_merge", "new.py")), "unsure",
        env_drop=("CLAUDE_PROJECT_DIR",))
    return cases


def run_case(root: Path, case):
    env = os.environ.copy()
    env["CLAUDE_PROJECT_DIR"] = str(root)
    env["PYTHONIOENCODING"] = "utf-8"
    for k in case.get("env_drop", ()):
        env.pop(k, None)
    env.update(case.get("env_set", {}))
    if "raw_stdin" in case:
        data = case["raw_stdin"]
    else:
        data = json.dumps({"hook_event_name": "PreToolUse", "tool_name": case["tool"],
                           "tool_input": case["input"], "cwd": str(root)}, ensure_ascii=False)
    p = subprocess.run([sys.executable, str(HOOK)], input=data.encode("utf-8"),
                       capture_output=True, cwd=root, env=env, timeout=60)
    return p.returncode, p.stderr.decode("utf-8", "replace")


def judge(expect, code, err):
    if expect == "allow":
        return code == 0 and "차단" not in err
    if expect == "block":
        return code == 2 and "차단" in err
    if expect == "unsure":
        return code == 2 and "판단 불가" in err
    raise ValueError(expect)


def main() -> int:
    if not sys.stdout.isatty():  # 파이프로 받을 때 한글이 깨지지 않게
        sys.stdout.reconfigure(encoding="utf-8")
    if not HOOK.exists():
        print(f"[알림] 훅 스크립트가 없다: {HOOK}  (T005 전에는 전부 FAIL이 정상)")
    passed = failed = skipped = 0
    with tempfile.TemporaryDirectory(prefix="hooktest_") as tmp:
        root = Path(tmp).resolve() / "repo"
        root.mkdir()
        try:
            make_fixture(root)
            for case in build_cases(root):
                if case.get("skip"):
                    skipped += 1
                    print(f"SKIP {case['name']}  ({case['skip']})")
                    continue
                code, err = run_case(root, case)
                if judge(case["expect"], code, err):
                    passed += 1
                    print(f"PASS {case['name']}")
                else:
                    failed += 1
                    first = (err.strip().splitlines() or [""])[0][:100]
                    print(f"FAIL {case['name']}  (기대 {case['expect']}, 종료 코드 {code}, {first})")
        finally:
            link = root / "src_link"
            if link.exists():
                try:
                    os.rmdir(link)  # 정션만 지운다(대상 폴더 내용은 건드리지 않는다)
                except OSError:
                    pass
    print(f"\n결과: PASS {passed} / SKIP {skipped}")
    print(f"FAIL {failed}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
