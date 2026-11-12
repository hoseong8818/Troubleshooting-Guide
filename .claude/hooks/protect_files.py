#!/usr/bin/env python3
"""PreToolUse 보호 훅 (tasks T005, T005c 확장). 계약과 규칙은 hook_test.py 맨 위 설명을 따른다.

- 차단: 종료 코드 2 + 표준오류 `차단: 이유`
- 판단 불가: 종료 코드 2 + 표준오류 `차단(판단 불가): 이유` (fail-closed)
- 허용: 종료 코드 0
- 어떤 예외가 나도 종료 코드 2로 끝낸다(미처리 예외는 종료 코드 1이라 차단되지 않기 때문).
- 이유 문구에는 파일 경로와 명령 종류만 넣고 명령 본문이나 파일 내용은 넣지 않는다.
"""
import json
import os
import re
import subprocess
import sys


class Block(Exception):
    pass


class Unsure(Exception):
    pass


FILE_TOOLS = {"Edit", "Write", "MultiEdit", "NotebookEdit"}
SHELL_TOOLS = {"Bash", "PowerShell"}
TASKS_REL = "specs/001-trouble-merge/tasks.md"
SPEC_NAMES = {"spec.md", "plan.md", "tasks.md"}
WORKTREES_PREFIX = ".claude/worktrees/"

# ---------------------------------------------------------------- worktree


def valid_worktree(root, wt_dir):
    """wt_dir가 root의 `.git/worktrees/` 아래에 등록된 실제 git worktree인가."""
    gitfile = os.path.join(wt_dir, ".git")
    if not os.path.isfile(gitfile):
        return False
    with open(gitfile, encoding="utf-8") as f:
        first = f.readline().strip()
    if not first.lower().startswith("gitdir:"):
        return False
    gitdir = first[7:].strip()
    if not os.path.isabs(gitdir):
        gitdir = os.path.join(wt_dir, gitdir)
    gitdir = os.path.normcase(os.path.realpath(gitdir))
    base = os.path.normcase(os.path.realpath(os.path.join(root, ".git", "worktrees")))
    return gitdir.startswith(base + os.sep)


# ---------------------------------------------------------------- 파일 도구


def protected_kind(rel):
    """rel: 저장소 기준, 소문자, 슬래시. 보호 대상이면 종류를, 아니면 None."""
    parts = rel.split("/")
    if parts[0] == ".claude":
        return "claude"
    if parts[0] in ("expected", "tests") and len(parts) > 1:
        return "dir"
    if rel.startswith("data/fake/"):
        return "dir"
    if parts[0] == "config" and rel.endswith(".csv"):
        return "dir"
    if rel in ("docs/constitution.md", "claude.md"):
        return "doc"
    if parts[0] == "specs" and parts[-1] in SPEC_NAMES:
        return "doc"
    return None


def clean_windows_path(p):
    """\\\\?\\ 접두사, :스트림, 후행 점·공백을 걷어 낸다."""
    if p.startswith("\\\\?\\") or p.startswith("//?/"):
        p = p[4:]
    drive, rest = os.path.splitdrive(p)
    out = []
    for c in re.split(r"[\\/]", rest):
        if c in ("", ".", ".."):
            out.append(c)
            continue
        c = c.split(":")[0].rstrip(". ")
        out.append(c)
    return drive + os.sep.join(out)


def resolve_rel(root, cwd, raw):
    """실제 경로(realpath)로 해석한 저장소 기준 상대 경로. 저장소 밖이면 None."""
    p = clean_windows_path(raw) if os.name == "nt" else raw
    if not os.path.isabs(p):
        p = os.path.join(cwd, p)
    real = os.path.realpath(p)
    try:
        rel = os.path.relpath(real, root)
    except ValueError:  # 다른 드라이브
        return None
    if rel == ".." or rel.startswith(".." + os.sep):
        return None
    return rel.replace(os.sep, "/")


def is_tracked(root, rel):
    r = subprocess.run(["git", "ls-files", "--error-unmatch", "--", rel],
                       cwd=root, capture_output=True, timeout=30)
    if r.returncode == 0:
        return True  # 커밋했거나 git add만 한 파일
    if r.returncode == 1:
        return False
    raise Unsure("git 상태를 알 수 없다")


def listed_in_tasks(root, rel_lower):
    try:
        with open(os.path.join(root, *TASKS_REL.split("/")), encoding="utf-8") as f:
            text = f.read().lower()
    except FileNotFoundError:
        return False
    pat = r"(?<![\w/.\-])" + re.escape(rel_lower) + r"(?![\w/\-]|\.\w)"
    return re.search(pat, text) is not None


def decide_file(root, cwd, raw):
    rel = resolve_rel(root, cwd, raw)
    if rel is None:
        return  # 저장소 밖: 이 훅의 대상이 아니다(실제 폴더는 deny가 막는다)
    work_root = root
    if rel.lower().startswith(WORKTREES_PREFIX):
        parts = rel.split("/")  # ['.claude', 'worktrees', '<이름>', ...]
        if len(parts) >= 4:
            wt_dir = os.path.join(root, ".claude", "worktrees", parts[2])
            if valid_worktree(root, wt_dir):
                # 실제 worktree: 그 폴더를 저장소 루트로 보고 같은 규칙을 적용한다.
                work_root = os.path.realpath(wt_dir)
                rel = "/".join(parts[3:])
        # 실제 worktree가 아니거나 worktree 폴더 자체이면 rel이 .claude 아래로 남아 차단된다.
    low = rel.lower()
    if low.split("/")[0] == ".git":
        raise Block("저장소 내부(.git) 수정: " + rel)
    kind = protected_kind(low)
    if kind is None:
        return
    if kind == "claude":
        raise Block("보호 장치(.claude)는 사람이 만든다: " + rel)
    if is_tracked(work_root, rel):
        raise Block("git에 추적되는 보호 대상 수정: " + rel)
    if listed_in_tasks(work_root, low):
        return
    raise Block("tasks.md에 이름이 없는 보호 폴더의 새 파일: " + rel)


# ------------------------------------------------------------ Bash·PowerShell

INTERP = {"python", "python3", "py", "pythonw", "powershell", "pwsh", "cmd", "bash", "sh",
          "node", "uv", "uvx", "perl", "ruby", "wsl", "mshta", "wscript", "cscript"}
PYTHON = {"python", "python3", "py", "pythonw"}
WRAPPERS = {"env", "command", "exec", "time", "nohup", "sudo", "start", "call"}
ALLOWED_PY_ARGS = [
    re.compile(r"^-m\s+unittest(\s|$)"),
    re.compile(r"^(\./)?\.claude/hooks/hook_test\.py(\s|$)"),
    re.compile(r"^-m\s+trouble_merge(\s|$)"),
    re.compile(r"^(\./)?tools/make_fake_data\.py(\s|$)"),
    re.compile(r"^(\./)?tools/make_sample_data\.py(\s|$)"),
]
SCRIPT_EXT_RE = re.compile(r"\.(ps1|sh|bat|cmd|vbs|js|py|pl|rb)$")
GIT_BRANCH_READ_OPTS = {"--show-current", "-a", "--all", "-r", "--remotes", "--list", "-v", "-vv",
                        "--no-color"}

WRITE_VERBS = (r"cp|copy|copy-item|cpi|mv|move|move-item|mi|rm|del|erase|rd|rmdir|ri|remove-item|"
               r"ren|rename|rename-item|rni|set-content|sc|add-content|ac|out-file|clear-content|"
               r"new-item|ni|tee|tee-object|touch|truncate|xcopy|robocopy|dd|ln|curl|wget|certutil|"
               r"invoke-webrequest|iwr|expand-archive|compress-archive|set-itemproperty")
VERB_RE = re.compile(r"(?<![\w.\-])(?:" + WRITE_VERBS + r")(?![\w\-])")
SED_RE = re.compile(r"(?<![\w.\-])sed\s+(?:-\w*i\b|--in-place)")
REDIR_RE = re.compile(r"(?<![<>\-=])>{1,2}(?!&)")
# /dev/null, nul, $null 로 가는 리다이렉트는 파일을 쓰지 않는다.
NULL_REDIR_RE = re.compile(r"(?:\d?|&)>{1,2}\s*(?:/dev/null|nul|\$null)(?![\w./\-])")
LINK_RE = re.compile(r"(?<![\w.\-])mklink(?![\w])|-itemtype\s+(?:junction|symboliclink|hardlink)"
                     r"|(?<![\w.\-])ln\s+-\w*s|fsutil\s+hardlink")
EXEC_RE = re.compile(r"(?<![\w.\-])(?:iex|invoke-expression|start-process|saps|invoke-command)(?![\w\-])")
UNRESOLVED_RE = re.compile(r"\$[\w{(]|%\w+%|`|['\"]\s*\+\s*['\"]")
PROT_CMD_RE = re.compile(
    r"(?<![\w.\-])(?:\./)?(?:expected|data/fake|tests|\.claude|\.git)(?![\w.\-])"
    r"|(?<![\w.\-])(?:\./)?config/[^\s\"']*\.csv"
    r"|(?<![\w.\-])claude\.md(?![\w])"
    r"|docs/constitution\.md"
    r"|specs/[^\s\"']*(?:spec|plan|tasks)\.md")
WT_RE = re.compile(r"\.claude/worktrees/([^/\s\"']+)((?:/[^\s\"']*)?)")
HEREDOC_RE = re.compile(r"<<-?\s*(['\"]?)(\w+)\1[^\n]*\n.*?\n[ \t]*\2[ \t]*(?:\n|$)", re.S)
HERESTRING_RE = re.compile(r"@(['\"]).*?\1@", re.S)
TOKEN_RE = re.compile(r"\"[^\"]*\"|'[^']*'|\S+")


def norm(s):
    return s.replace("\\", "/").lower()


def strip_worktrees(t, root):
    """명령 안의 `.claude/worktrees/<이름>/`을, 실제 worktree일 때만 걷어 낸다.
    이름이 `..`이거나 뒤따르는 경로에 `..`가 끼면 걷어 내지 않는다(그러면 .claude가 남아 차단된다)."""
    def repl(m):
        name, rest = m.group(1), m.group(2)
        if name in (".", "..") or ".." in rest.split("/"):
            return m.group(0)
        if not valid_worktree(root, os.path.join(root, ".claude", "worktrees", name)):
            return m.group(0)
        return rest[1:]
    return WT_RE.sub(repl, t)


def no_null(t):
    return NULL_REDIR_RE.sub(" ", t)


def strip_wrappers(stmt):
    s = stmt.strip()
    while True:
        m = re.match(r"^(?:&\s*|\w+=\S*\s+)", s)
        if m:
            s = s[m.end():]
            continue
        m = re.match(r"^(\w+)\s+", s)
        if m and m.group(1).lower() in WRAPPERS:
            s = s[m.end():]
            continue
        return s


def first_word(stmt):
    s = strip_wrappers(stmt)
    m = TOKEN_RE.match(s)
    tok = m.group(0).strip("\"'") if m else ""
    base = re.split(r"[\\/]", tok)[-1].lower()
    return re.sub(r"\.(exe|cmd|bat)$", "", base), tok, s[m.end():] if m else ""


def check_git(stmt):
    toks = [t.strip("\"'") for t in TOKEN_RE.findall(strip_wrappers(stmt))][1:]
    sub, i = None, 0
    while i < len(toks):
        t = toks[i]
        if t in ("-c", "--exec-path") or t.startswith(("--git-dir", "--work-tree", "-c=")):
            raise Block("git 전역 옵션(-c, --git-dir, --work-tree)은 허용하지 않는다")
        if t == "-C":
            i += 2
            continue
        if t.startswith("-"):
            i += 1
            continue
        sub = t
        break
    if sub is None:
        return
    rest = toks[i + 1:]
    if sub in ("status", "diff", "log", "show"):
        if any(t.startswith("--output") for t in rest):
            raise Block("git " + sub + " --output은 허용하지 않는다")
        return
    if sub in ("add", "rev-parse", "ls-files"):
        return
    if sub == "commit":
        if any(t.startswith("--amend") for t in rest):
            raise Block("git commit --amend는 허용하지 않는다")
        return
    if sub == "switch":
        if any(t in ("-c", "-C", "--create", "--force-create") for t in rest):
            return
        raise Block("git switch는 -c(새 브랜치)만 허용한다")
    if sub == "branch":
        if all(t in GIT_BRANCH_READ_OPTS for t in rest):
            return
        raise Block("git branch는 목록 조회 옵션만 허용한다")
    if sub == "worktree":
        if rest[:1] == ["list"]:
            return
        raise Block("git worktree는 list만 허용한다")
    raise Block("허용 목록 밖 git 명령: git " + sub)


def decide_shell(cmd, root):
    cmd = HEREDOC_RE.sub(" ", cmd)
    cmd = HERESTRING_RE.sub(" ", cmd)
    generic = []
    for stmt in (s for s in re.split(r"&&|\|\||[;|\n]", cmd) if s.strip()):
        n = strip_worktrees(norm(stmt), root)
        if LINK_RE.search(n):
            raise Block("링크·정션 생성 명령은 허용하지 않는다")
        if EXEC_RE.search(n):
            raise Block("실행 대상을 알 수 없는 명령(iex, Start-Process 등)")
        word, tok, args = first_word(stmt)
        if word == "git":
            check_git(stmt)
            nn = no_null(n)
            if REDIR_RE.search(nn) and PROT_CMD_RE.search(nn):
                raise Block("보호 경로로 리다이렉트하는 git 명령")
            continue
        if word in INTERP:
            if word not in PYTHON:
                raise Block("허용 목록 밖 인터프리터: " + word)
            a = args.strip().replace("\\", "/")
            if not any(p.match(a) for p in ALLOWED_PY_ARGS):
                raise Block("허용 목록 밖 python 실행(-c, 임의 스크립트 등)")
            m = re.search(r"--out(?:=|\s+)(\S+)", a)
            if m and PROT_CMD_RE.search(strip_worktrees(norm(m.group(1)), root)):
                raise Block("도구 출력 폴더가 보호 경로다")
        elif SCRIPT_EXT_RE.search(tok.lower()):
            raise Block("스크립트 파일 직접 실행은 허용하지 않는다")
        generic.append(stmt)
    t = no_null(strip_worktrees(norm(" ; ".join(generic)), root))
    if VERB_RE.search(t) or SED_RE.search(t) or REDIR_RE.search(t):
        if PROT_CMD_RE.search(t):
            raise Block("보호 경로에 쓰거나 지우거나 옮기는 명령")
        if UNRESOLVED_RE.search(t):
            raise Unsure("쓰기 명령의 대상이 변수나 이어붙인 문자열이라 해석할 수 없다")


# ------------------------------------------------------------------- 진입점


def decide():
    data = json.loads(sys.stdin.buffer.read().decode("utf-8"))
    if not isinstance(data, dict):
        raise Unsure("입력 형식을 알 수 없다")
    tool = data.get("tool_name")
    if not isinstance(tool, str):
        raise Unsure("tool_name이 없다")
    if tool not in FILE_TOOLS and tool not in SHELL_TOOLS:
        return
    env_root = os.environ.get("CLAUDE_PROJECT_DIR")
    if not env_root:
        raise Unsure("CLAUDE_PROJECT_DIR을 모른다")
    root = os.path.realpath(env_root)
    tin = data.get("tool_input")
    if not isinstance(tin, dict):
        raise Unsure("tool_input이 없다")
    if tool in FILE_TOOLS:
        raw = tin.get("file_path") or tin.get("notebook_path")
        if not isinstance(raw, str) or not raw:
            raise Unsure("대상 경로가 없다")
        cwd = data.get("cwd") if isinstance(data.get("cwd"), str) and data.get("cwd") else root
        decide_file(root, cwd, raw)
    else:
        cmd = tin.get("command")
        if not isinstance(cmd, str) or not cmd.strip():
            raise Unsure("command가 없다")
        decide_shell(cmd, root)


def emit(prefix, msg):
    sys.stderr.buffer.write((prefix + " " + msg + "\n").encode("utf-8"))
    sys.stderr.buffer.flush()


def main():
    code = 0
    try:
        decide()
    except Block as e:
        emit("차단:", str(e))
        code = 2
    except Unsure as e:
        emit("차단(판단 불가):", str(e))
        code = 2
    except Exception as e:  # 모든 실패는 차단(fail-closed). 예외 메시지는 버리고 종류만 낸다.
        emit("차단(판단 불가):", "훅 내부 예외 " + type(e).__name__)
        code = 2
    sys.exit(code)


if __name__ == "__main__":
    main()
