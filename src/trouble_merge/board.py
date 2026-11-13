"""통합 엑셀을 QnA 게시판 HTML 한 장으로 만든다 (T054, 데모).

외부 서비스·LLM 없이 브라우저 안에서만 검색한다(헌법 원칙 1). 게시글 데이터는 JSON으로 넣고
화면에는 textContent로만 그려서 셀 값이 코드로 실행되지 않게 한다. 이미 있는 파일은 덮어쓰지 않는다.
사용: python -m trouble_merge.board --merged <통합 xlsx> --out <폴더>
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from .writer import unique_path

PAGE = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>트러블슈팅 QnA 게시판</title>
<style>
:root{--bg:#f6f7f9;--card:#fff;--text:#1d2530;--muted:#667085;--line:#e3e7ee;--accent:#2457d6;--accent-bg:#e8eefc;--warn:#b54708;--warn-bg:#fff3e0;--dup:#6941c6;--dup-bg:#f1ebfd}
@media (prefers-color-scheme:dark){:root{--bg:#14171c;--card:#1d222a;--text:#e7eaf0;--muted:#98a2b3;--line:#2e3540;--accent:#8fb0ff;--accent-bg:#26324f;--warn:#ffb866;--warn-bg:#3a2a14;--dup:#c4b0ff;--dup-bg:#2c2545}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:15px/1.55 "Malgun Gothic","Apple SD Gothic Neo",system-ui,sans-serif}
header{padding:20px 16px 8px;max-width:1100px;margin:0 auto}
h1{margin:0 0 4px;font-size:22px}
header p{margin:0;color:var(--muted);font-size:13px}
main{max-width:1100px;margin:0 auto;padding:8px 16px 40px}
#q{width:100%;padding:12px 14px;font-size:16px;border:1px solid var(--line);border-radius:10px;background:var(--card);color:var(--text)}
#q:focus{outline:2px solid var(--accent);outline-offset:1px}
#cats{display:flex;flex-wrap:wrap;gap:6px;margin:12px 0}
.chip{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:999px;padding:4px 12px;font-size:13px;cursor:pointer}
.chip[aria-pressed="true"]{background:var(--accent-bg);border-color:var(--accent);color:var(--accent);font-weight:600}
#count{color:var(--muted);font-size:13px;margin:4px 0 8px}
.cols{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.1fr);gap:16px;align-items:start}
@media (max-width:760px){.cols{grid-template-columns:1fr}}
#list{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:8px}
.item{width:100%;text-align:left;background:var(--card);color:var(--text);border:1px solid var(--line);border-radius:10px;padding:10px 12px;cursor:pointer;font:inherit}
.item:hover{border-color:var(--accent)}
.item[aria-current="true"]{border-color:var(--accent);box-shadow:0 0 0 2px var(--accent-bg)}
.title{font-weight:600}
.meta{color:var(--muted);font-size:12px;margin-top:2px}
.badge{display:inline-block;font-size:11px;padding:1px 7px;border-radius:999px;margin-left:6px;vertical-align:middle}
.badge.warn{background:var(--warn-bg);color:var(--warn)}
.badge.dup{background:var(--dup-bg);color:var(--dup)}
#detail{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;position:sticky;top:12px}
#detail h2{margin:0 0 4px;font-size:18px}
#detail dt{margin-top:14px;font-size:12px;font-weight:700;color:var(--muted)}
#detail dd{margin:2px 0 0;white-space:pre-wrap;overflow-wrap:anywhere}
#detail .answer{background:var(--accent-bg);border-radius:8px;padding:10px 12px}
#hint,#empty{color:var(--muted);padding:24px 8px;text-align:center}
.notice{margin-top:10px;padding:8px 10px;border-radius:8px;background:var(--warn-bg);color:var(--warn);font-size:13px}
</style>
</head>
<body>
<header>
<h1>트러블슈팅 QnA 게시판</h1>
<p>증상이나 원인을 입력하면 비슷한 사례와 조치를 보여 줍니다. 이 화면은 이 파일 안의 사례만 찾습니다(외부 전송 없음).</p>
</header>
<main>
<input id="q" type="search" placeholder="예: 베어링 소음, 압력 저하, PMP01" aria-label="증상·원인·설비 검색" autocomplete="off">
<div id="cats" role="group" aria-label="분류"></div>
<div id="count" aria-live="polite"></div>
<div class="cols">
<ul id="list"></ul>
<section id="detail" aria-live="polite"><p id="hint">왼쪽 목록에서 사례를 고르면 원인과 조치가 여기에 나옵니다.</p></section>
</div>
<p id="empty" hidden>해당 사례 없음. 다른 단어로 검색하거나 분류를 바꿔 보세요.</p>
</main>
<script type="application/json" id="board-data">__DATA__</script>
<script>
(function () {
  "use strict";
  var posts = JSON.parse(document.getElementById("board-data").textContent);
  var qEl = document.getElementById("q"), catsEl = document.getElementById("cats");
  var listEl = document.getElementById("list"), detailEl = document.getElementById("detail");
  var countEl = document.getElementById("count"), emptyEl = document.getElementById("empty");
  var NEED = "확인 필요", cat = "", selected = -1, PAGE_SIZE = 100, limit = PAGE_SIZE;
  var WEIGHT = {"증상": 3, "원인": 3, "조치": 1, "설비": 2, "분류": 1, "기타 항목": 1};

  function text(v) { return v === null || v === undefined ? "" : String(v); }
  function el(tag, cls, str) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (str !== undefined) e.textContent = str;
    return e;
  }
  function categories() {
    var seen = {}, order = [];
    posts.forEach(function (p) {
      var c = text(p["분류"]) || "(분류 없음)";
      if (!(c in seen)) { seen[c] = 0; order.push(c); }
      seen[c] += 1;
    });
    return {seen: seen, order: order};
  }
  function inCat(p) { return !cat || (text(p["분류"]) || "(분류 없음)") === cat; }
  function score(p, tokens) {
    var total = 0;
    for (var i = 0; i < tokens.length; i++) {
      var hit = 0;
      for (var f in WEIGHT) {
        if (text(p[f]).toLowerCase().indexOf(tokens[i]) !== -1) hit += WEIGHT[f];
      }
      if (!hit) return 0;
      total += hit;
    }
    return total;
  }
  function renderCats() {
    var info = categories();
    catsEl.textContent = "";
    function chip(label, value, n) {
      var b = el("button", "chip", label + " (" + n + ")");
      b.type = "button";
      b.setAttribute("aria-pressed", String(cat === value));
      b.addEventListener("click", function () { cat = value; selected = -1; limit = PAGE_SIZE; render(); });
      catsEl.appendChild(b);
    }
    chip("전체", "", posts.length);
    info.order.forEach(function (c) { chip(c, c, info.seen[c]); });
  }
  function badges(p, parent) {
    if (p["확인 필요 이유"]) parent.appendChild(el("span", "badge warn", NEED));
    if (p["중복 의심"]) parent.appendChild(el("span", "badge dup", "중복 의심"));
  }
  function showDetail(index) {
    var p = posts[index];
    selected = index;
    detailEl.textContent = "";
    var h = el("h2", null, text(p["증상"]) || "(증상 없음)");
    badges(p, h);
    detailEl.appendChild(h);
    detailEl.appendChild(el("div", "meta", [p["설비"], p["분류"], p["일자"]].map(text).filter(Boolean).join(" · ")));
    var dl = el("dl");
    function row(label, value, cls) {
      dl.appendChild(el("dt", null, label));
      dl.appendChild(el("dd", cls, text(value) || "(내용 없음)"));
    }
    row("원인", p["원인"]);
    row("조치(해결 방법)", p["조치"], "answer");
    if (p["기타 항목"]) row("기타 항목", p["기타 항목"]);
    row("출처", text(p["원본 파일"]) + " / " + text(p["원본 시트"]) + " / " + text(p["원본 행"]) + "행");
    detailEl.appendChild(dl);
    if (p["확인 필요 이유"]) {
      detailEl.appendChild(el("div", "notice", NEED + ": " + text(p["확인 필요 이유"]) +
        (p["원래 값"] ? " (원래 값: " + text(p["원래 값"]) + ")" : "")));
    }
    if (p["중복 의심"]) detailEl.appendChild(el("div", "notice", text(p["중복 의심"])));
    Array.prototype.forEach.call(listEl.querySelectorAll(".item"), function (b) {
      b.setAttribute("aria-current", String(Number(b.dataset.index) === index));
    });
  }
  function render() {
    renderCats();
    var tokens = qEl.value.toLowerCase().split(/\s+/).filter(Boolean);
    var found = [];
    posts.forEach(function (p, i) {
      if (!inCat(p)) return;
      var s = tokens.length ? score(p, tokens) : 1;
      if (s) found.push({i: i, s: s});
    });
    if (tokens.length) found.sort(function (a, b) { return b.s - a.s || a.i - b.i; });
    listEl.textContent = "";
    found.slice(0, limit).forEach(function (f) {
      var p = posts[f.i];
      var li = el("li");
      var b = el("button", "item");
      b.type = "button";
      b.dataset.index = String(f.i);
      var t = el("div", "title", text(p["증상"]) || "(증상 없음)");
      badges(p, t);
      b.appendChild(t);
      b.appendChild(el("div", "meta", [p["설비"], p["분류"], p["일자"]].map(text).filter(Boolean).join(" · ")));
      b.addEventListener("click", function () { showDetail(f.i); });
      li.appendChild(b);
      listEl.appendChild(li);
    });
    if (found.length > limit) {
      var moreLi = el("li");
      var more = el("button", "chip", "더 보기 (" + (found.length - limit) + "건 남음)");
      more.type = "button";
      more.addEventListener("click", function () { limit += PAGE_SIZE; render(); });
      moreLi.appendChild(more);
      listEl.appendChild(moreLi);
    }
    countEl.textContent = found.length + "건" + (found.length > limit ? " (" + limit + "건 표시)" : "");
    emptyEl.hidden = found.length !== 0;
    if (!found.length || !found.some(function (f) { return f.i === selected; })) {
      selected = -1;
      detailEl.textContent = "";
      var hint = el("p", null, "왼쪽 목록에서 사례를 고르면 원인과 조치가 여기에 나옵니다.");
      hint.id = "hint";
      detailEl.appendChild(hint);
    } else {
      showDetail(selected);
    }
  }
  qEl.addEventListener("input", function () { limit = PAGE_SIZE; render(); });
  render();
})();
</script>
</body>
</html>
"""


def read_posts(path):
    """통합 엑셀의 첫 시트를 게시글 목록(헤더 이름 -> 값)으로 읽는다. 빈 문자열은 None."""
    wb = load_workbook(path, data_only=True)
    try:
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    header = list(rows[0])
    return [
        dict(zip(header, [None if v == "" else v for v in row]))
        for row in rows[1:] if any(v is not None for v in row)
    ]


def to_json(posts):
    """</script>로 데이터 블록을 닫지 못하도록 <, >, &와 줄 구분자를 이스케이프한다."""
    raw = json.dumps(posts, ensure_ascii=False, default=str)
    for char, escaped in (("<", "\\u003c"), (">", "\\u003e"), ("&", "\\u0026"),
                          (" ", "\\u2028"), (" ", "\\u2029")):
        raw = raw.replace(char, escaped)
    return raw


def build_board(merged, out_dir, stamp):
    posts = read_posts(merged)
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    path = unique_path(out_dir, "board", ".html", stamp)
    path.write_text(PAGE.replace("__DATA__", to_json(posts)), encoding="utf-8")
    return path, len(posts)


def main(argv=None):
    parser = argparse.ArgumentParser(prog="trouble_merge.board")
    parser.add_argument("--merged", required=True, help="통합 엑셀(merged.xlsx)")
    parser.add_argument("--out", required=True, help="게시판 HTML을 만들 폴더")
    args = parser.parse_args(argv)
    try:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        _, count = build_board(args.merged, args.out, stamp)
    except Exception:  # 메시지와 종류를 버리고 고정 문구만 낸다(셀 값이 섞일 수 있다)
        print("게시판을 만들 수 없음: 통합 엑셀을 열 수 없다", file=sys.stderr)
        return 2
    print("완료: 게시글 %d건" % count)
    return 0


if __name__ == "__main__":
    sys.exit(main())
