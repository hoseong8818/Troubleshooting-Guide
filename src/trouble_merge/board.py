"""통합 엑셀을 블로그형 QnA 게시판 HTML 한 장으로 만든다 (T054, 데모).

외부 서비스·LLM 없이 브라우저 안에서만 검색한다(헌법 원칙 1). 게시글 데이터는 JSON으로 넣고
화면에는 textContent로만 그려서 셀 값이 코드로 실행되지 않게 한다. 외부 폰트·이미지·스크립트를 쓰지 않는다.
이미 있는 파일은 덮어쓰지 않는다.
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
<html lang="ko" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>설비 트러블슈팅 노트</title>
<style>
:root{--bg:#ffffff;--surface:#f6f7fb;--text:#1b1d26;--muted:#6b7080;--line:#e6e8ef;--accent:#e23a5e;--accent-soft:#fdebef;--answer:#eef4ff;--answer-line:#3d6fe0;--warn:#b54708;--warn-bg:#fff3e0;--dup:#6941c6;--dup-bg:#f1ebfd}
[data-theme="dark"]{--bg:#14161f;--surface:#1c1f2b;--text:#e9eaf0;--muted:#9a9db0;--line:#2a2e3d;--accent:#ff5c7c;--accent-soft:#3a1d27;--answer:#1d2740;--answer-line:#7fa2ff;--warn:#ffb866;--warn-bg:#3a2a14;--dup:#c4b0ff;--dup-bg:#2c2545}
*{box-sizing:border-box}
[hidden]{display:none!important}
html,body{margin:0;background:var(--bg);color:var(--text)}
body{font:16px/1.7 "Pretendard","Malgun Gothic","Apple SD Gothic Neo",system-ui,sans-serif;transition:background .2s,color .2s}
button{font:inherit;color:inherit}
a{color:inherit;text-decoration:none}
.top{position:sticky;top:0;z-index:5;background:var(--bg);border-bottom:1px solid var(--line)}
.top-in{max-width:1040px;margin:0 auto;padding:12px 16px;display:flex;align-items:center;gap:12px}
.logo{font-weight:800;font-size:18px;letter-spacing:-.3px;cursor:pointer;background:none;border:0;padding:0}
.logo span{margin-right:6px}
.spacer{flex:1}
.nav{font-size:14px;font-weight:600;background:none;border:0;cursor:pointer;padding:6px 8px}
.icon{width:36px;height:36px;border-radius:8px;border:0;background:none;cursor:pointer;font-size:17px;display:grid;place-items:center}
.icon:hover,.nav:hover{background:var(--surface)}
.searchbar{max-width:1040px;margin:0 auto;padding:0 16px 14px}
#q{width:100%;padding:12px 14px;font-size:16px;border:1px solid var(--line);border-radius:10px;background:var(--surface);color:var(--text)}
#q:focus{outline:2px solid var(--accent);outline-offset:1px}
.wrap{max-width:1040px;margin:0 auto;padding:28px 16px 64px}
.layout{display:grid;grid-template-columns:200px minmax(0,1fr);gap:40px}
.cats{min-width:0}
.cats h2{margin:0 0 10px;font-size:14px;color:var(--accent)}
.cats ul{list-style:none;margin:0;padding:0}
.cat{display:block;width:100%;text-align:left;background:none;border:0;padding:4px 8px;margin-left:-8px;border-radius:6px;font-size:13px;font-weight:600;color:var(--muted);cursor:pointer}
.cat:hover{color:var(--text)}
.cat[aria-pressed="true"]{color:var(--accent);background:var(--accent-soft)}
#count{font-size:13px;color:var(--muted);margin:0 0 8px}
#list{list-style:none;margin:0;padding:0}
.post{padding:20px 0;border-bottom:1px solid var(--line)}
.post:first-child{padding-top:0}
.date{font-size:13px;color:var(--muted)}
.ptitle{display:block;margin:2px 0 4px;font-size:20px;font-weight:800;line-height:1.4;letter-spacing:-.3px;background:none;border:0;padding:0;text-align:left;cursor:pointer}
.ptitle:hover{color:var(--accent)}
.tags{display:flex;flex-wrap:wrap;gap:4px 10px;font-size:12px;font-weight:700}
.tag{color:var(--accent)}
.badge{font-size:11px;padding:0 7px;border-radius:999px;font-weight:700}
.badge.warn{background:var(--warn-bg);color:var(--warn)}
.badge.dup{background:var(--dup-bg);color:var(--dup)}
.summary{margin:6px 0 0;font-size:14px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.more{margin-top:20px;width:100%;padding:10px;border:1px solid var(--line);border-radius:10px;background:var(--surface);cursor:pointer;font-weight:600}
#empty{padding:48px 8px;text-align:center;color:var(--muted)}
#article{max-width:760px;margin:0 auto}
.back{background:none;border:0;padding:0;color:var(--muted);cursor:pointer;font-size:14px}
.back:hover{color:var(--accent)}
.ahead{text-align:center;padding:24px 0 28px;border-bottom:1px solid var(--line);margin-bottom:8px}
.ahead h1{margin:6px 0 10px;font-size:32px;line-height:1.35;letter-spacing:-.6px;overflow-wrap:anywhere}
.ahead .tags{justify-content:center}
#article h2{margin:36px 0 10px;font-size:22px;letter-spacing:-.3px}
#article h2::before{content:"";display:inline-block;width:4px;height:18px;background:var(--accent);border-radius:2px;margin-right:10px;vertical-align:-2px}
#article p{margin:0;white-space:pre-wrap;overflow-wrap:anywhere}
.answer{background:var(--answer);border-left:4px solid var(--answer-line);border-radius:8px;padding:14px 16px;font-weight:600}
.notice{margin-top:24px;padding:12px 14px;border-radius:8px;background:var(--warn-bg);color:var(--warn);font-size:14px}
.source{font-size:14px;color:var(--muted)}
.related{list-style:none;margin:0;padding:0}
.related li{padding:8px 0;border-bottom:1px dashed var(--line)}
.related button{background:none;border:0;padding:0;cursor:pointer;text-align:left;font-weight:700}
.related button:hover{color:var(--accent)}
.related small{display:block;color:var(--muted);font-weight:400}
@media (max-width:760px){
  .layout{grid-template-columns:1fr;gap:16px}
  .cats ul{display:flex;gap:6px;overflow-x:auto;padding-bottom:4px}
  .cat{white-space:nowrap;margin-left:0;border:1px solid var(--line);border-radius:999px;padding:4px 12px}
  .ahead h1{font-size:24px}
}
</style>
</head>
<body>
<header class="top">
<div class="top-in">
<button class="logo" id="home" type="button"><span aria-hidden="true">🛠️</span>설비 트러블슈팅 노트</button>
<div class="spacer"></div>
<button class="nav" id="all" type="button">전체 글</button>
<button class="icon" id="find" type="button" aria-label="검색으로 이동">🔍</button>
<button class="icon" id="theme" type="button" aria-label="밝기 전환">☾</button>
</div>
<div class="searchbar" id="searchbar">
<input id="q" type="search" placeholder="증상이나 원인을 물어보세요. 예: 베어링 소음, 압력 저하, PMP01" aria-label="증상·원인·설비 검색" autocomplete="off">
</div>
</header>
<div class="wrap">
<div class="layout" id="listview">
<nav class="cats" aria-label="분류">
<h2>분류</h2>
<ul id="cats"></ul>
</nav>
<section>
<p id="count" aria-live="polite"></p>
<ul id="list"></ul>
<p id="empty" hidden>해당 사례 없음. 다른 단어로 검색하거나 분류를 바꿔 보세요.</p>
</section>
</div>
<article id="article" hidden></article>
</div>
<script type="application/json" id="board-data">__DATA__</script>
<script>
(function () {
  "use strict";
  var posts = JSON.parse(document.getElementById("board-data").textContent);
  var $ = function (id) { return document.getElementById(id); };
  var qEl = $("q"), catsEl = $("cats"), listEl = $("list"), countEl = $("count"), emptyEl = $("empty");
  var listView = $("listview"), articleEl = $("article"), searchbar = $("searchbar");
  var NEED = "확인 필요", NO_CAT = "(분류 없음)", PAGE_SIZE = 50;
  var cat = "", limit = PAGE_SIZE;
  var WEIGHT = {"증상": 3, "원인": 3, "조치": 1, "설비": 2, "분류": 1, "기타 항목": 1};

  function text(v) { return v === null || v === undefined ? "" : String(v); }
  function el(tag, cls, str) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (str !== undefined) e.textContent = str;
    return e;
  }
  function catOf(p) { return text(p["분류"]) || NO_CAT; }
  function titleOf(p) { return text(p["증상"]) || "(증상 없음)"; }
  var DATE_RE = /^(\\d{4})-(\\d{2})-(\\d{2})$/;
  function dateKey(p) { var s = text(p["일자"]); return DATE_RE.test(s) ? s : ""; }
  function prettyDate(v) {
    var s = text(v), m = DATE_RE.exec(s);
    if (m) return m[1] + "년 " + Number(m[2]) + "월 " + Number(m[3]) + "일";
    if (s === NEED) return "일자 확인 필요";
    return s || "일자 미상";
  }
  function tagsOf(p) {
    var box = el("div", "tags");
    if (p["분류"]) box.appendChild(el("span", "tag", text(p["분류"])));
    if (p["설비"]) box.appendChild(el("span", "tag", text(p["설비"])));
    if (p["확인 필요 이유"]) box.appendChild(el("span", "badge warn", NEED));
    if (p["중복 의심"]) box.appendChild(el("span", "badge dup", "중복 의심"));
    return box;
  }
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
  function byDateDesc(a, b) {
    var da = dateKey(posts[a]), db = dateKey(posts[b]);  // 해석된 날짜만 앞에, 나머지는 뒤에
    if (da !== db) { if (!da) return 1; if (!db) return -1; return da < db ? 1 : -1; }
    return a - b;
  }

  function renderCats() {
    var counts = {}, order = [];
    posts.forEach(function (p) {
      var c = catOf(p);
      if (!(c in counts)) { counts[c] = 0; order.push(c); }
      counts[c] += 1;
    });
    order.sort(function (a, b) { return counts[b] - counts[a] || (a < b ? -1 : 1); });
    catsEl.textContent = "";
    function item(label, value, n) {
      var li = el("li"), b = el("button", "cat", label + " (" + n + ")");
      b.type = "button";
      b.setAttribute("aria-pressed", String(cat === value));
      b.addEventListener("click", function () { cat = value; limit = PAGE_SIZE; go(""); renderList(); });
      li.appendChild(b);
      catsEl.appendChild(li);
    }
    item("전체", "", posts.length);
    order.forEach(function (c) { item(c, c, counts[c]); });
  }

  function renderList() {
    renderCats();
    var tokens = qEl.value.toLowerCase().split(/\\s+/).filter(Boolean);
    var found = [];
    posts.forEach(function (p, i) {
      if (cat && catOf(p) !== cat) return;
      var s = tokens.length ? score(p, tokens) : 1;
      if (s) found.push({i: i, s: s});
    });
    if (tokens.length) found.sort(function (a, b) { return b.s - a.s || byDateDesc(a.i, b.i); });
    else found.sort(function (a, b) { return byDateDesc(a.i, b.i); });

    listEl.textContent = "";
    found.slice(0, limit).forEach(function (f) {
      var p = posts[f.i], li = el("li", "post");
      li.appendChild(el("div", "date", prettyDate(p["일자"])));
      var t = el("button", "ptitle", titleOf(p));
      t.type = "button";
      t.addEventListener("click", function () { go("post/" + f.i); });
      li.appendChild(t);
      li.appendChild(tagsOf(p));
      li.appendChild(el("p", "summary", "원인: " + (text(p["원인"]) || "(내용 없음)") + "  →  조치: " + (text(p["조치"]) || "(내용 없음)")));
      listEl.appendChild(li);
    });
    if (found.length > limit) {
      var more = el("button", "more", "더 보기 (" + (found.length - limit) + "건 남음)");
      more.type = "button";
      more.addEventListener("click", function () { limit += PAGE_SIZE; renderList(); });
      var li = el("li");
      li.appendChild(more);
      listEl.appendChild(li);
    }
    var label = (cat ? cat + " · " : "") + (tokens.length ? "‘" + qEl.value.trim() + "’ 검색 결과 " : "전체 ");
    countEl.textContent = label + found.length + "건";
    emptyEl.hidden = found.length !== 0;
  }

  function section(title, body, cls) {
    articleEl.appendChild(el("h2", null, title));
    articleEl.appendChild(el("p", cls, text(body) || "(내용 없음)"));
  }
  function renderArticle(index) {
    var p = posts[index];
    if (!p) { go(""); return; }
    articleEl.textContent = "";
    var back = el("button", "back", "← 목록으로");
    back.type = "button";
    back.addEventListener("click", function () { go(""); });
    articleEl.appendChild(back);
    var head = el("div", "ahead");
    head.appendChild(el("div", "date", prettyDate(p["일자"])));
    head.appendChild(el("h1", null, titleOf(p)));
    head.appendChild(tagsOf(p));
    articleEl.appendChild(head);
    section("문제 상황", p["증상"]);
    section("원인", p["원인"]);
    section("해결 방법", p["조치"], "answer");
    if (p["기타 항목"]) section("기타 항목", p["기타 항목"]);
    if (p["확인 필요 이유"]) {
      articleEl.appendChild(el("div", "notice", NEED + ": " + text(p["확인 필요 이유"]) +
        (p["원래 값"] ? " (원래 값: " + text(p["원래 값"]) + ")" : "")));
    }
    if (p["중복 의심"]) articleEl.appendChild(el("div", "notice", text(p["중복 의심"])));

    var key = text(p["증상"]).trim().toLowerCase(), related = [];
    if (key) {
      posts.forEach(function (o, i) {
        if (i !== index && text(o["증상"]).trim().toLowerCase() === key) related.push(i);
      });
      related.sort(byDateDesc);
    }
    if (related.length) {
      articleEl.appendChild(el("h2", null, "같은 증상의 다른 사례 (" + related.length + "건)"));
      var ul = el("ul", "related");
      related.slice(0, 8).forEach(function (i) {
        var o = posts[i], li = el("li"), b = el("button", null, (text(o["조치"]) || "(조치 없음)"));
        b.type = "button";
        b.addEventListener("click", function () { go("post/" + i); });
        li.appendChild(b);
        li.appendChild(el("small", null, [prettyDate(o["일자"]), text(o["설비"]), "원인: " + (text(o["원인"]) || "-")].join(" · ")));
        ul.appendChild(li);
      });
      articleEl.appendChild(ul);
    }
    articleEl.appendChild(el("h2", null, "출처"));
    articleEl.appendChild(el("p", "source", text(p["원본 파일"]) + " / " + text(p["원본 시트"]) + " / " + text(p["원본 행"]) + "행"));
  }

  // 주소의 # 뒤에 화면 위치를 남겨 뒤로 가기가 되게 한다. 주소를 바꿀 수 없는 환경(미리보기 등)에서도 화면은 바뀐다.
  function go(route) {
    try { history.pushState(null, "", "#" + route); } catch (e) { /* 주소만 못 바꾼다 */ }
    route_(route);
  }
  function route_(route) {
    if (typeof route !== "string") route = location.hash.slice(1);
    var m = /^post\\/(\\d+)$/.exec(route);
    var isPost = !!m;
    listView.hidden = isPost;
    articleEl.hidden = !isPost;
    searchbar.hidden = isPost;
    if (isPost) { renderArticle(Number(m[1])); window.scrollTo(0, 0); }
    else renderList();
  }

  qEl.addEventListener("input", function () { limit = PAGE_SIZE; renderList(); });
  $("home").addEventListener("click", function () { cat = ""; qEl.value = ""; limit = PAGE_SIZE; go(""); });
  $("all").addEventListener("click", function () { cat = ""; limit = PAGE_SIZE; go(""); });
  $("find").addEventListener("click", function () { go(""); qEl.focus(); });
  $("theme").addEventListener("click", function () {
    var root = document.documentElement, dark = root.getAttribute("data-theme") === "dark";
    root.setAttribute("data-theme", dark ? "light" : "dark");
    this.textContent = dark ? "☀" : "☾";
  });
  window.addEventListener("popstate", function () { route_(); });
  route_();
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
