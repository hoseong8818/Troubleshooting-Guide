"""통합 엑셀을 QnA 게시판 HTML 한 장으로 만든다 (T054, 데모).

화면은 Apple Human Interface Guidelines의 원칙을 따른다: 큰 제목과 반투명 막대, 검색 칸, 둥근 묶음 목록,
시스템 색, 44px 이상의 누르는 영역, 라이트·다크·대비 강화·투명도 줄이기·동작 줄이기 대응, 색만으로 뜻을 전하지 않음.

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
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="color-scheme" content="light dark">
<title>트러블슈팅 노트</title>
<style>
:root{
  --bg:#f2f2f7; --cell:#ffffff; --cell-press:#e5e5ea; --fill:rgba(118,118,128,.12);
  --label:#000000; --label2:rgba(60,60,67,.6); --label3:rgba(60,60,67,.3);
  --sep:rgba(60,60,67,.29); --blue:#007aff; --blue-tint:rgba(0,122,255,.10);
  --orange:#c75100; --orange-tint:rgba(255,149,0,.16); --purple:#8944ab; --purple-tint:rgba(175,82,222,.14);
  --bar:rgba(249,249,249,.78); --bar-solid:#f9f9f9;
}
:root[data-theme="dark"]{
  --bg:#000000; --cell:#1c1c1e; --cell-press:#2c2c2e; --fill:rgba(118,118,128,.24);
  --label:#ffffff; --label2:rgba(235,235,245,.6); --label3:rgba(235,235,245,.3);
  --sep:rgba(84,84,88,.65); --blue:#0a84ff; --blue-tint:rgba(10,132,255,.18);
  --orange:#ff9f0a; --orange-tint:rgba(255,159,10,.18); --purple:#bf5af2; --purple-tint:rgba(191,90,242,.2);
  --bar:rgba(22,22,24,.78); --bar-solid:#161618;
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#000000; --cell:#1c1c1e; --cell-press:#2c2c2e; --fill:rgba(118,118,128,.24);
  --label:#ffffff; --label2:rgba(235,235,245,.6); --label3:rgba(235,235,245,.3);
  --sep:rgba(84,84,88,.65); --blue:#0a84ff; --blue-tint:rgba(10,132,255,.18);
  --orange:#ff9f0a; --orange-tint:rgba(255,159,10,.18); --purple:#bf5af2; --purple-tint:rgba(191,90,242,.2);
  --bar:rgba(22,22,24,.78); --bar-solid:#161618;
}}
@media (prefers-contrast:more){:root{--label2:var(--label);--sep:var(--label3);--blue:#0040dd}}
*{box-sizing:border-box}
[hidden]{display:none!important}
html,body{margin:0;background:var(--bg);color:var(--label)}
body{font:400 17px/1.47 -apple-system,BlinkMacSystemFont,"SF Pro Text","Apple SD Gothic Neo","Segoe UI","Malgun Gothic",system-ui,sans-serif;letter-spacing:-.01em;-webkit-font-smoothing:antialiased}
button{font:inherit;color:inherit;letter-spacing:inherit}
:focus-visible{outline:3px solid var(--blue);outline-offset:2px;border-radius:8px}

/* 막대: 반투명 재질(투명도 줄이기 설정이면 불투명) */
.bar{position:sticky;top:0;z-index:10;background:var(--bar);-webkit-backdrop-filter:saturate(180%) blur(20px);backdrop-filter:saturate(180%) blur(20px);border-bottom:.5px solid transparent;transition:border-color .2s}
.bar.scrolled{border-bottom-color:var(--sep)}
@media (prefers-reduced-transparency:reduce){.bar{background:var(--bar-solid);-webkit-backdrop-filter:none;backdrop-filter:none}}
.bar-in{max-width:1080px;margin:0 auto;height:52px;padding:0 16px;display:flex;align-items:center;gap:8px}
.bar-title{flex:1;text-align:center;font-weight:600;font-size:17px;opacity:0;transition:opacity .2s;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar.scrolled .bar-title{opacity:1}
.tbtn{min-width:44px;min-height:44px;padding:0 8px;border:0;background:none;color:var(--blue);font-size:17px;cursor:pointer;border-radius:10px;display:inline-flex;align-items:center;gap:2px}
.tbtn:hover{background:var(--fill)}
.tbtn .chev{font-size:24px;line-height:1;margin-top:-2px}
.side{min-width:88px}
.side.r{justify-content:flex-end}

.wrap{max-width:1080px;margin:0 auto;padding:0 16px 64px}
.large{margin:4px 0 12px}
.large h1{margin:0;font-size:34px;line-height:1.2;font-weight:700;letter-spacing:.01em}
.large p{margin:2px 0 0;font-size:15px;color:var(--label2)}
.search{position:relative;margin-bottom:16px}
.search .glyph{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--label2);font-size:15px;pointer-events:none}
#q{width:100%;height:40px;padding:0 12px 0 34px;border:0;border-radius:10px;background:var(--fill);color:var(--label);font-size:17px}
#q::placeholder{color:var(--label2)}
#q:focus{outline:none;box-shadow:0 0 0 3px var(--blue-tint),inset 0 0 0 1px var(--blue)}

.layout{display:grid;grid-template-columns:240px minmax(0,1fr);gap:28px;align-items:start}
.cats{position:sticky;top:64px;min-width:0}
.sec-h{margin:0 0 6px 16px;font-size:13px;font-weight:400;color:var(--label2);text-transform:none}
.group{background:var(--cell);border-radius:12px;overflow:hidden;margin:0 0 24px;padding:0;list-style:none}
.cat{width:100%;min-height:44px;display:flex;align-items:center;gap:8px;padding:0 14px;border:0;background:none;cursor:pointer;text-align:left;font-size:17px;position:relative}
.cat:hover{background:var(--cell-press)}
.cat .n{margin-left:auto;color:var(--label2);font-size:15px;font-variant-numeric:tabular-nums}
.cat[aria-pressed="true"]{background:var(--blue);color:#fff}
.cat[aria-pressed="true"] .n{color:rgba(255,255,255,.85)}
.group > li + li .cat::before,.group > li + li .row::before{content:"";position:absolute;left:16px;right:0;top:0;border-top:.5px solid var(--sep)}

#count{margin:0 0 6px 16px;font-size:13px;color:var(--label2)}
.row{width:100%;min-height:44px;display:flex;align-items:center;gap:12px;padding:11px 14px 11px 16px;border:0;background:none;cursor:pointer;text-align:left;position:relative}
.row:hover{background:var(--cell-press)}
.row .body{flex:1;min-width:0;display:block}
.row .s,.row .m{display:block}
.row .t{font-size:17px;font-weight:600;line-height:1.3;display:flex;flex-wrap:wrap;align-items:center;gap:6px}
.row .s{margin-top:2px;font-size:15px;color:var(--label2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.row .m{margin-top:3px;font-size:13px;color:var(--label2)}
.chev-r{color:var(--label3);font-size:20px;flex:none}
.caps-inline{display:inline-flex;gap:6px}
.caps-inline:empty{display:none}
.cap{display:inline-flex;align-items:center;height:20px;padding:0 8px;border-radius:10px;font-size:12px;font-weight:600;letter-spacing:0}
.cap.warn{background:var(--orange-tint);color:var(--orange)}
.cap.dup{background:var(--purple-tint);color:var(--purple)}
.more{justify-content:center;color:var(--blue);font-weight:400}
.empty{padding:56px 16px;text-align:center}
.empty .icon{font-size:44px;line-height:1;color:var(--label3)}
.empty h2{margin:12px 0 4px;font-size:22px;font-weight:600}
.empty p{margin:0;color:var(--label2);font-size:15px}

/* 글 보기 */
#article{max-width:720px;margin:0 auto}
.meta-line{margin:6px 0 0;font-size:15px;color:var(--label2)}
.caps{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
.text-cell{padding:12px 16px;font-size:17px;white-space:pre-wrap;overflow-wrap:anywhere}
.answer{background:var(--blue-tint)}
.answer .text-cell{font-weight:600}
.kv{display:flex;align-items:center;gap:12px;min-height:44px;padding:10px 16px;position:relative}
.group > li + li.kv::before,.group > li + li > .kv::before{content:"";position:absolute;left:16px;right:0;top:0;border-top:.5px solid var(--sep)}
.kv .k{flex:none}
.kv .v{margin-left:auto;color:var(--label2);text-align:right;overflow-wrap:anywhere}
.note{margin:-16px 16px 24px;font-size:13px;color:var(--label2)}
.note.warn{color:var(--orange)}

.view{animation:fade .2s ease-out}
@keyframes fade{from{opacity:0}to{opacity:1}}
@media (prefers-reduced-motion:reduce){.view{animation:none}.bar,.bar-title{transition:none}}
@media (max-width:860px){
  .layout{grid-template-columns:1fr;gap:0}
  .cats{position:static}
  .cats .sec-h{display:none}
  .cats .group{display:flex;gap:8px;overflow-x:auto;background:none;border-radius:0;margin:0 -16px 16px;padding:0 16px 2px;scrollbar-width:none}
  .cats .group > li + li .cat::before{display:none}
  .cat{width:auto;min-height:36px;white-space:nowrap;border-radius:18px;background:var(--fill);font-size:15px;padding:0 14px}
  .cat .n{margin-left:4px}
  .large h1{font-size:30px}
}
</style>
</head>
<body>
<header class="bar" id="bar">
<div class="bar-in">
<div class="side"><button class="tbtn" id="back" type="button" hidden><span class="chev" aria-hidden="true">‹</span>목록</button></div>
<div class="bar-title" id="bar-title">트러블슈팅 노트</div>
<div class="side r"><button class="tbtn" id="theme" type="button" aria-label="화면 모드 바꾸기">화면 모드</button></div>
</div>
</header>
<main class="wrap">
<section id="listview" class="view">
<div class="large">
<h1>트러블슈팅 노트</h1>
<p id="sub"></p>
</div>
<div class="search">
<span class="glyph" aria-hidden="true">⌕</span>
<input id="q" type="search" placeholder="증상이나 원인으로 검색 (예: 베어링 소음, PMP01)" aria-label="증상·원인·설비 검색" autocomplete="off" enterkeyhint="search">
</div>
<div class="layout">
<nav class="cats" aria-label="분류">
<h2 class="sec-h">분류</h2>
<ul class="group" id="cats"></ul>
</nav>
<div>
<p id="count" aria-live="polite"></p>
<ul class="group" id="list"></ul>
<div class="empty" id="empty" hidden>
<div class="icon" aria-hidden="true">⌕</div>
<h2>해당 사례 없음</h2>
<p>다른 단어로 검색하거나 분류를 바꿔 보세요.</p>
</div>
</div>
</div>
</section>
<article id="article" class="view" hidden></article>
</main>
<script type="application/json" id="board-data">__DATA__</script>
<script>
(function () {
  "use strict";
  var posts = JSON.parse(document.getElementById("board-data").textContent);
  var $ = function (id) { return document.getElementById(id); };
  var qEl = $("q"), catsEl = $("cats"), listEl = $("list"), countEl = $("count"), emptyEl = $("empty");
  var listView = $("listview"), articleEl = $("article"), backBtn = $("back"), bar = $("bar"), barTitle = $("bar-title");
  var NEED = "확인 필요", NO_CAT = "(분류 없음)", PAGE_SIZE = 50;
  var cat = "", limit = PAGE_SIZE, listScroll = 0;
  var WEIGHT = {"증상": 3, "원인": 3, "조치": 1, "설비": 2, "분류": 1, "기타 항목": 1};
  var DATE_RE = /^(\\d{4})-(\\d{2})-(\\d{2})$/;

  function text(v) { return v === null || v === undefined ? "" : String(v); }
  function el(tag, cls, str) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (str !== undefined) e.textContent = str;
    return e;
  }
  function catOf(p) { return text(p["분류"]) || NO_CAT; }
  function titleOf(p) { return text(p["증상"]) || "(증상 없음)"; }
  function dateKey(p) { var s = text(p["일자"]); return DATE_RE.test(s) ? s : ""; }
  function prettyDate(v) {
    var s = text(v), m = DATE_RE.exec(s);
    if (m) return m[1] + "년 " + Number(m[2]) + "월 " + Number(m[3]) + "일";
    if (s === NEED) return "일자 확인 필요";
    return s || "일자 없음";
  }
  function caps(p) {
    var box = el("span", "caps-inline");
    if (p["확인 필요 이유"]) box.appendChild(el("span", "cap warn", NEED));
    if (p["중복 의심"]) box.appendChild(el("span", "cap dup", "중복 의심"));
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
      var li = el("li"), b = el("button", "cat");
      b.type = "button";
      b.appendChild(el("span", null, label));
      b.appendChild(el("span", "n", String(n)));
      b.setAttribute("aria-pressed", String(cat === value));
      b.addEventListener("click", function () { cat = value; limit = PAGE_SIZE; renderList(); });
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
      var p = posts[f.i], li = el("li"), b = el("button", "row");
      b.type = "button";
      var body = el("span", "body");
      var t = el("span", "t");
      t.appendChild(el("span", null, titleOf(p)));
      t.appendChild(caps(p));
      body.appendChild(t);
      body.appendChild(el("span", "s", (text(p["원인"]) || "원인 없음") + " → " + (text(p["조치"]) || "조치 없음")));
      body.appendChild(el("span", "m", [prettyDate(p["일자"]), text(p["설비"]), text(p["분류"])].filter(Boolean).join(" · ")));
      b.appendChild(body);
      b.appendChild(el("span", "chev-r", "›"));
      b.setAttribute("aria-label", titleOf(p) + ", " + prettyDate(p["일자"]));
      b.addEventListener("click", function () { listScroll = window.scrollY; go("post/" + f.i); });
      li.appendChild(b);
      listEl.appendChild(li);
    });
    if (found.length > limit) {
      var li = el("li"), more = el("button", "row more", "더 보기 (" + (found.length - limit) + "건 남음)");
      more.type = "button";
      more.addEventListener("click", function () { limit += PAGE_SIZE; renderList(); });
      li.appendChild(more);
      listEl.appendChild(li);
    }
    listEl.hidden = found.length === 0;
    emptyEl.hidden = found.length !== 0;
    var where = cat ? cat + " · " : "";
    countEl.textContent = where + (tokens.length ? "검색 결과 " : "전체 ") + found.length + "건";
  }

  function group(title, items, cls) {
    if (title) articleEl.appendChild(el("h2", "sec-h", title));
    var ul = el("ul", "group" + (cls ? " " + cls : ""));
    items.forEach(function (node) { ul.appendChild(node); });
    articleEl.appendChild(ul);
    return ul;
  }
  function textCell(value) { var li = el("li"); li.appendChild(el("div", "text-cell", text(value) || "(내용 없음)")); return li; }
  function kv(k, v) { var li = el("li", "kv"); li.appendChild(el("span", "k", k)); li.appendChild(el("span", "v", text(v) || "-")); return li; }

  function renderArticle(index) {
    var p = posts[index];
    if (!p) { go(""); return; }
    articleEl.textContent = "";
    var head = el("div", "large");
    head.appendChild(el("h1", null, titleOf(p)));
    head.appendChild(el("p", "meta-line", [prettyDate(p["일자"]), text(p["설비"]), text(p["분류"])].filter(Boolean).join(" · ")));
    var c = el("div", "caps");
    if (p["확인 필요 이유"]) c.appendChild(el("span", "cap warn", NEED));
    if (p["중복 의심"]) c.appendChild(el("span", "cap dup", "중복 의심"));
    if (c.childNodes.length) head.appendChild(c);
    articleEl.appendChild(head);

    group("문제 상황", [textCell(p["증상"])]);
    group("원인", [textCell(p["원인"])]);
    group("해결 방법", [textCell(p["조치"])], "answer");
    if (p["확인 필요 이유"]) {
      articleEl.appendChild(el("p", "note warn", NEED + ": " + text(p["확인 필요 이유"]) +
        (p["원래 값"] ? " (원래 값: " + text(p["원래 값"]) + ")" : "")));
    }
    group("정보", [kv("설비", p["설비"]), kv("분류", p["분류"]), kv("일자", prettyDate(p["일자"])),
      kv("출처", text(p["원본 파일"]) + " / " + text(p["원본 시트"]) + " / " + text(p["원본 행"]) + "행")]);
    if (p["기타 항목"]) group("기타 항목", [textCell(p["기타 항목"])]);
    if (p["중복 의심"]) articleEl.appendChild(el("p", "note", "같은 설비·일자·증상의 행이 더 있습니다."));

    var key = text(p["증상"]).trim().toLowerCase(), related = [];
    if (key) {
      posts.forEach(function (o, i) {
        if (i !== index && text(o["증상"]).trim().toLowerCase() === key) related.push(i);
      });
      related.sort(byDateDesc);
    }
    if (related.length) {
      var rows = related.slice(0, 8).map(function (i) {
        var o = posts[i], li = el("li"), b = el("button", "row");
        b.type = "button";
        var body = el("span", "body");
        body.appendChild(el("span", "t", text(o["조치"]) || "(조치 없음)"));
        body.appendChild(el("span", "m", [prettyDate(o["일자"]), text(o["설비"]), "원인: " + (text(o["원인"]) || "-")].join(" · ")));
        b.appendChild(body);
        b.appendChild(el("span", "chev-r", "›"));
        b.addEventListener("click", function () { go("post/" + i); });
        li.appendChild(b);
        return li;
      });
      group("같은 증상의 다른 사례 " + related.length + "건", rows);
    }
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
    backBtn.hidden = !isPost;
    if (isPost) {
      renderArticle(Number(m[1]));
      barTitle.textContent = titleOf(posts[Number(m[1])] || {});
      window.scrollTo(0, 0);
    } else {
      barTitle.textContent = "트러블슈팅 노트";
      renderList();
      window.scrollTo(0, listScroll);
    }
    onScroll();
  }
  function onScroll() { bar.classList.toggle("scrolled", window.scrollY > 40); }

  $("sub").textContent = "사례 " + posts.length + "건 · 이 파일 안에서만 검색합니다";
  qEl.addEventListener("input", function () { limit = PAGE_SIZE; listScroll = 0; renderList(); });
  backBtn.addEventListener("click", function () { go(""); });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && !articleEl.hidden) go("");
    if (e.key === "/" && document.activeElement !== qEl && articleEl.hidden) { e.preventDefault(); qEl.focus(); }
  });
  $("theme").addEventListener("click", function () {
    var root = document.documentElement, cur = root.getAttribute("data-theme");
    var dark = cur ? cur === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches;
    root.setAttribute("data-theme", dark ? "light" : "dark");
  });
  window.addEventListener("scroll", onScroll, {passive: true});
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
