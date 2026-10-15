"""MarginIQ Help viewer: renders the Markdown guides to styled, linked HTML.

Standard library only (no extra dependency, works offline). Every guide in
``DOCS`` is converted to ``<name>.html`` in a temp folder with a shared
sidebar, an "On this page" table of contents and an in-page search box; links
between guides are rewritten to the rendered pages, links to other repository
files open the local file. A "Supported Projects & Spec Limits" page is
generated from ``projects.json`` so it always matches the active registry.

    python rmt_help.py [readme|quickstart|cli|install|pipeline|projects]
                       [--anchor ID] [--out DIR] [--no-open]
"""
from __future__ import annotations

import argparse
import html
import re
import sys
import tempfile
import webbrowser
from pathlib import Path
from typing import Callable

from rmt_project import TOOL_NAME, TOOL_SUBTITLE, TOOL_VERSION, load_projects, project_label

HERE = Path(__file__).resolve().parent
DEFAULT_OUT = Path(tempfile.gettempdir()) / "MarginIQ_help"

# key -> (markdown file, sidebar title)
DOCS: dict[str, tuple[str, str]] = {
    "readme": ("README.md", "User Guide"),
    "quickstart": ("RMT_QUICKSTART.md", "Quick Start"),
    "cli": ("CLI_AUTOMATION.md", "Command Line & Automation"),
    "install": ("INSTALL.md", "Installation Guide"),
    "pipeline": ("RMT_LOG_PIPELINE_README.md", "Pipeline Reference"),
}
PROJECTS_PAGE = ("projects", "projects.html", "Supported Projects & Spec Limits")

# Section anchors used by the GUI Help menu (GitHub-style heading ids).
ANCHOR_WHATS_NEW = "-whats-new-in-20"
ANCHOR_FAQ = "-frequently-asked-questions"

_INLINE_TAGS = r"br|b|i|em|strong|sup|sub|kbd|img|code|span"


# ---------------------------------------------------------------------------
# Markdown -> HTML (the subset used by the MarginIQ guides)
# ---------------------------------------------------------------------------

def slugify(text: str) -> str:
    """GitHub heading id: strip links/markup, lower-case, drop punctuation."""
    t = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text.strip())
    t = t.lower().replace("`", "")
    t = re.sub(r"[^\w\- ]", "", t)
    return t.replace(" ", "-")


class _Renderer:
    def __init__(self, link_fn: Callable[[str], str]) -> None:
        self.link_fn = link_fn
        self.headings: list[tuple[int, str, str]] = []   # (level, id, text)
        self._ids: dict[str, int] = {}
        self._codes: list[str] = []

    # ── inline ─────────────────────────────────────────────────────────
    def inline(self, s: str) -> str:
        spans: list[str] = []

        def _code(m: re.Match) -> str:
            code = m.group(1) if m.group(1) is not None else m.group(2)
            spans.append(f"<code>{html.escape(code)}</code>")
            return f"\x02{len(spans) - 1}\x02"

        s = re.sub(r"``\s?(.+?)\s?``|`([^`]+)`", _code, s)
        s = s.replace("\\|", "|").replace("\\*", "\x03")
        s = html.escape(s, quote=False)
        s = re.sub(rf"&lt;(/?(?:{_INLINE_TAGS})\b.*?)&gt;",
                   lambda m: "<" + self._fix_tag(html.unescape(m.group(1))) + ">", s)
        s = re.sub(r"&lt;(https?://[^\s&]+)&gt;", r'<a href="\1">\1</a>', s)
        s = re.sub(r"!\[([^\]]*)\]\(([^)\s]+)\)",
                   lambda m: f'<img alt="{m.group(1)}" src="{self.link_fn(m.group(2))}">', s)
        s = re.sub(r"\[([^\]]+)\]\(([^)\s]+)(?:\s+&quot;[^)]*&quot;)?\)",
                   lambda m: self._link(m.group(1), m.group(2)), s)
        s = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)
        s = re.sub(r"(?<![\w*])\*(?![\s*])(.+?)(?<![\s*])\*(?![\w*])", r"<em>\1</em>", s)
        s = s.replace("\x03", "*")
        return re.sub(r"\x02(\d+)\x02", lambda m: spans[int(m.group(1))], s)

    def _fix_tag(self, tag: str) -> str:
        """Rewrite src= of inline <img> tags to the resolved location."""
        return re.sub(r'src="([^"]+)"', lambda m: f'src="{self.link_fn(m.group(1))}"', tag)

    def _link(self, text: str, href: str) -> str:
        target = self.link_fn(html.unescape(href))
        ext = ' target="_blank" rel="noopener"' if target.startswith("http") else ""
        return f'<a href="{html.escape(target)}"{ext}>{text}</a>'

    # ── blocks ─────────────────────────────────────────────────────────
    def render(self, text: str) -> str:
        lines = text.replace("\r\n", "\n").split("\n")
        return self._blocks(self._extract_fences(lines))

    def _extract_fences(self, lines: list[str]) -> list[str]:
        out: list[str] = []
        i = 0
        while i < len(lines):
            m = re.match(r"^(\s*)(```|~~~)\s*([\w+-]*)", lines[i])
            if not m:
                out.append(lines[i]); i += 1; continue
            indent, fence, lang = m.group(1), m.group(2), m.group(3)
            body: list[str] = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith(fence):
                ln = lines[i]
                body.append(ln[len(indent):] if ln.startswith(indent) else ln.lstrip())
                i += 1
            i += 1
            cls = f' class="language-{lang}"' if lang else ""
            self._codes.append(f"<pre><code{cls}>{html.escape(chr(10).join(body))}</code></pre>")
            out.append(f"{indent}\x04{len(self._codes) - 1}\x04")
        return out

    def _heading_id(self, text: str) -> str:
        base = slugify(text)
        n = self._ids.get(base, 0)
        self._ids[base] = n + 1
        return base if n == 0 else f"{base}-{n}"

    @staticmethod
    def _is_list_item(line: str) -> re.Match | None:
        return re.match(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$", line)

    @staticmethod
    def _is_table_sep(line: str) -> bool:
        return bool(re.match(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$", line)) and "-" in line

    @staticmethod
    def _cells(line: str) -> list[str]:
        s = line.strip()
        s = s[1:] if s.startswith("|") else s
        s = s[:-1] if s.endswith("|") and not s.endswith("\\|") else s
        return [c.strip() for c in re.split(r"(?<!\\)\|", s)]

    def _blocks(self, lines: list[str]) -> str:
        out: list[str] = []
        i = 0
        while i < len(lines):
            line = lines[i]
            st = line.strip()
            if not st:
                i += 1; continue
            m = re.match(r"^\x04(\d+)\x04$", st)
            if m:
                out.append(self._codes[int(m.group(1))]); i += 1; continue
            m = re.match(r"^(#{1,6})\s+(.*?)\s*#*$", st)
            if m:
                lvl, txt = len(m.group(1)), m.group(2)
                hid = self._heading_id(txt)
                plain = re.sub(r"<[^>]+>", "", self.inline(txt))
                self.headings.append((lvl, hid, html.unescape(plain)))
                out.append(f'<h{lvl} id="{hid}">{self.inline(txt)}'
                           f'<a class="anchor" href="#{hid}" aria-label="link">#</a></h{lvl}>')
                i += 1; continue
            if re.match(r"^(-{3,}|\*{3,}|_{3,})$", st):
                out.append("<hr>"); i += 1; continue
            if st.startswith("|") and i + 1 < len(lines) and self._is_table_sep(lines[i + 1]):
                head = self._cells(st)
                aligns = []
                for c in self._cells(lines[i + 1]):
                    a = "center" if c.startswith(":") and c.endswith(":") else (
                        "right" if c.endswith(":") else "")
                    aligns.append(f' style="text-align:{a}"' if a else "")
                rows = []
                i += 2
                while i < len(lines) and lines[i].strip().startswith("|"):
                    rows.append(self._cells(lines[i])); i += 1
                al = lambda k: aligns[k] if k < len(aligns) else ""
                th = "".join(f"<th{al(k)}>{self.inline(c)}</th>" for k, c in enumerate(head))
                body = "".join(
                    "<tr>" + "".join(f"<td{al(k)}>{self.inline(c)}</td>" for k, c in enumerate(r)) + "</tr>"
                    for r in rows)
                out.append(f'<div class="tbl"><table><thead><tr>{th}</tr></thead>'
                           f"<tbody>{body}</tbody></table></div>")
                continue
            if st.startswith(">"):
                quote = []
                while i < len(lines) and lines[i].strip().startswith(">"):
                    quote.append(re.sub(r"^\s*>\s?", "", lines[i])); i += 1
                out.append(f"<blockquote>{self._blocks(quote)}</blockquote>")
                continue
            if self._is_list_item(line):
                i = self._list(lines, i, out)
                continue
            if re.match(r"^<(img|div|p|table|details|br)\b", st, re.I):
                out.append(self._fix_tag(st)); i += 1; continue
            para = [st]
            i += 1
            while i < len(lines):
                nxt = lines[i].strip()
                if (not nxt or nxt.startswith(("#", ">", "|")) or re.match(r"^\x04\d+\x04$", nxt)
                        or self._is_list_item(lines[i]) or re.match(r"^(-{3,}|\*{3,})$", nxt)):
                    break
                para.append(nxt); i += 1
            out.append(f"<p>{self.inline(' '.join(para))}</p>")
        return "\n".join(out)

    def _list(self, lines: list[str], i: int, out: list[str]) -> int:
        first = self._is_list_item(lines[i])
        base = len(first.group(1))
        ordered = first.group(2)[0].isdigit()
        items: list[list[str]] = []
        while i < len(lines):
            line = lines[i]
            m = self._is_list_item(line)
            indent = len(line) - len(line.lstrip())
            if m and len(m.group(1)) == base and m.group(2)[0].isdigit() == ordered:
                items.append([m.group(3)]); i += 1; continue
            if not line.strip():
                nxt = next((l for l in lines[i + 1:] if l.strip()), "")
                n_ind = len(nxt) - len(nxt.lstrip())
                if nxt and n_ind > base:
                    items[-1].append(""); i += 1; continue
                break
            if indent > base:
                items[-1].append(line[min(indent, base + 2):] if indent >= base + 2 else line.strip())
                i += 1; continue
            break
        tag = "ol" if ordered else "ul"
        start = ""
        if ordered:
            num = int(re.match(r"\d+", first.group(2)).group(0))
            start = f' start="{num}"' if num != 1 else ""
        parts = []
        for body in items:
            inner = self._blocks(body)
            if inner.startswith("<p>") and inner.count("<p>") == 1 and inner.endswith("</p>"):
                inner = inner[3:-4]
            elif inner.startswith("<p>"):
                inner = re.sub(r"^<p>(.*?)</p>", r"\1", inner, count=1, flags=re.S)
            parts.append(f"<li>{inner}</li>")
        out.append(f"<{tag}{start}>" + "".join(parts) + f"</{tag}>")
        return i


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

_CSS = """
:root{--blue:#0071C5;--navy:#1E293B;--bg:#f6f8fb;--line:#d8dee8;--muted:#5b6575}
*{box-sizing:border-box}html{scroll-padding-top:70px}
body{margin:0;font:15px/1.6 "Segoe UI",system-ui,sans-serif;color:#1f2937;background:var(--bg)}
header{position:sticky;top:0;z-index:5;display:flex;align-items:center;gap:12px;padding:10px 22px;
 background:linear-gradient(90deg,#00305e,var(--blue));color:#fff;box-shadow:0 2px 6px rgba(0,0,0,.18)}
header img{width:34px;height:34px}header .t{font-size:18px;font-weight:600}
header .s{font-size:12px;opacity:.85}header .sp{flex:1}
#q{width:260px;padding:6px 10px;border:0;border-radius:6px;font:inherit}
#qn{font-size:12px;min-width:70px;opacity:.9}
.wrap{display:grid;grid-template-columns:250px minmax(0,1fr) 230px;gap:22px;max-width:1500px;margin:0 auto;padding:18px 22px}
nav.side,aside.toc{position:sticky;top:72px;align-self:start;max-height:calc(100vh - 90px);overflow:auto;font-size:14px}
nav.side a,aside.toc a{display:block;padding:5px 10px;border-radius:6px;color:var(--navy);text-decoration:none}
nav.side a:hover,aside.toc a:hover{background:#e6eef8}nav.side a.cur{background:var(--blue);color:#fff;font-weight:600}
nav.side h4,aside.toc h4{margin:14px 10px 6px;font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted)}
aside.toc a.l3{padding-left:22px;font-size:13px;color:var(--muted)}
main{background:#fff;border:1px solid var(--line);border-radius:10px;padding:26px 38px;min-width:0}
h1,h2,h3,h4{color:var(--navy);line-height:1.3}h1{font-size:28px;margin-top:0}
h2{font-size:22px;border-bottom:1px solid var(--line);padding-bottom:6px;margin-top:34px}
h3{font-size:18px;margin-top:26px}.anchor{margin-left:8px;color:#b8c2d0;text-decoration:none;font-size:.8em;visibility:hidden}
h1:hover .anchor,h2:hover .anchor,h3:hover .anchor,h4:hover .anchor{visibility:visible}
a{color:var(--blue)}code{background:#eef2f7;padding:1px 5px;border-radius:4px;font:13px Consolas,"Cascadia Mono",monospace}
pre{background:#0f172a;color:#e2e8f0;padding:14px 16px;border-radius:8px;overflow:auto;position:relative}
pre code{background:none;color:inherit;padding:0;font-size:13px}
pre button{position:absolute;top:8px;right:8px;font-size:11px;padding:2px 8px;border:0;border-radius:4px;background:#334155;color:#e2e8f0;cursor:pointer}
.tbl{overflow-x:auto;margin:14px 0}table{border-collapse:collapse;width:100%;font-size:14px}
th,td{border:1px solid var(--line);padding:6px 10px;vertical-align:top;text-align:left}
th{background:#eef3fa}tr:nth-child(even) td{background:#fafcff}
blockquote{margin:14px 0;padding:8px 16px;border-left:4px solid var(--blue);background:#f1f6fc;color:#334155}
blockquote p{margin:4px 0}hr{border:0;border-top:1px solid var(--line);margin:26px 0}
mark{background:#ffe066;padding:0 1px}mark.on{background:#ff9f1c}
footer{font-size:12px;color:var(--muted);text-align:center;padding:12px}
@media (max-width:1100px){.wrap{grid-template-columns:220px minmax(0,1fr)}aside.toc{display:none}}
@media print{header,nav.side,aside.toc{display:none}.wrap{display:block}main{border:0}}
"""

_JS = """
(function(){
const main=document.querySelector('main'),q=document.getElementById('q'),qn=document.getElementById('qn');
let hits=[],cur=-1;
function clear(){document.querySelectorAll('main mark').forEach(m=>{m.replaceWith(document.createTextNode(m.textContent))});main.normalize();hits=[];cur=-1;qn.textContent='';}
function find(t){clear();if(t.length<2)return;const re=new RegExp(t.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\\\$&'),'gi');
 const w=document.createTreeWalker(main,NodeFilter.SHOW_TEXT,{acceptNode:n=>n.parentNode.closest('script,style')?2:1});
 const nodes=[];while(w.nextNode())nodes.push(w.currentNode);
 nodes.forEach(n=>{const s=n.nodeValue;if(!re.test(s))return;re.lastIndex=0;const f=document.createDocumentFragment();let last=0,m;
  while((m=re.exec(s))){f.append(s.slice(last,m.index));const k=document.createElement('mark');k.textContent=m[0];f.append(k);hits.push(k);last=m.index+m[0].length;}
  f.append(s.slice(last));n.replaceWith(f);});
 qn.textContent=hits.length?hits.length+' match'+(hits.length>1?'es':''):'no match';go(0);}
function go(i){if(!hits.length)return;if(cur>=0)hits[cur].classList.remove('on');cur=(i+hits.length)%hits.length;
 hits[cur].classList.add('on');hits[cur].scrollIntoView({block:'center'});qn.textContent=(cur+1)+' / '+hits.length;}
let tm;q.addEventListener('input',()=>{clearTimeout(tm);tm=setTimeout(()=>find(q.value.trim()),250);});
q.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();go(cur+(e.shiftKey?-1:1));}if(e.key==='Escape'){q.value='';clear();}});
document.addEventListener('keydown',e=>{if((e.ctrlKey&&e.key==='f')||(e.key==='/'&&document.activeElement!==q)){e.preventDefault();q.focus();q.select();}});
document.querySelectorAll('pre').forEach(p=>{const b=document.createElement('button');b.textContent='Copy';
 b.onclick=()=>{navigator.clipboard.writeText(p.querySelector('code').innerText).then(()=>{b.textContent='Copied';setTimeout(()=>b.textContent='Copy',1200);});};p.append(b);});
})();
"""


def _page(title: str, current: str, body: str, headings: list[tuple[int, str, str]],
          source_note: str) -> str:
    icon = (HERE / "assets" / "marginiq.png").as_uri()
    side = ['<h4>Guides</h4>']
    for key, (md, name) in DOCS.items():
        cls = ' class="cur"' if key == current else ""
        side.append(f'<a{cls} href="{Path(md).stem}.html">{html.escape(name)}</a>')
    cls = ' class="cur"' if current == PROJECTS_PAGE[0] else ""
    side.append('<h4>Reference</h4>')
    side.append(f'<a{cls} href="{PROJECTS_PAGE[1]}">{html.escape(PROJECTS_PAGE[2])}</a>')
    toc = ['<h4>On this page</h4>'] + [
        f'<a class="l{lvl}" href="#{hid}">{html.escape(txt.strip())}</a>'
        for lvl, hid, txt in headings if lvl in (2, 3)]
    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(title)} - {TOOL_NAME} Help</title><style>{_CSS}</style></head>
<body><header><img src="{icon}" alt=""><div><div class="t">{TOOL_NAME} Help</div>
<div class="s">{html.escape(TOOL_SUBTITLE)} &middot; v{TOOL_VERSION}</div></div><div class="sp"></div>
<input id="q" type="search" placeholder="Search this page  (Ctrl+F or /)"><span id="qn"></span></header>
<div class="wrap"><nav class="side">{''.join(side)}</nav>
<main>{body}</main><aside class="toc">{''.join(toc)}</aside></div>
<footer>{html.escape(source_note)}</footer><script>{_JS}</script></body></html>"""


def _make_link_fn(doc_dir: Path) -> Callable[[str], str]:
    rendered = {Path(md).name.lower(): f"{Path(md).stem}.html" for md, _ in DOCS.values()}

    def _fn(href: str) -> str:
        if re.match(r"^(https?:|mailto:|#|file:)", href):
            return href
        path, _, frag = href.partition("#")
        name = Path(path).name.lower()
        if name in rendered and (doc_dir / path).resolve().parent == HERE:
            return rendered[name] + (f"#{frag}" if frag else "")
        target = (doc_dir / path).resolve()
        return target.as_uri() + (f"#{frag}" if frag else "") if target.exists() else href
    return _fn


def render_doc(key: str) -> tuple[str, str]:
    """(html file name, page html) for one guide."""
    md, title = DOCS[key]
    src = HERE / md
    try:
        text = src.read_text(encoding="utf-8")
    except (FileNotFoundError, PermissionError) as exc:
        text = f"# {title}\n\nCould not read `{src}`: {exc}"
    r = _Renderer(_make_link_fn(src.parent))
    body = r.render(text)
    note = f"Rendered from the local copy of {md} ({src}). Edit the .md file to change this page."
    return f"{Path(md).stem}.html", _page(title, key, body, r.headings, note)


def _projects_markdown() -> str:
    """Markdown for the Supported Projects & Spec Limits page (from projects.json)."""
    reg = load_projects()
    out = ["# Supported Projects & Spec Limits", "",
           f"Generated from `projects.json` (registry default: **{reg['default']}**). "
           "Pick the project on Tab 1 or pass `--project <key>` on the command line.", ""]
    for key, p in reg["projects"].items():
        out += [f"## {project_label(key, reg)}", "", "| Rule | Value |", "|------|-------|",
                f"| Key / code | `{key}` / `{p.get('code', key)}` |",
                f"| Memory | {p.get('memory', '-')} |"]
        aliases = p.get("param_aliases") or {}
        out.append("| Header aliases | " + (", ".join(f"`{a}` \u2192 `{b}`" for a, b in aliases.items())
                                         or "none (built-in `RxVref` \u2192 `RxDqVrefByte` always applies)") + " |")
        fm = p.get("rmt_fallback_markers") or []
        out.append("| RMT block markers | `START_RMT`" + "".join(f", `{m}`" for m in fm) + " |")
        out.append("| RMT data only inside the Rank Margin Tool task | "
                   + ("**yes** - logs without it print `NO RMT DATA` and are skipped"
                      if p.get("require_rmt_task") else "no (built-in behaviour)") + " |")
        dtr = p.get("dtr") or {}
        out.append("| Thermal (DTR) flow | " + (
            "first / last *Rank Margin Tool* run; temperatures from "
            + ("**BCRH / BHRC** in the file name" if dtr.get("temps_from_filename") else "PHY Temperature1")
            if dtr.get("layout") == "rmt_task_runs" else
            "*Temp Drift RMT Test* + *Hit Enter* + *Rank Margin Tool*; temperatures from PHY Temperature1") + " |")
        pv = p.get("platform_variants") or {}
        out.append("| Platform variants | " + (", ".join(f"{k} \u2192 **{v}**" for k, v in pv.items()) or "-") + " |")
        lim = p.get("rmt_limits") or {}
        out.append("| \u00b1Ref lines | " + ("**spec RMT limits** per memory type + frequency (below); "
                                         "`--no-rmt-limits` = axis JSON" if lim else
                                         "axis JSON (`jmp_axis_settings.json`, default \u00b110)") + " |")
        out.append("")
        if lim:
            units = lim.get("physical_units", {})
            if lim.get("_source"):
                out += [f"> Source: {lim['_source']}. {lim.get('_mapping', '')}", ""]
            for fam, fd in lim.items():
                if fam.startswith("_") or not isinstance(fd, dict) or "frequencies" not in fd:
                    continue
                freqs = fd["frequencies"]
                params = list(next(iter(freqs.values()))["ticks"])
                out += [f"### {fam} spec RMT limits (ticks, physical in brackets)", "",
                        "| MT/s | " + " | ".join(f"{q} ({units.get(q, '')})" for q in params) + " |",
                        "|---:|" + "---:|" * len(params)]
                for f, entry in sorted(freqs.items(), key=lambda kv: int(kv[0])):
                    cells = []
                    for q in params:
                        t, ph = entry["ticks"][q], entry.get("physical", {}).get(q)
                        tv = f"{t[1]:g}" if t[0] == t[1] else f"-{t[0]:g}/+{t[1]:g}"
                        cells.append(f"{tv} ({ph[1]:g})" if ph else tv)
                    out.append(f"| {f} | " + " | ".join(cells) + " |")
                out += ["", f"CPU vdd: {fd.get('cpu_vdd', '-')} V. A log frequency that is not listed "
                        "uses the nearest one (with a warning).", ""]
    return "\n".join(out)


def build_help(out_dir: Path = DEFAULT_OUT) -> Path:
    """Render every guide + the projects page into *out_dir*; returns the folder."""
    out_dir.mkdir(parents=True, exist_ok=True)
    for key in DOCS:
        name, page = render_doc(key)
        (out_dir / name).write_text(page, encoding="utf-8")
    r = _Renderer(_make_link_fn(HERE))
    body = r.render(_projects_markdown())
    (out_dir / PROJECTS_PAGE[1]).write_text(
        _page(PROJECTS_PAGE[2], PROJECTS_PAGE[0], body, r.headings,
              f"Generated from {HERE / 'projects.json'} - always matches the active registry."),
        encoding="utf-8")
    return out_dir


def open_help(page: str = "readme", anchor: str | None = None,
              out_dir: Path = DEFAULT_OUT, open_browser: bool = True) -> Path:
    """Build the help pages and open *page* (DOCS key or "projects") at *anchor*.

    Windows drops the #fragment when a file URL is handed to the shell, so a
    tiny redirect page carries the anchor.
    """
    build_help(out_dir)
    name = PROJECTS_PAGE[1] if page == PROJECTS_PAGE[0] else f"{Path(DOCS[page][0]).stem}.html"
    target = name + (f"#{anchor}" if anchor else "")
    launcher = out_dir / "open.html"
    launcher.write_text(
        f'<!DOCTYPE html><meta charset="utf-8"><meta http-equiv="refresh" content="0;url={target}">'
        f'<script>location.replace("{target}")</script><a href="{target}">Open {TOOL_NAME} Help</a>',
        encoding="utf-8")
    if open_browser and not webbrowser.open(launcher.as_uri()):
        raise OSError("No web browser could be started.")
    return out_dir / name


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=f"{TOOL_NAME} Help viewer (Markdown guides as HTML).")
    ap.add_argument("page", nargs="?", default="readme", choices=[*DOCS, PROJECTS_PAGE[0]])
    ap.add_argument("--anchor", default=None, help="Section id to open, e.g. -whats-new-in-20")
    ap.add_argument("--out", default=str(DEFAULT_OUT), help="Output folder for the HTML pages")
    ap.add_argument("--no-open", action="store_true", help="Only build the pages")
    args = ap.parse_args(argv)
    try:
        path = open_help(args.page, args.anchor, Path(args.out), open_browser=not args.no_open)
    except (OSError, PermissionError) as exc:
        print(f"Could not build/open the help: {exc}")
        return 2
    print(f"Help page: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
