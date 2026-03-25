#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build an interactive HTML vertical timeline from a spreadsheet.
March 2026
Authors: Jason Keller, assisted by ChatGPT

Expected columns (first three required):
  1. Era
  2. Year / year label / decade / range
  3. Event text

Features:
- Vertical scrolling modern timeline UI
- Era / Year / Event zoom levels
- Live search
- Era filter
- Expand / collapse controls
- Basic Markdown in event text
- Optional image per event
- Optional full-page background image via --bg
- Intro header paragraph and external links

Usage:
  python build_timeline_fixed.py input.xlsx output.html [--bg path/to/image optional]
  (ex. python build_timeline.py neuroHistory.xlsx output.html --bg cajalReticularFormation.webp)
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import mimetypes
import pathlib
import re
from typing import Any

import pandas as pd

YEAR_RE = re.compile(r"-?\d+")
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp", ".tif", ".tiff"}

INTRO_HTML = """
<div class="header-card">
  <p>
    This is a timeline of events in the history of neuroscience that have 
    significantly interested me and/or shaped my view, gathered from a variety of sources and with a bit of my interpretation added. 
    It's definitely NOT comprehensive and is notably biased towards recent and mammalian studies, as well as Western history. 
    It’s also a constant work in progress (especially for work in the past decade or two where long term influence remains to 
    be seen) and has some approximate dates. Feel free to make suggestions for corrections and/or additions (open an issue or DM 
    @neurojak.bsky.social), or use it to start your own mix tape.
  </p>
  
  <p>
    The history of neuroscience is an exciting story that we don’t yet know the ending to.
    In the interest of good story-telling, I try to be as concise as possible, but each of the timepoints below could probably be a book in itself, and 
    many insightful books and articles have previously been written on the topic
    (e.g. see <a href="https://www.sfn.org/about/history-of-neuroscience" target="_blank" rel="noopener">SFN</a>,
    the Journal of the History of the Neurosciences, Matthew Cobb’s <em>The Idea of the Brain</em>).
  </p>

  <h3>Links to some other neuroscience history resources:</h3>
  <ul>
    <li><a href="https://faculty.washington.edu/chudler/hist.html" target="_blank" rel="noopener">Eric Chudler</a></li>
    <li><a href="https://www.scaruffi.com/mind/ns.html" target="_blank" rel="noopener">Piero Scaruffi</a></li>
    <li><a href="https://en.wikipedia.org/wiki/History_of_neuroscience" target="_blank" rel="noopener">Wikipedia</a></li>
    <li><a href="http://www.ishn.org/" target="_blank" rel="noopener">International Society for the History of the Neurosciences</a></li>
  </ul>
</div>
""".strip()


def looks_like_image(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return False
    if text.startswith(("http://", "https://", "data:image/")):
        return True
    return pathlib.Path(text).suffix.lower() in IMAGE_EXTENSIONS


def to_data_url(path: str | None) -> str:
    if not path:
        return ""
    p = pathlib.Path(path)
    mime = mimetypes.guess_type(str(p))[0] or "image/png"
    encoded = base64.b64encode(p.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def normalize_image_ref(raw: Any, base_dir: pathlib.Path) -> str | None:
    if not looks_like_image(raw):
        return None
    text = str(raw).strip()
    if text.startswith(("http://", "https://", "data:image/")):
        return text
    candidate = pathlib.Path(text)
    if not candidate.is_absolute():
        candidate = (base_dir / candidate).resolve()
    if candidate.exists():
        return to_data_url(str(candidate))
    return text


def parse_year_value(raw: Any) -> tuple[int | None, str]:
    if pd.isna(raw):
        return None, ""

    if isinstance(raw, (int, float)) and not pd.isna(raw):
        numeric = int(raw)
        return numeric, str(numeric)

    text = str(raw).strip()
    if not text:
        return None, ""

    simple_int = re.fullmatch(r"-?\d+", text)
    if simple_int:
        value = int(text)
        return value, text

    matches = [int(m.group()) for m in YEAR_RE.finditer(text)]
    if matches:
        return matches[0], text

    return None, text


def load_events(path: str) -> list[dict[str, Any]]:
    src = pathlib.Path(path)
    if src.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(src, header=None)
    else:
        df = pd.read_csv(src, header=None)

    df = df.iloc[:, :4].copy()
    while df.shape[1] < 4:
        df[df.shape[1]] = None
    df.columns = ["Era", "YearRaw", "Text", "ImageRaw"]

    if not df.empty:
        first_row = [str(x).strip().lower() for x in df.iloc[0].tolist()[:4]]
        if first_row[:3] in (["era", "year", "event"], ["era", "year", "text"]):
            df = df.iloc[1:].copy()

    # Carry forward blank era cells to match grouped spreadsheet sections.
    df["Era"] = df["Era"].replace(r"^\s*$", pd.NA, regex=True).ffill()
    df["Era"] = df["Era"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    df["Text"] = df["Text"].apply(lambda x: "" if pd.isna(x) else str(x).strip())

    parsed = df["YearRaw"].apply(parse_year_value)
    df["SortYear"] = parsed.apply(lambda x: x[0])
    df["YearLabel"] = parsed.apply(lambda x: x[1])
    df["Image"] = df["ImageRaw"].apply(lambda v: normalize_image_ref(v, src.parent))

    df = df[(df["Era"] != "") & (df["Text"] != "") & df["SortYear"].notna()].copy()
    df["SortYear"] = df["SortYear"].astype(int)
    df = df.sort_values(by=["SortYear", "Era", "Text"], ascending=[True, True, True]).reset_index(drop=True)

    return [
        {
            "Era": row["Era"],
            "SortYear": int(row["SortYear"]),
            "YearLabel": row["YearLabel"],
            "Text": row["Text"],
            "Image": row["Image"],
        }
        for _, row in df.iterrows()
    ]


def build_html(events: list[dict[str, Any]], bg_url: str = "") -> str:
    data_json = json.dumps(events, ensure_ascii=False).replace("</", "<\\/")
    page_title = "neurohistory mix-tape"
    bg_css = (
        f"background-image:url('{bg_url}'); background-size:cover; background-attachment:fixed;"
        "background-position:center top;"
        if bg_url
        else "background: radial-gradient(circle at top, #16213a 0%, #0b1220 42%, #060b14 100%);"
    )

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{html.escape(page_title)}</title>
<style>
:root {{
  --page-bg: #09101d;
  --panel: rgba(10, 18, 32, 0.82);
  --panel-strong: rgba(10, 18, 32, 0.94);
  --card: rgba(16, 24, 40, 0.95);
  --text: #e7eefb;
  --muted: #a8b6d2;
  --line: rgba(109, 155, 255, 0.22);
  --accent: #73a5ff;
  --accent-2: #64e1c6;
  --mark: rgba(255, 214, 102, 0.35);
  --radius: 18px;
  --shadow: 0 24px 60px rgba(0, 0, 0, 0.28);
}}
* {{ box-sizing: border-box; }}
html, body {{ height: 100%; }}
body {{
  margin: 0;
  color: var(--text);
  font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  {bg_css}
}}
body::before {{
  content: "";
  position: fixed;
  inset: 0;
  background:
    radial-gradient(1000px 500px at 80% 10%, rgba(115,165,255,.12), transparent 60%),
    radial-gradient(800px 400px at 10% 0%, rgba(100,225,198,.09), transparent 55%),
    linear-gradient(180deg, rgba(0,0,0,.6), rgba(0,0,0,.8));
  pointer-events: none;
}}
.app {{ position: relative; min-height: 100vh; }}
.container {{ max-width: 1180px; margin: 0 auto; padding: 0 22px 42px; }}
.banner {{
  position: sticky; top: 0; z-index: 20;
  backdrop-filter: blur(16px) saturate(1.2);
  background: linear-gradient(180deg, rgba(5,10,18,.88), rgba(5,10,18,.72));
  border-bottom: 1px solid rgba(255,255,255,0.06);
}}
.banner-inner {{ padding-top: 20px; padding-bottom: 18px; }}
.kicker {{ color: var(--accent-2); font-size: 12px; letter-spacing: .18em; text-transform: uppercase; margin-bottom: 10px; }}
.banner h1 {{ margin: 0; font-size: clamp(1.5rem, 2vw + 1rem, 2.5rem); line-height: 1.1; }}
.controls {{
  display: grid; grid-template-columns: minmax(240px, 1.2fr) repeat(4, auto); gap: 12px;
  align-items: end; margin-top: 18px;
}}
.control {{ display: grid; gap: 6px; font-size: 13px; color: var(--muted); }}
input[type=search], select, button {{
  background: var(--panel);
  color: var(--text);
  border: 1px solid rgba(255,255,255,.08);
  border-radius: 12px;
  min-height: 42px;
  padding: 10px 12px;
  font: inherit;
}}
button {{ cursor: pointer; transition: transform .12s ease, border-color .12s ease, background .12s ease; }}
button:hover {{ transform: translateY(-1px); border-color: rgba(115,165,255,.5); background: rgba(18,28,48,.92); }}
button.secondary {{ min-width: 132px; }}
.summary-row {{ display: flex; flex-wrap: wrap; gap: 10px; align-items: center; margin-top: 14px; }}
.badge {{
  display: inline-flex; align-items: center; gap: 8px;
  background: rgba(255,255,255,.05); color: var(--text);
  border: 1px solid rgba(255,255,255,.08); border-radius: 999px; padding: 8px 12px; font-size: 12px;
}}
.main {{ padding-top: 26px; }}
.header-card {{
  max-width: 980px;
  margin: 0 auto 24px auto;
  padding: 24px 28px;
  background: rgba(9,14,24,.76);
  border:1px solid rgba(255,255,255,.08);
  border-radius:24px;
  color: var(--text);
  border-radius: 18px;
  box-shadow: 0 18px 40px rgba(0,0,0,0.18);
}}
.header-card h2, .header-card h3 {{ color: var(--accent-2); }}
.header-card p {{ line-height: 1.7; }}
.header-card a {{ color: #2563eb; text-decoration: none; }}
.header-card a:hover {{ text-decoration: underline; }}
.timeline-wrap {{ position: relative; padding-left: 38px; }}
.vertical-line {{
  position: absolute; left: 12px; top: 4px; bottom: 0; width: 2px;
  background: linear-gradient(180deg, rgba(115,165,255,.4), rgba(115,165,255,.1));
}}
.era-block {{
  position: relative;
  margin: 0 0 18px 0;
  border-radius: 24px;
  overflow: hidden;
  border: 1px solid rgba(255,255,255,.08);
  box-shadow: var(--shadow);
}}
.era-block::before {{
  content: "";
  position: absolute; left: -33px; top: 28px; width: 14px; height: 14px;
  border-radius: 50%; background: var(--accent); box-shadow: 0 0 0 6px rgba(115,165,255,.15);
}}
.era-header {{
  display: flex; justify-content: space-between; align-items: center; gap: 14px;
  padding: 18px 20px; cursor: pointer; user-select: none;
}}
.era-heading {{ display: grid; gap: 6px; }}
.era-title {{ font-size: 1.06rem; font-weight: 700; }}
.era-meta {{ color: rgba(231,238,251,.82); font-size: 13px; }}
.era-toggle {{
  background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.10);
  padding: 7px 12px; border-radius: 999px; text-transform: uppercase; letter-spacing: .08em; font-size: 11px;
}}
.era-content {{ padding: 0 18px 18px 20px; }}
.year-group {{
  position: relative; margin-top: 12px; padding-left: 18px;
  border-left: 2px solid rgba(255,255,255,.08);
}}
.year-header {{
  display: inline-flex; align-items: center; gap: 8px;
  padding: 4px 0; color: var(--accent-2); font-weight: 700; cursor: pointer;
}}
.event-list {{ margin-top: 8px; }}
.event-card {{
  background: var(--card);
  border: 1px solid rgba(255,255,255,.08);
  border-radius: var(--radius);
  padding: 14px 15px;
  margin: 0 0 12px 0;
}}
.event-topline {{ display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 10px; }}
.year-chip {{
  display: inline-flex; align-items: center;
  background: rgba(255,255,255,.06);
  color: var(--text); border: 1px solid rgba(255,255,255,.07);
  border-radius: 999px; padding: 5px 9px; font-size: 12px;
}}
.event-body {{ color: #eef4ff; line-height: 1.6; white-space: normal; overflow-wrap: anywhere; }}
.event-body p {{ margin: 0 0 10px 0; }}
.event-body p:last-child {{ margin-bottom: 0; }}
.event-body code {{ background: rgba(255,255,255,.08); padding: 1px 5px; border-radius: 6px; }}
.event-body a {{ color: var(--accent-2); }}
.event-body ul {{ margin: 8px 0 8px 20px; padding: 0; }}
.event-media img {{ display: block; max-width: 100%; height: auto; margin-top: 12px; border-radius: 14px; border: 1px solid rgba(255,255,255,.08); }}
mark {{ background: var(--mark); color: inherit; padding: 0 .12em; border-radius: .18em; }}
.hidden {{ display: none !important; }}
.empty-state {{
  background: var(--panel-strong); border: 1px solid rgba(255,255,255,.08); border-radius: 24px;
  padding: 26px; color: var(--muted);
}}
.footer-note {{ color: var(--muted); font-size: 12px; text-align: center; padding: 16px 0 0; }}
@media (max-width: 980px) {{
  .controls {{ grid-template-columns: 1fr 1fr; }}
}}
@media (max-width: 700px) {{
  .controls {{ grid-template-columns: 1fr; }}
  .timeline-wrap {{ padding-left: 28px; }}
  .vertical-line {{ left: 8px; }}
  .era-block::before {{ left: -23px; }}
  .era-header {{ align-items: flex-start; flex-direction: column; }}
}}
@media print {{
  .banner, .footer-note {{ position: static; }}
  .controls button {{ display: none; }}
  body {{ background: white !important; color: black; }}
  body::before {{ display: none; }}
  .event-card, .era-block, .empty-state {{ box-shadow: none; }}
}}
</style>
</head>
<body>
<div class="app">
  <div class="banner">
    <div class="container banner-inner">
      <div class="kicker">an interactive vertical timeline by Jason Keller:</div>
      <h1>{html.escape(page_title)}</h1>
      <div class="controls">
        <label class="control">Search events
          <input id="search" type="search" placeholder="Search event text…" />
        </label>
        <label class="control">Zoom
          <select id="zoom">
            <option value="era" selected>Era view</option>
            <option value="year">Year view</option>
            <option value="event">Event view</option>
          </select>
        </label>
        <label class="control">Era filter
          <select id="era-filter">
            <option value="all" selected>All eras</option>
          </select>
        </label>
        <label class="control">Expand / collapse
          <button id="expand-all" type="button" class="secondary">Expand all</button>
        </label>
        <label class="control">Export
          <button id="export-html" type="button" class="secondary">Print / Save PDF</button>
        </label>
      </div>
      <div class="summary-row">
        <span class="badge" id="count-badge">0 events</span>
        <span class="badge" id="range-badge">No range</span>
      </div>
    </div>
  </div>

  <main class="container main">
    {INTRO_HTML}
    <div class="timeline-wrap">
      <div class="vertical-line"></div>
      <div id="timeline"></div>
    </div>
    <div class="footer-note">Tip: click era or year headers to collapse and expand sections.</div>
  </main>
</div>

<script>
const DATA = {data_json};
const ERA_COLORS = ['#16315f', '#17495d', '#293462', '#4c2a85', '#244466', '#4e3a24', '#374151', '#0f4c5c'];
const timeline = document.getElementById('timeline');
const countBadge = document.getElementById('count-badge');
const rangeBadge = document.getElementById('range-badge');
const zoomSel = document.getElementById('zoom');
const eraFilter = document.getElementById('era-filter');
const searchEl = document.getElementById('search');
const expandAllBtn = document.getElementById('expand-all');

const uniqueEras = [...new Set(DATA.map(d => d.Era))];
uniqueEras.forEach((era) => {{
  const opt = document.createElement('option');
  opt.value = era;
  opt.textContent = era;
  eraFilter.appendChild(opt);
}});

function escapeHtml(text) {{
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}}

function escapeRegex(text) {{
  const specials = '.^$*+?()[]{{}}|\\\\';
  let out = '';
  for (const ch of String(text)) {{
    out += specials.includes(ch) ? '\\\\' + ch : ch;
  }}
  return out;
}}

function markdownToHtml(text) {{
  if (!text) return '';
  const NL = String.fromCharCode(10);
  const CR = String.fromCharCode(13);
  let safe = escapeHtml(text).split(CR + NL).join(NL).split(CR).join(NL);

  safe = safe.replace(/\\[([^\\]]+)\\]\\((https?:[^\\s)]+)\\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');
  safe = safe.replace(/`([^`]+)`/g, '<code>$1</code>');
  safe = safe.replace(/\\*\\*([^*]+)\\*\\*/g, '<strong>$1</strong>');
  safe = safe.replace(/\\*([^*]+)\\*/g, '<em>$1</em>');

  const lines = safe.split(NL);
  const out = [];
  let inList = false;
  for (const line of lines) {{
    const bullet = line.match(/^\\s*[-*]\\s+(.+)$/);
    if (bullet) {{
      if (!inList) {{ out.push('<ul>'); inList = true; }}
      out.push('<li>' + bullet[1] + '</li>');
      continue;
    }}
    if (inList) {{ out.push('</ul>'); inList = false; }}
    if (line.trim()) out.push('<p>' + line + '</p>');
  }}
  if (inList) out.push('</ul>');
  return out.join('');
}}

function highlightHtml(htmlString, query) {{
  const q = (query || '').trim();
  if (!q) return htmlString;
  const re = new RegExp('(' + escapeRegex(q) + ')', 'ig');
  return htmlString.replace(re, '<mark>$1</mark>');
}}

function getFilteredEvents(filterEra, query) {{
  const q = (query || '').trim().toLowerCase();
  return DATA.filter((ev) => {{
    const eraOk = filterEra === 'all' || ev.Era === filterEra;
    const haystack = (ev.Text + ' ' + ev.YearLabel + ' ' + ev.Era).toLowerCase();
    return eraOk && (!q || haystack.includes(q));
  }});
}}

function groupByEraAndYear(events) {{
  const byEra = new Map();
  for (const ev of events) {{
    if (!byEra.has(ev.Era)) byEra.set(ev.Era, []);
    byEra.get(ev.Era).push(ev);
  }}
  for (const arr of byEra.values()) {{
    arr.sort((a, b) => a.SortYear - b.SortYear || a.YearLabel.localeCompare(b.YearLabel) || a.Text.localeCompare(b.Text));
  }}
  return byEra;
}}

function setYearGroupOpen(groupEl, open) {{
  const header = groupEl.querySelector('.year-header');
  const list = groupEl.querySelector('.event-list');
  list.style.display = open ? 'block' : 'none';
  header.dataset.open = open ? '1' : '0';
  const label = header.dataset.label || '';
  header.textContent = (open ? '▾ ' : '▸ ') + label;
}}

function setEraOpen(blockEl, open) {{
  const content = blockEl.querySelector('.era-content');
  const toggle = blockEl.querySelector('.era-toggle');
  content.classList.toggle('hidden', !open);
  toggle.textContent = open ? 'collapse' : 'expand';
  blockEl.dataset.open = open ? '1' : '0';
}}

function render() {{
  const zoom = zoomSel.value;
  const filterEra = eraFilter.value;
  const query = searchEl.value;
  const events = getFilteredEvents(filterEra, query);
  const groups = groupByEraAndYear(events);

  timeline.innerHTML = '';

  if (!events.length) {{
    countBadge.textContent = '0 events';
    rangeBadge.textContent = 'No matching years';
    timeline.innerHTML = '<section class="empty-state">No events match the current filters. Try clearing the search or switching eras.</section>';
    return;
  }}

  const years = events.map(ev => ev.SortYear);
  countBadge.textContent = events.length + (events.length === 1 ? ' event' : ' events');
  rangeBadge.textContent = Math.min(...years) + ' to ' + Math.max(...years);

  let eraIndex = 0;
  for (const [era, items] of groups.entries()) {{
    const block = document.createElement('section');
    block.className = 'era-block';
    const yearsSorted = items.map(i => i.SortYear);
    const color = ERA_COLORS[eraIndex % ERA_COLORS.length];
    eraIndex += 1;
    block.style.background = 'linear-gradient(180deg, ' + color + 'E6, ' + color + 'B8)';

    const minYear = Math.min(...yearsSorted);
    const maxYear = Math.max(...yearsSorted);
    block.innerHTML = [
      '<div class="era-header">',
        '<div class="era-heading">',
          '<div class="era-title"></div>',
          '<div class="era-meta"></div>',
        '</div>',
        '<div class="era-toggle">expand</div>',
      '</div>',
      '<div class="era-content"></div>'
    ].join('');
    block.querySelector('.era-title').textContent = era;
    block.querySelector('.era-meta').textContent = minYear + ' to ' + maxYear + ' • ' + items.length + (items.length === 1 ? ' event' : ' events');

    const content = block.querySelector('.era-content');
    const byYear = new Map();
    for (const ev of items) {{
      const key = ev.YearLabel || String(ev.SortYear);
      if (!byYear.has(key)) byYear.set(key, []);
      byYear.get(key).push(ev);
    }}

    const sortedYearEntries = [...byYear.entries()].sort((a, b) => a[1][0].SortYear - b[1][0].SortYear || a[0].localeCompare(b[0]));
    for (const [yearLabel, yearEvents] of sortedYearEntries) {{
      const yg = document.createElement('div');
      yg.className = 'year-group';
      yg.innerHTML = '<div class="year-header"></div><div class="event-list"></div>';
      const yearHeader = yg.querySelector('.year-header');
      yearHeader.dataset.label = yearLabel;
      const listEl = yg.querySelector('.event-list');

      for (const ev of yearEvents) {{
        const card = document.createElement('article');
        card.className = 'event-card';
        const bodyHtml = highlightHtml(markdownToHtml(ev.Text || ''), query);
        card.innerHTML = [
          '<div class="event-topline"><span class="year-chip"></span></div>',
          '<div class="event-body"></div>',
          ev.Image ? '<div class="event-media"><img loading="lazy" alt="Event image" /></div>' : ''
        ].join('');
        card.querySelector('.year-chip').textContent = ev.YearLabel || String(ev.SortYear);
        card.querySelector('.event-body').innerHTML = bodyHtml;
        if (ev.Image) {{
          card.querySelector('img').src = ev.Image;
        }}
        listEl.appendChild(card);
      }}

      yearHeader.addEventListener('click', () => {{
        const open = yearHeader.dataset.open !== '1';
        setYearGroupOpen(yg, open);
      }});

      const openByDefault = zoom === 'event';
      setYearGroupOpen(yg, openByDefault);
      content.appendChild(yg);
    }}

    block.querySelector('.era-header').addEventListener('click', () => {{
      const open = block.dataset.open !== '1';
      setEraOpen(block, open);
    }});

    const eraOpenByDefault = zoom !== 'era';
    setEraOpen(block, eraOpenByDefault);
    timeline.appendChild(block);
  }}
}}

expandAllBtn.addEventListener('click', () => {{
  const blocks = [...document.querySelectorAll('.era-block')];
  const anyCollapsed = blocks.some(block => block.dataset.open !== '1');
  blocks.forEach(block => {{
    setEraOpen(block, anyCollapsed);
    block.querySelectorAll('.year-group').forEach(group => setYearGroupOpen(group, anyCollapsed));
  }});
  expandAllBtn.textContent = anyCollapsed ? 'Collapse all' : 'Expand all';
}});

document.getElementById('export-html').addEventListener('click', () => window.print());
zoomSel.addEventListener('change', () => {{
  expandAllBtn.textContent = 'Expand all';
  render();
}});
eraFilter.addEventListener('change', () => {{
  expandAllBtn.textContent = 'Expand all';
  render();
}});
searchEl.addEventListener('input', () => {{
  expandAllBtn.textContent = 'Expand all';
  render();
}});

render();
</script>
</body>
</html>
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Build an interactive HTML vertical timeline from a spreadsheet.")
    parser.add_argument("input_path", help="Input spreadsheet (.xlsx/.xls/.csv)")
    parser.add_argument("output_path", help="Output HTML path")
    parser.add_argument("--bg", dest="bg_path", default="", help="Optional background image path")
    args = parser.parse_args()

    events = load_events(args.input_path)
    bg_url = to_data_url(args.bg_path) if args.bg_path else ""
    html_text = build_html(events, bg_url=bg_url)
    pathlib.Path(args.output_path).write_text(html_text, encoding="utf-8")
    print(f"Wrote {args.output_path} with {len(events)} events")


if __name__ == "__main__":
    main()
