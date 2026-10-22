#!/usr/bin/env python3
"""
xlsx → md/html 変換スクリプト
リポジトリルートで `python scripts/xlsx_to_mdhtml.py` 実行
"""
import os
import sys
import html
import re
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[1]
XLSX_PATH = REPO_ROOT / "master" / "ゲーム業界・時事情報収集.xlsx"
MD_DIR = REPO_ROOT / "md"
DOCS_DIR = REPO_ROOT / "docs"

SHEET_SLUGS = {
    "README": "00-readme",
    "日次ニュース": "01-daily-news",
    "新作リリースログ": "02-releases",
    "IP・タイトルウォッチ": "03-ip-watch",
    "企業分析（国内）": "04-companies-jp",
    "企業分析（海外大手）": "05-companies-overseas",
    "四半期決算サマリ": "06-earnings",
    "業界トレンド・技術": "07-trends",
    "インディー注目": "08-indie",
    "今後起きそうなこと": "09-future",
    "掘り下げ・分析": "10-deep-dive",
    "業界マトリクス": "11-matrix",
    "横断的問い": "12-cross-questions",
    "週間ランキング": "13-rankings",
    "イベントカレンダー": "14-events",
    "ゲームデザイン・トレンド": "15-game-design",
    "デザイン派生系統": "16-design-lineage",
    "メディアミックス追跡": "17-media-mix",
    "【映画】作品ウォッチ": "20-movie-works",
    "【映画】企業・スタジオ": "21-movie-companies",
    "【映画】公開・興行": "22-movie-releases",
    "【映画】トレンド": "23-movie-trends",
    "【アニメ】作品ウォッチ": "30-anime-works",
    "【アニメ】企業・制作": "31-anime-companies",
    "【アニメ】放送・配信ログ": "32-anime-releases",
    "【アニメ】トレンド": "33-anime-trends",
    "【漫画】作品ウォッチ": "40-manga-works",
    "【漫画】出版社・PF": "41-manga-companies",
    "【漫画】新刊・話題作": "42-manga-releases",
    "【漫画】トレンド": "43-manga-trends",
    "【小説】作品ウォッチ": "50-novel-works",
    "【小説】出版社・レーベル": "51-novel-companies",
    "【小説】新刊・受賞": "52-novel-releases",
    "【小説】トレンド": "53-novel-trends",
    "【音楽】アーティストウォッチ": "60-music-works",
    "【音楽】企業・レーベル": "61-music-companies",
    "【音楽】リリース・チャート": "62-music-releases",
    "【音楽】トレンド": "63-music-trends",
}

# 「行ごとセクション形式」で表示するシート
# ナビ・INDEXのグループ分け
SHEET_GROUPS = [
    ("共通", ["日次ニュース", "横断的問い", "掘り下げ・分析", "今後起きそうなこと",
              "業界マトリクス", "メディアミックス追跡", "イベントカレンダー"]),
    ("ゲーム", ["新作リリースログ", "IP・タイトルウォッチ", "企業分析（国内）", "企業分析（海外大手）",
                "四半期決算サマリ", "業界トレンド・技術", "インディー注目", "週間ランキング",
                "ゲームデザイン・トレンド", "デザイン派生系統"]),
    ("映画", ["【映画】作品ウォッチ", "【映画】企業・スタジオ", "【映画】公開・興行", "【映画】トレンド"]),
    ("アニメ", ["【アニメ】作品ウォッチ", "【アニメ】企業・制作", "【アニメ】放送・配信ログ", "【アニメ】トレンド"]),
    ("漫画", ["【漫画】作品ウォッチ", "【漫画】出版社・PF", "【漫画】新刊・話題作", "【漫画】トレンド"]),
    ("小説", ["【小説】作品ウォッチ", "【小説】出版社・レーベル", "【小説】新刊・受賞", "【小説】トレンド"]),
    ("音楽", ["【音楽】アーティストウォッチ", "【音楽】企業・レーベル", "【音楽】リリース・チャート", "【音楽】トレンド"]),
]

# ナビ表示時にグループ名を落として短くする
def _short(sn):
    for g in ("映画", "アニメ", "漫画", "小説", "音楽"):
        p = f"【{g}】"
        if sn.startswith(p):
            return sn[len(p):]
    return sn

SECTION_SHEETS = {"日次ニュース", "今後起きそうなこと", "掘り下げ・分析", "横断的問い"}

CSS = """
<style>
:root{
  --bg:#0d1117; --panel:#161b22; --border:#30363d;
  --text:#c9d1d9; --muted:#8b949e; --accent:#58a6ff; --accent-2:#79c0ff;
  --header-bg:#1f3864; --header-text:#fff;
}
@media (prefers-color-scheme: light){
  :root{ --bg:#f6f8fa; --panel:#fff; --border:#d0d7de;
         --text:#1f2328; --muted:#656d76; --accent:#0969da; --accent-2:#0550ae;
         --header-bg:#1f3864; --header-text:#fff; }
}
*{box-sizing:border-box}
body{
  background:var(--bg); color:var(--text);
  font-family:-apple-system,BlinkMacSystemFont,"Helvetica Neue","Hiragino Sans","Yu Gothic","Noto Sans JP",sans-serif;
  margin:0; padding:24px; line-height:1.65;
}
.container{max-width:1400px; margin:0 auto}
header{display:flex; align-items:baseline; gap:16px; flex-wrap:wrap; margin-bottom:20px}
h1{margin:0; font-size:22px}
.subtle{color:var(--muted); font-size:13px}
nav{margin:16px 0 24px; padding:12px 16px; background:var(--panel); border:1px solid var(--border); border-radius:8px}
nav a{color:var(--accent); margin-right:14px; text-decoration:none; font-size:13px; white-space:nowrap}
nav a:hover{text-decoration:underline}
.search-wrap{margin:0 0 12px}
input.search{width:100%; padding:10px 12px; background:var(--panel); color:var(--text);
  border:1px solid var(--border); border-radius:6px; font-size:14px}
.table-wrap{overflow-x:auto; background:var(--panel); border:1px solid var(--border); border-radius:8px}
table{border-collapse:collapse; width:100%; font-size:13px}
th,td{padding:10px 12px; border-bottom:1px solid var(--border); vertical-align:top; text-align:left}
th{background:var(--header-bg); color:var(--header-text); position:sticky; top:0; font-weight:600; cursor:pointer; user-select:none}
th:hover{background:#2a4f8c}
td{white-space:pre-wrap; word-break:break-word; max-width:480px}
tr:nth-child(even){background:rgba(127,127,127,0.04)}
tr:hover{background:rgba(88,166,255,0.08)}
a{color:var(--accent)}
.section{background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:18px 22px; margin-bottom:14px}
.section h2{margin:0 0 8px; font-size:17px}
.section .meta{color:var(--muted); font-size:12px; margin-bottom:10px}
.section dl{margin:0; display:grid; grid-template-columns:max-content 1fr; gap:6px 16px; font-size:13px}
.section dt{color:var(--muted); white-space:nowrap}
.section dd{margin:0; white-space:pre-wrap}
.empty{color:var(--muted); font-style:italic; padding:30px; text-align:center}
footer{margin-top:30px; color:var(--muted); font-size:12px; text-align:center}
.tag{display:inline-block; background:rgba(88,166,255,0.15); color:var(--accent-2);
  padding:2px 8px; border-radius:10px; font-size:11px; margin-right:4px; margin-bottom:3px}
.chips{margin:2px 0 6px}
.tagbar{display:flex; flex-wrap:wrap; gap:6px; margin:0 0 14px; align-items:center}
.tagbar .tb-label{color:var(--muted); font-size:12px; margin-right:2px}
.tagfilter{background:var(--panel); color:var(--muted); border:1px solid var(--border);
  border-radius:14px; padding:4px 12px; font-size:12px; cursor:pointer; line-height:1.2}
.tagfilter:hover{border-color:var(--accent); color:var(--accent)}
.tagfilter.active{background:var(--accent); color:#fff; border-color:var(--accent)}

nav.grouped{margin:16px 0 24px; padding:6px 14px; background:var(--panel);
  border:1px solid var(--border); border-radius:8px}
nav.grouped .navrow{display:flex; align-items:baseline; gap:10px; flex-wrap:wrap;
  padding:7px 0; border-bottom:1px solid var(--border)}
nav.grouped .navrow:last-child{border-bottom:none}
nav.grouped .navlabel{flex:0 0 62px; color:var(--muted); font-size:11px; font-weight:700;
  letter-spacing:.04em}
nav.grouped a{color:var(--accent); text-decoration:none; font-size:12.5px; white-space:nowrap}
nav.grouped a:hover{text-decoration:underline}
nav.grouped a.active{color:var(--text); font-weight:700}
nav.grouped a .nvdot,nav.grouped .navleg .nvdot{color:#e8583a; font-size:9px; vertical-align:super; margin-left:2px; line-height:0}
nav.grouped .navleg{flex:0 0 auto; color:var(--muted); font-size:11px; margin-left:auto}
.genre-h{margin:22px 0 6px; font-size:13px; color:var(--muted); font-weight:700;
  letter-spacing:.04em; border-bottom:1px solid var(--border); padding-bottom:5px}
</style>
"""

JS = """
<script>
(function(){
  const inp=document.getElementById('search');
  const bar=document.getElementById('tagbar');
  const active=new Set();
  function apply(){
    const q=(inp?inp.value.toLowerCase():'');
    document.querySelectorAll('.searchable').forEach(el=>{
      const textOk = !q || el.textContent.toLowerCase().includes(q);
      let tagOk = true;
      if(active.size>0){
        const tags=(el.dataset.tags||'').split('|').filter(Boolean);
        tagOk = tags.some(t=>active.has(t));
      }
      el.style.display=(textOk&&tagOk)?'':'none';
    });
  }
  if(inp) inp.addEventListener('input', apply);
  if(bar){
    const allBtn=bar.querySelector('.tagfilter[data-tag="__all__"]');
    bar.querySelectorAll('.tagfilter').forEach(btn=>{
      btn.addEventListener('click',()=>{
        const t=btn.dataset.tag;
        if(t==='__all__'){
          active.clear();
          bar.querySelectorAll('.tagfilter').forEach(b=>b.classList.remove('active'));
          btn.classList.add('active');
        } else {
          if(active.has(t)){active.delete(t); btn.classList.remove('active');}
          else{active.add(t); btn.classList.add('active');}
          if(active.size===0){allBtn.classList.add('active');} else {allBtn.classList.remove('active');}
        }
        apply();
      });
    });
  }
  // sort
  document.querySelectorAll('table').forEach(table=>{
    table.querySelectorAll('th').forEach((th,idx)=>{
      let asc=true;
      th.addEventListener('click',()=>{
        const tbody=table.tBodies[0];
        const rows=Array.from(tbody.rows);
        rows.sort((a,b)=>{
          const av=(a.cells[idx]?.innerText||'').trim();
          const bv=(b.cells[idx]?.innerText||'').trim();
          const an=parseFloat(av.replace(/,/g,''));
          const bn=parseFloat(bv.replace(/,/g,''));
          if(!isNaN(an)&&!isNaN(bn)) return (an-bn)*(asc?1:-1);
          return av.localeCompare(bv,'ja')*(asc?1:-1);
        });
        rows.forEach(r=>tbody.appendChild(r));
        asc=!asc;
      });
    });
  });
})();
</script>
"""

_RECENT_SLUGS = set()
_RECENT_DATES = {}

def nav_html(active_slug=None):
    parts = ['<nav class="grouped">']
    _top = ('<div class="navrow"><span class="navlabel">TOP</span>'
            '<a href="index.html">ダッシュボード</a>'
            '<a href="design-map.html">デザイン連動マップ</a>'
            '<a href="industry-map.html">全体連動マップ</a>')
    if _RECENT_SLUGS:
        _top += '<span class="navleg"><span class="nvdot">●</span> 最近更新</span>'
    _top += '</div>'
    parts.append(_top)
    for gname, sheets in SHEET_GROUPS:
        links = []
        for sn in sheets:
            slug = SHEET_SLUGS.get(sn)
            if not slug:
                continue
            cls = ' class="active"' if slug == active_slug else ''
            _dot = (f'<span class="nvdot" title="最近更新: {html.escape(_RECENT_DATES.get(slug,""))}">\u25cf</span>'
                    if slug in _RECENT_SLUGS else '')
            links.append(f'<a href="{slug}.html"{cls}>{html.escape(_short(sn))}{_dot}</a>')
        if links:
            parts.append(f'<div class="navrow"><span class="navlabel">{html.escape(gname)}</span>'
                         + ''.join(links) + '</div>')
    parts.append('</nav>')
    return ''.join(parts)

def url_to_link(text):
    if not isinstance(text, str): return html.escape(str(text)) if text is not None else ''
    pattern = re.compile(r'(https?://[^\s<>"]+)')
    parts = []
    last = 0
    for m in pattern.finditer(text):
        parts.append(html.escape(text[last:m.start()]))
        url = m.group(1)
        parts.append(f'<a href="{html.escape(url)}" target="_blank" rel="noopener">{html.escape(url[:60] + ("…" if len(url)>60 else ""))}</a>')
        last = m.end()
    parts.append(html.escape(text[last:]))
    return ''.join(parts)

def fmt_cell(v):
    if v is None: return ''
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, float):
        return f"{v:,.0f}" if v == int(v) else f"{v:,}"
    return str(v)

def fmt_tag_cell(v):
    text = fmt_cell(v)
    if not text: return ''
    if any(sep in text for sep in [',', '、']):
        tags = re.split(r'[,、]\s*', text)
        return ''.join(f'<span class="tag">{html.escape(t.strip())}</span>' for t in tags if t.strip())
    return html.escape(text)

def render_table_html(sheet_name, headers, rows):
    nav = nav_html(SHEET_SLUGS.get(sheet_name))
    title = html.escape(sheet_name)
    if not rows:
        body = '<div class="empty">データなし</div>'
    else:
        ths = ''.join(f'<th>{html.escape(h or "")}</th>' for h in headers)
        trs = []
        for row in rows:
            tds = []
            for i, v in enumerate(row):
                h = headers[i] if i < len(headers) else ''
                text = fmt_cell(v)
                if not text:
                    cell_html = ''
                elif h and ('URL' in h or 'リンク' in h):
                    cell_html = url_to_link(text)
                elif h == 'タグ':
                    cell_html = fmt_tag_cell(v)
                else:
                    cell_html = url_to_link(text) if 'http' in text else html.escape(text)
                tds.append(f'<td>{cell_html}</td>')
            trs.append(f'<tr class="searchable">{"".join(tds)}</tr>')
        body = (
            '<div class="search-wrap"><input class="search" id="search" placeholder="全列を絞り込み検索…"></div>'
            '<div class="table-wrap"><table><thead><tr>'
            + ths + '</tr></thead><tbody>'
            + ''.join(trs) + '</tbody></table></div>'
        )
    updated = datetime.now().strftime('%Y-%m-%d %H:%M')
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>{title} | ゲーム業界トラッカー</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{CSS}</head><body><div class="container">
<header><h1>{title}</h1><span class="subtle">最終更新: {updated}</span></header>
{nav}{body}
<footer>ゲーム業界・時事情報収集 · 自動生成</footer>
</div>{JS}</body></html>"""

def render_section_html(sheet_name, headers, rows):
    nav = nav_html(SHEET_SLUGS.get(sheet_name))
    title = html.escape(sheet_name)
    sections = []
    all_tags = []
    for row in rows:
        d = {headers[i]: fmt_cell(v) for i, v in enumerate(row) if i < len(headers)}
        if not any(d.values()): continue
        head_field = headers[1] if len(headers) > 1 else headers[0]
        h2 = html.escape(d.get(head_field, '') or '(無題)')
        tags = [t.strip() for t in re.split(r'[,、]\s*', d.get('タグ', '') or '') if t.strip()]
        for t in tags:
            if t not in all_tags: all_tags.append(t)
        data_tags = html.escape('|'.join(tags))
        chips = ''.join(f'<span class="tag">{html.escape(t)}</span>' for t in tags)
        meta_parts = []
        if d.get('日付'): meta_parts.append(d['日付'])
        if d.get('記録日'): meta_parts.append(d['記録日'])
        if d.get('関連企業'): meta_parts.append(f"関連企業: {d['関連企業']}")
        if d.get('関連IP/タイトル') or d.get('関連企業/IP'):
            ip = d.get('関連IP/タイトル') or d.get('関連企業/IP')
            meta_parts.append(f"IP: {ip}")
        if d.get('確度'): meta_parts.append(f"確度: {d['確度']}")
        meta = ' · '.join(meta_parts)
        dls = []
        skip_keys = {'日付', '記録日', '関連企業', '関連IP/タイトル', '関連企業/IP', 'タグ', '確度'}
        if head_field in d:
            skip_keys.add(head_field)
        for k in headers:
            if not k or k in skip_keys: continue
            v = d.get(k, '')
            if not v: continue
            if 'URL' in k or 'リンク' in k:
                vh = url_to_link(v)
            else:
                vh = url_to_link(v) if 'http' in v else html.escape(v).replace('\n','<br>')
            dls.append(f'<dt>{html.escape(k)}</dt><dd>{vh}</dd>')
        sections.append(f'<div class="section searchable" data-tags="{data_tags}"><h2>{h2}</h2>'
                        + (f'<div class="meta">{html.escape(meta)}</div>' if meta else '')
                        + (f'<div class="chips">{chips}</div>' if chips else '')
                        + (f'<dl>{"".join(dls)}</dl>' if dls else '') + '</div>')
    tagbar = ''
    if all_tags:
        btns = ['<span class="tb-label">タグで絞り込み:</span>',
                '<button class="tagfilter active" data-tag="__all__">すべて</button>']
        btns += [f'<button class="tagfilter" data-tag="{html.escape(t)}">{html.escape(t)}</button>' for t in all_tags]
        tagbar = '<div class="tagbar" id="tagbar">' + ''.join(btns) + '</div>'
    body = ('<div class="search-wrap"><input class="search" id="search" placeholder="全文検索…"></div>'
            + tagbar
            + (''.join(sections) if sections else '<div class="empty">データなし</div>'))
    updated = datetime.now().strftime('%Y-%m-%d %H:%M')
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>{title} | ゲーム業界トラッカー</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{CSS}</head><body><div class="container">
<header><h1>{title}</h1><span class="subtle">最終更新: {updated}</span></header>
{nav}{body}
<footer>ゲーム業界・時事情報収集 · 自動生成</footer>
</div>{JS}</body></html>"""


DAILY_CSS = """
<style>
.dbar{position:sticky;top:0;z-index:5;background:var(--bg);padding:10px 0 6px}
.dbar .row{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:6px}
.dbar .lbl{color:var(--muted);font-size:12px;margin-right:2px;min-width:52px}
.gf{border:1px solid var(--border);background:var(--panel);color:var(--text);border-radius:999px;padding:4px 12px;font-size:12px;cursor:pointer}
.gf.active{background:var(--accent);border-color:var(--accent);color:#fff}
.gf .n{color:var(--muted);margin-left:4px;font-size:11px}
.gf.active .n{color:#fff}
.dayh{display:flex;align-items:baseline;gap:10px;margin:26px 0 10px;padding-bottom:6px;border-bottom:2px solid var(--border);scroll-margin-top:110px}
.dayh h2{margin:0;font-size:18px}
.dayh .c{color:var(--muted);font-size:12px}
.dayh .mix{display:flex;gap:4px;flex-wrap:wrap}
.nitem{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:12px 16px;margin-bottom:8px}
.nitem .top{display:flex;gap:8px;align-items:flex-start}
.nitem h3{margin:0;font-size:14.5px;line-height:1.5}
.nitem h3 a{color:var(--text);text-decoration:none}
.nitem h3 a:hover{color:var(--accent)}
.nitem .sum{font-size:13px;color:var(--text);margin:6px 0 4px;line-height:1.65}
.nitem .meta{color:var(--muted);font-size:11.5px}
.nitem details summary{cursor:pointer;color:var(--accent);font-size:12px;margin-top:6px}
.nitem dl{margin:8px 0 0;display:grid;grid-template-columns:max-content 1fr;gap:6px 14px;font-size:12.5px}
.nitem dt{color:var(--muted);white-space:nowrap}
.nitem dd{margin:0}
.gb{display:inline-block;font-size:11px;font-weight:700;padding:1px 8px;border-radius:999px;white-space:nowrap;flex:0 0 auto;margin-top:3px}
.g0{background:rgba(124,92,255,.18);color:#a08cff}.g1{background:rgba(229,83,75,.18);color:#ff7b72}
.g2{background:rgba(240,136,62,.18);color:#f0883e}.g3{background:rgba(46,160,67,.18);color:#3fb950}
.g4{background:rgba(187,128,9,.2);color:#d29922}.g5{background:rgba(219,97,162,.18);color:#f778ba}
.g6{background:rgba(88,166,255,.18);color:#79c0ff}
@media (prefers-color-scheme: light){.g0{color:#6639ba}.g1{color:#c93c37}.g2{color:#bc4c00}.g3{color:#1a7f37}.g4{color:#9a6700}.g5{color:#bf3989}.g6{color:#0969da}}
.dnav{display:flex;gap:6px;overflow-x:auto;padding-bottom:4px}
.dnav a{font-size:11.5px;white-space:nowrap;color:var(--accent);text-decoration:none;border:1px solid var(--border);border-radius:6px;padding:2px 8px}
.more{display:block;margin:18px auto;padding:9px 18px;border:1px solid var(--border);background:var(--panel);color:var(--text);border-radius:999px;cursor:pointer}
</style>
"""

DAILY_JS = """
<script>
(function(){
  const inp=document.getElementById('search');
  const gsel=new Set(), tsel=new Set();
  let limitDays=14, filtering=false;
  const days=[...document.querySelectorAll('.dayblock')];
  function apply(){
    const q=(inp.value||'').toLowerCase();
    filtering = !!q || gsel.size>0 || tsel.size>0;
    days.forEach((b,i)=>{
      let vis=0;
      b.querySelectorAll('.nitem').forEach(el=>{
        const okq=!q||el.textContent.toLowerCase().includes(q);
        const okg=!gsel.size||gsel.has(el.dataset.g);
        const tags=(el.dataset.tags||'').split('|');
        const okt=!tsel.size||tags.some(t=>tsel.has(t));
        const ok=okq&&okg&&okt; el.style.display=ok?'':'none'; if(ok)vis++;
      });
      const within = filtering || i<limitDays;
      b.style.display=(vis&&within)?'':'none';
      const c=b.querySelector('.vis'); if(c) c.textContent = filtering ? ('（該当 '+vis+'件）') : '';
    });
    document.getElementById('more').style.display = (!filtering && limitDays<days.length)?'':'none';
  }
  function toggle(set,btn,key){
    if(key==='__all__'){set.clear();btn.parentNode.querySelectorAll('.gf').forEach(b=>b.classList.remove('active'));btn.classList.add('active');}
    else{ if(set.has(key)){set.delete(key);btn.classList.remove('active');}else{set.add(key);btn.classList.add('active');}
      btn.parentNode.querySelector('[data-k="__all__"]').classList.toggle('active',set.size===0);}
    apply();
  }
  document.querySelectorAll('#gfil .gf').forEach(b=>b.onclick=()=>toggle(gsel,b,b.dataset.k));
  document.querySelectorAll('#tfil .gf').forEach(b=>b.onclick=()=>toggle(tsel,b,b.dataset.k));
  inp.addEventListener('input',apply);
  document.getElementById('more').onclick=()=>{limitDays+=14;apply();};
  // ハッシュで古い日付に飛んだときも表示する
  function hash(){const h=location.hash.slice(1);const i=days.findIndex(b=>b.id===h);
    if(i>=limitDays){limitDays=i+1;apply();document.getElementById(h).scrollIntoView();}}
  window.addEventListener('hashchange',hash);
  apply(); hash();
})();
</script>
"""

_GENRES = ["ゲーム", "映画", "アニメ", "漫画", "小説", "音楽", "横断"]
_WD = "月火水木金土日"

def _gb(g):
    g = g or "横断"
    i = _GENRES.index(g) if g in _GENRES else 6
    return f'<span class="gb g{i}">{html.escape(g)}</span>'

def render_daily_html(sheet_name, headers, rows):
    from collections import OrderedDict, Counter
    nav = nav_html(SHEET_SLUGS.get(sheet_name))
    days = OrderedDict()
    for row in rows:
        d = {headers[i]: fmt_cell(v) for i, v in enumerate(row) if i < len(headers)}
        if not any(d.values()):
            continue
        m = re.search(r'\d{4}-\d{2}-\d{2}', d.get('日付', '') or '')
        days.setdefault(m.group(0) if m else '日付なし', []).append(d)
    gcount = Counter(); tcount = Counter()
    blocks = []
    dnav = []
    for day, items in days.items():
        try:
            dt = datetime.strptime(day, '%Y-%m-%d'); lab = f"{dt.year}年{dt.month}月{dt.day}日（{_WD[dt.weekday()]}）"
        except Exception:
            lab = day
        gc = Counter((x.get('ジャンル') or '横断') for x in items)
        mix = ''.join(f'{_gb(g)}<span class="c">{gc[g]}</span>' for g in _GENRES if gc.get(g))
        its = []
        for x in items:
            g = x.get('ジャンル') or '横断'
            gcount[g] += 1
            tags = [t.strip() for t in re.split(r'[,、]\s*', x.get('タグ', '') or '') if t.strip()]
            for t in tags: tcount[t] += 1
            chips = ''.join(f'<span class="tag">{html.escape(t)}</span>' for t in tags)
            url = x.get('URL', '')
            title = html.escape(x.get('見出し', '') or '(無題)')
            th = f'<a href="{html.escape(url)}" target="_blank" rel="noopener">{title}</a>' if url else title
            meta = ' · '.join(p for p in [
                f"関連企業: {x['関連企業']}" if x.get('関連企業') else '',
                f"IP: {x['関連IP/タイトル']}" if x.get('関連IP/タイトル') else ''] if p)
            dls = []
            for k in ('背景にある出来事', '重要ポイント', '今後の予測', '勉強しておくべきテーマ', '備考'):
                v = x.get(k, '')
                if v:
                    vh = html.escape(v).replace('\n', '<br>')
                    if k == '重要ポイント':
                        vh = '<ul style="margin:0;padding-left:18px">' + ''.join(
                            f'<li>{html.escape(l.lstrip("・- "))}</li>' for l in v.split('\n') if l.strip()) + '</ul>'
                    dls.append(f'<dt>{html.escape(k)}</dt><dd>{vh}</dd>')
            its.append(f'<div class="nitem" data-g="{html.escape(g)}" data-tags="{html.escape("|".join(tags))}">'
                       f'<div class="top">{_gb(g)}<h3>{th}</h3></div>'
                       + (f'<div class="sum">{html.escape(x.get("サマリ", ""))}</div>' if x.get('サマリ') else '')
                       + f'<div class="chips">{chips}</div>'
                       + (f'<div class="meta">{html.escape(meta)}</div>' if meta else '')
                       + (f'<details><summary>背景・ポイント・予測を開く</summary><dl>{"".join(dls)}</dl></details>' if dls else '')
                       + '</div>')
        blocks.append(f'<div class="dayblock" id="d-{day}"><div class="dayh"><h2>{lab}</h2>'
                      f'<span class="c">{len(items)}件</span><span class="mix">{mix}</span><span class="c vis"></span></div>'
                      + ''.join(its) + '</div>')
        if len(dnav) < 21:
            dnav.append(f'<a href="#d-{day}">{day[5:].replace("-", "/")}</a>')
    gbtn = ['<span class="lbl">ジャンル</span><button class="gf active" data-k="__all__">すべて</button>']
    gbtn += [f'<button class="gf" data-k="{g}">{g}<span class="n">{gcount[g]}</span></button>' for g in _GENRES if gcount.get(g)]
    tbtn = ['<span class="lbl">タグ</span><button class="gf active" data-k="__all__">すべて</button>']
    tbtn += [f'<button class="gf" data-k="{html.escape(t)}">{html.escape(t)}<span class="n">{n}</span></button>'
             for t, n in tcount.most_common()]
    bar = (f'<div class="dbar"><input class="search" id="search" placeholder="見出し・本文・企業名で検索…">'
           f'<div class="row" id="gfil" style="margin-top:8px">{"".join(gbtn)}</div>'
           f'<details><summary class="subtle" style="cursor:pointer">タグで絞り込む</summary><div class="row" id="tfil">{"".join(tbtn)}</div></details>'
           f'<div class="dnav">{"".join(dnav)}</div></div>')
    updated = datetime.now().strftime('%Y-%m-%d %H:%M')
    total = sum(len(v) for v in days.values())
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>日次ニュース | エンタメ業界トラッカー</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{CSS}{DAILY_CSS}</head><body><div class="container">
<header><h1>日次ニュース</h1><span class="subtle">全{total:,}件 · {len(days)}日分 · 最終更新: {updated}</span></header>
{nav}{bar}{''.join(blocks)}
<button class="more" id="more">さらに古い日付を表示</button>
<footer>エンタメ業界・時事情報収集 · 自動生成</footer>
</div>{DAILY_JS}</body></html>"""


RANKING_SHEET = "週間ランキング"
RANK_GROUP_ORDER = ["国内ファミ通", "Steamトップセラー", "NSW DL", "PS DL"]

def _week_key(s):
    m = re.search(r'(\d{4})-(\d{2})(?:-(\d{2}))?', str(s or ''))
    return m.group(0) if m else ''

def _rank_groups(headers, rows):
    idx = {h: i for i, h in enumerate(headers) if h}
    ki, wi, ri = idx.get('区分'), idx.get('集計週'), idx.get('順位')
    groups = {}
    for r in rows:
        g = fmt_cell(r[ki]) if ki is not None else 'その他'
        groups.setdefault(g or 'その他', []).append(r)
    order = [g for g in RANK_GROUP_ORDER if g in groups] + [g for g in groups if g not in RANK_GROUP_ORDER]
    def rank_val(r):
        try: return int(str(r[ri]).replace(',', ''))
        except Exception: return 999
    for g in groups:
        groups[g].sort(key=rank_val)
        groups[g].sort(key=lambda r: _week_key(fmt_cell(r[wi])), reverse=True)
    return idx, wi, order, groups

def render_ranking_html(sheet_name, headers, rows):
    nav = nav_html(SHEET_SLUGS.get(sheet_name))
    title = html.escape(sheet_name)
    idx, wi, order, groups = _rank_groups(headers, rows)
    disp = [h for h in headers if h and h not in ('区分', '集計週')]
    sections = []
    for gi, g in enumerate(order):
        grows = groups[g]
        weeks = sorted({fmt_cell(r[wi]) for r in grows}, key=_week_key, reverse=True)
        opts = ''.join(
            f'<option value="{html.escape(w)}"{" selected" if i == 0 else ""}>{html.escape(w)}</option>'
            for i, w in enumerate(weeks))
        opts += '<option value="__all__">すべての週</option>'
        ths = ''.join(f'<th>{html.escape(h)}</th>' for h in disp)
        trs = []
        for r in grows:
            week = fmt_cell(r[wi])
            tds = []
            for h in disp:
                v = r[idx[h]] if idx.get(h) is not None and idx[h] < len(r) else None
                text = fmt_cell(v)
                if not text:
                    cell = ''
                elif 'URL' in h or 'リンク' in h or '出典' in h:
                    cell = url_to_link(text)
                else:
                    cell = url_to_link(text) if 'http' in text else html.escape(text)
                tds.append(f'<td>{cell}</td>')
            trs.append(f'<tr data-group="g{gi}" data-week="{html.escape(week)}">{"".join(tds)}</tr>')
        sections.append(
            f'<div class="section"><h2>{html.escape(g)}</h2>'
            f'<div class="meta" style="margin-bottom:10px">対象週: '
            f'<select class="week-sel" data-group="g{gi}" '
            f'style="padding:6px 10px;background:var(--panel);color:var(--text);'
            f'border:1px solid var(--border);border-radius:6px;font-size:13px">{opts}</select></div>'
            f'<div class="table-wrap"><table><thead><tr>{ths}</tr></thead>'
            f'<tbody>{"".join(trs)}</tbody></table></div></div>')
    body = ''.join(sections) if sections else '<div class="empty">データなし</div>'
    week_js = """
<script>
document.querySelectorAll('select.week-sel').forEach(sel=>{
  const apply=()=>{
    const v=sel.value;
    document.querySelectorAll('tr[data-group=\"'+sel.dataset.group+'\"]').forEach(tr=>{
      tr.style.display=(v==='__all__'||tr.dataset.week===v)?'':'none';
    });
  };
  sel.addEventListener('change',apply); apply();
});
</script>"""
    updated = datetime.now().strftime('%Y-%m-%d %H:%M')
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>{title} | ゲーム業界トラッカー</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{CSS}</head><body><div class="container">
<header><h1>{title}</h1><span class="subtle">最終更新: {updated}</span></header>
{nav}{body}
<footer>ゲーム業界・時事情報収集 · 自動生成</footer>
</div>{JS}{week_js}</body></html>"""

def render_ranking_md(sheet_name, headers, rows):
    lines = [f"# {sheet_name}", "", f"_最終更新: {datetime.now().strftime('%Y-%m-%d %H:%M')}_", ""]
    idx, wi, order, groups = _rank_groups(headers, rows)
    disp = [h for h in headers if h and h not in ('区分', '集計週')]
    for g in order:
        lines.append(f"## {g}")
        lines.append("")
        weeks = sorted({fmt_cell(r[wi]) for r in groups[g]}, key=_week_key, reverse=True)
        for w in weeks:
            lines.append(f"### {w}")
            lines.append("")
            lines.append('| ' + ' | '.join(disp) + ' |')
            lines.append('| ' + ' | '.join(['---'] * len(disp)) + ' |')
            for r in groups[g]:
                if fmt_cell(r[wi]) != w: continue
                cells = []
                for h in disp:
                    v = r[idx[h]] if idx.get(h) is not None and idx[h] < len(r) else None
                    cells.append(fmt_cell(v).replace('|', '\\|').replace('\n', '<br>'))
                lines.append('| ' + ' | '.join(cells) + ' |')
            lines.append("")
    return '\n'.join(lines) + '\n'

def render_table_md(sheet_name, headers, rows):
    lines = [f"# {sheet_name}", "", f"_最終更新: {datetime.now().strftime('%Y-%m-%d %H:%M')}_", ""]
    safe_headers = [h or '' for h in headers]
    lines.append('| ' + ' | '.join(safe_headers) + ' |')
    lines.append('| ' + ' | '.join(['---'] * len(safe_headers)) + ' |')
    for row in rows:
        cells = []
        for v in row:
            t = fmt_cell(v).replace('|', '\\|').replace('\n', '<br>')
            cells.append(t)
        lines.append('| ' + ' | '.join(cells) + ' |')
    return '\n'.join(lines) + '\n'

def render_section_md(sheet_name, headers, rows):
    lines = [f"# {sheet_name}", "", f"_最終更新: {datetime.now().strftime('%Y-%m-%d %H:%M')}_", ""]
    for row in rows:
        d = {headers[i]: fmt_cell(v) for i, v in enumerate(row) if i < len(headers)}
        if not any(d.values()): continue
        head_field = headers[1] if len(headers) > 1 else headers[0]
        lines.append(f"## {d.get(head_field, '(無題)')}")
        for k in headers:
            if not k or k == head_field: continue
            v = d.get(k, '')
            if not v: continue
            lines.append(f"- **{k}**: {v}")
        lines.append("")
    return '\n'.join(lines) + '\n'

def render_index_html(stats):
    nav = nav_html('00-readme')
    blocks = []
    for gname, sheets in SHEET_GROUPS:
        rows = []
        for sn in sheets:
            slug = SHEET_SLUGS.get(sn)
            if not slug:
                continue
            n = stats.get(sn, 0)
            rows.append(f'<tr><td><a href="{slug}.html">{html.escape(_short(sn))}</a></td>'
                        f'<td style="text-align:right">{n}</td></tr>')
        if rows:
            blocks.append(f'<div class="genre-h">{html.escape(gname)}</div>'
                          '<div class="table-wrap"><table><thead><tr><th>シート</th>'
                          '<th style="text-align:right">行数</th></tr></thead><tbody>'
                          + ''.join(rows) + '</tbody></table></div>')
    updated = datetime.now().strftime('%Y-%m-%d %H:%M')
    total = sum(stats.values())
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>エンタメ業界・時事情報収集</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{CSS}</head><body><div class="container">
<header><h1>エンタメ業界・時事情報収集</h1><span class="subtle">最終更新: {updated} · 全{total}行</span></header>
{nav}
<div class="section">
  <h2>シート一覧</h2>
  {''.join(blocks)}
</div>
<div class="section">
  <h2>更新スケジュール</h2>
  <ul>
    <li><b>毎朝 7:00</b> — 6ジャンル横断の日次ニュース + 横断的問い</li>
    <li><b>毎週月 9:00</b> — 全ジャンルの企業・レーベル業績 + 作品ウォッチ</li>
    <li><b>毎週水 12:00</b> — 全ジャンルのトレンド + 掘り下げ・分析</li>
    <li><b>毎週金 18:00</b> — 全ジャンルのリリース・公開・チャート</li>
    <li><b>毎月 1日 9:00</b> — 業界マトリクス + 今後起きそうなこと</li>
    <li><b>2/5/8/11月 上旬</b> — 四半期決算サマリ</li>
  </ul>
</div>
<footer>自動生成 · GitHub Pages</footer>
</div></body></html>"""

def render_index_md(stats):
    lines = ["# エンタメ業界・時事情報収集", "",
             f"_最終更新: {datetime.now().strftime('%Y-%m-%d %H:%M')}_", "",
             "## シート一覧", ""]
    for sn, slug in SHEET_SLUGS.items():
        if sn == 'README': continue
        n = stats.get(sn, 0)
        lines.append(f"- [{sn}](md/{slug}.md) ({n}行)")
    lines += ["", "## 更新スケジュール", "",
              "- **毎朝 7:00** — 日次ニュース10件追加",
              "- **毎週月 9:00** — 国内/海外大手の株価・IR更新",
              "- **毎週金 18:00** — 新作リリース週次レポート",
              "- **毎週水 12:00** — 業界トレンド・話題のゲームデザイン",
              "- **2/5/8/11月 上旬** — 四半期決算サマリ", "",
              "ダッシュボード（GitHub Pages）から閲覧できます。"]
    return '\n'.join(lines) + '\n'

def main():
    if not XLSX_PATH.exists():
        print(f"ERROR: not found: {XLSX_PATH}", file=sys.stderr); sys.exit(1)
    MD_DIR.mkdir(parents=True, exist_ok=True)
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(XLSX_PATH, data_only=True)
    # --- 最近更新カテゴリの判定（ヘッダーメニューに●印） ---
    from datetime import date as _date
    UPDATE_COLS = ('日付', '記録日', '更新日', '最終更新日')
    RECENT_WINDOW_DAYS = 3
    _dpat = re.compile(r'(\d{4}-\d{2}-\d{2})')
    _latest = {}
    for _sn in wb.sheetnames:
        if _sn not in SHEET_SLUGS or _sn == "README":
            continue
        _ws = wb[_sn]
        _hdr = [c.value for c in _ws[1]]
        _cols = [i for i, h in enumerate(_hdr) if h in UPDATE_COLS]
        if not _cols:
            continue
        _mx = ''
        for _row in _ws.iter_rows(min_row=2, values_only=True):
            for _i in _cols:
                if _i < len(_row):
                    _v = _row[_i]
                    if isinstance(_v, datetime):
                        _sd = _v.strftime('%Y-%m-%d')
                    else:
                        _m = _dpat.search(str(_v) if _v is not None else '')
                        _sd = _m.group(1) if _m else ''
                    if _sd > _mx:
                        _mx = _sd
        if _mx:
            _latest[SHEET_SLUGS[_sn]] = _mx
    if _latest:
        _newest = _date.fromisoformat(max(_latest.values()))
        for _slug, _d in _latest.items():
            if (_newest - _date.fromisoformat(_d)).days <= RECENT_WINDOW_DAYS:
                _RECENT_SLUGS.add(_slug)
                _RECENT_DATES[_slug] = _d
    stats = {}
    sheet_data = {}
    for sn in wb.sheetnames:
        if sn not in SHEET_SLUGS: continue
        if sn == "README": continue
        ws = wb[sn]
        headers = [c.value for c in ws[1]]
        # 末尾の全Noneヘッダはトリム
        while headers and headers[-1] is None:
            headers.pop()
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            row = list(row[:len(headers)])
            if any(v not in (None, '') for v in row):
                rows.append(row)
        # 全シートを日付列を基準に新しい順（降順）で並べ替え（最新を上に）。
        # 各シートで最初に見つかった日付系ヘッダを基準にする。
        # YYYY-MM-DD形式を抽出してキーに使用。日付として解釈できない行は最下段。
        # 週間ランキングは専用レンダラ側でソートするため除外。
        DATE_HEADER_PRIORITY = ['日付', '記録日', '発売日', '更新日', '決算発表日',
                                '開催日', '集計週', '最終更新日']
        if sn != RANKING_SHEET:
            date_col = None
            for _cand in DATE_HEADER_PRIORITY:
                if _cand in headers:
                    date_col = headers.index(_cand)
                    break
            if date_col is not None:
                _date_pat = re.compile(r'(\d{4}-\d{2}-\d{2})')
                def _date_key(r, _c=date_col):
                    v = r[_c] if _c < len(r) else None
                    if isinstance(v, datetime):
                        return v.strftime('%Y-%m-%d')
                    s = str(v) if v is not None else ''
                    m = _date_pat.search(s)
                    return m.group(1) if m else ''
                rows.sort(key=_date_key, reverse=True)
        stats[sn] = len(rows)
        sheet_data[sn] = (headers, rows)
        slug = SHEET_SLUGS[sn]
        if sn == RANKING_SHEET:
            (MD_DIR / f"{slug}.md").write_text(render_ranking_md(sn, headers, rows), encoding='utf-8')
            (DOCS_DIR / f"{slug}.html").write_text(render_ranking_html(sn, headers, rows), encoding='utf-8')
        elif sn == "日次ニュース":
            (MD_DIR / f"{slug}.md").write_text(render_section_md(sn, headers, rows), encoding='utf-8')
            (DOCS_DIR / f"{slug}.html").write_text(render_daily_html(sn, headers, rows), encoding='utf-8')
        elif sn in SECTION_SHEETS:
            (MD_DIR / f"{slug}.md").write_text(render_section_md(sn, headers, rows), encoding='utf-8')
            (DOCS_DIR / f"{slug}.html").write_text(render_section_html(sn, headers, rows), encoding='utf-8')
        else:
            (MD_DIR / f"{slug}.md").write_text(render_table_md(sn, headers, rows), encoding='utf-8')
            (DOCS_DIR / f"{slug}.html").write_text(render_table_html(sn, headers, rows), encoding='utf-8')
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import dashboard
        idx = dashboard.render_dashboard_html(sheet_data, stats, SHEET_GROUPS, SHEET_SLUGS, _short,
                                              _RECENT_SLUGS, nav_html, CSS)
    except Exception as e:
        print("dashboard fallback:", e, file=sys.stderr)
        idx = render_index_html(stats)
    (DOCS_DIR / "index.html").write_text(idx, encoding='utf-8')
    (REPO_ROOT / "README.md").write_text(render_index_md(stats), encoding='utf-8')
    print(f"Generated for {len(stats)} sheets. Total rows: {sum(stats.values())}")
    try:
        import gen_maps
        gen_maps.generate()
    except Exception as e:
        print("map generation skipped:", e, file=sys.stderr)

if __name__ == "__main__":
    main()
