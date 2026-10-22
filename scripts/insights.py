"""「問い・掘り下げ」統合ページと「予測マップ」の生成。

- 論点テーマ（xlsxの「論点テーマ」シート）のキーワードで、横断的問い・掘り下げ・予測を自動分類する。
- 各テーマに「いまの見立て」と、問い/掘り下げの推移（新しい順）を表示する。
- 予測は「想定時期」を期間に変換し、タイムライン（ガント風）で見せる。
"""
import html
import re
from collections import Counter, defaultdict
from datetime import date, datetime

GENRES = ["ゲーム", "映画", "アニメ", "漫画", "小説", "音楽", "横断"]
_DP = re.compile(r"(\d{4}-\d{2}-\d{2})")


def esc(s):
    return html.escape("" if s is None else str(s))


def dstr(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    m = _DP.search(str(v))
    return m.group(1) if m else ""


def norm_genre(g):
    g = str(g or "").strip()
    for x in GENRES:
        if g == x:
            return x
    if g.startswith("横断"):
        return "横断"
    for x in GENRES:
        if x in g:
            return x
    return "横断"


def gbadge(g):
    g = norm_genre(g)
    return f'<span class="gb g{GENRES.index(g)}">{esc(g)}</span>'


def short(s, n):
    s = (s or "").strip().replace("\n", " ")
    return s if len(s) <= n else s[:n] + "…"


def dicts(sheet_data, name):
    if name not in sheet_data:
        return []
    h, rows = sheet_data[name]
    return [{k: r[i] for i, k in enumerate(h) if k and i < len(r)} for r in rows]


# ------------------------------------------------------------------ テーマ分類
def load_themes(sheet_data):
    out = []
    for t in dicts(sheet_data, "論点テーマ"):
        if not t.get("テーマID"):
            continue
        kw = [k.strip() for k in str(t.get("キーワード（分類用・カンマ区切り）") or "").split(",") if k.strip()]
        out.append(dict(id=t["テーマID"], name=t.get("テーマ名") or "", q=t.get("問いの形") or "",
                        kw=kw, view=t.get("いまの見立て") or "", vdate=dstr(t.get("見立て更新日"))))
    out.append(dict(id="T99", name="その他の論点", q="分類外の問い", kw=[], view="", vdate=""))
    return out


def classify(themes, title, body=""):
    best, score = "T99", 0
    tl, bl = title.lower(), body.lower()
    for t in themes:
        s = 0
        for k in t["kw"]:
            kl = k.lower()
            s += 3 * tl.count(kl) + bl.count(kl)
        if s > score:
            best, score = t["id"], s
    return best


def collect(sheet_data):
    themes = load_themes(sheet_data)
    items = defaultdict(list)
    for q in dicts(sheet_data, "横断的問い"):
        title = str(q.get("問い") or "")
        if not title:
            continue
        tid = classify(themes, title, str(q.get("気付き") or ""))
        items[tid].append(dict(kind="問い", d=dstr(q.get("記録日")), title=title, body=q.get("気付き") or "",
                               src=q.get("発火元（ニュース等）") or "", rel=q.get("関連企業/IP") or "",
                               ext=q.get("他業界への接続候補") or "", status=q.get("ステータス") or "",
                               g=norm_genre(q.get("ジャンル"))))
    for d in dicts(sheet_data, "掘り下げ・分析"):
        title = str(d.get("テーマ") or "")
        if not title:
            continue
        tid = classify(themes, title, str(d.get("詳細") or ""))
        items[tid].append(dict(kind="掘り下げ", d=dstr(d.get("記録日")), title=title, body=d.get("詳細") or "",
                               imp=d.get("示唆") or "", kw=d.get("キーワード・構造解説") or "",
                               rel=d.get("関連企業/IP") or "", link=d.get("参考リンク") or "",
                               g=norm_genre(d.get("ジャンル"))))
    preds = []
    for f in dicts(sheet_data, "今後起きそうなこと"):
        text = str(f.get("今後起きそうなこと") or "")
        if not text:
            continue
        p = dict(d=dstr(f.get("記録日")), text=text, conf=str(f.get("確度") or "中"), when=str(f.get("想定時期") or ""),
                 rel=f.get("関連企業/IP") or "", src=f.get("根拠となるニュース") or "", check=f.get("検証メモ") or "",
                 g=norm_genre(f.get("ジャンル")), verdict=str(f.get("判定") or "").strip())
        p["tid"] = classify(themes, text, p["src"])
        p["s"], p["e"], p["fuzzy"] = window(p["d"], p["when"])
        preds.append(p)
    for v in items.values():
        v.sort(key=lambda x: x["d"], reverse=True)
    return themes, items, preds


# ------------------------------------------------------------------ 想定時期 → 期間
def _ym_add(y, m, k):
    t = y * 12 + (m - 1) + k
    return t // 12, t % 12 + 1


def window(rec, when):
    """記録日と「想定時期」の文字列から (開始, 終了) の (年, 月) と、あいまいフラグを推定する。"""
    try:
        r = date.fromisoformat(rec)
    except Exception:
        r = date.today()
    R = (r.year, r.month)
    w = str(when or "")
    for a_, b_ in [("ヵ月", "ヶ月"), ("カ月", "ヶ月"), ("か月", "ヶ月"), ("ケ月", "ヶ月"), ("〜", "~"), ("～", "~"), ("（", "("), ("）", ")")]:
        w = w.replace(a_, b_)
    w = re.sub(r"(\d)\s*-\s*(\d)", lambda m: m.group(1) + "~" + m.group(2) if len(m.group(0)) < 4 or True else m.group(0), w)
    w = re.sub(r"(20\d{2})-(\d{2})(?!\d)", r"\1年\2月", w) if re.search(r"20\d{2}-\d{2}(?!\d)", w) and not re.search(r"20\d{2}-20\d{2}", w) else w
    w = w.replace("半年", "6ヶ月").replace("数ヶ月", "3ヶ月")
    w = re.sub(r"(?<!\d)(\d{1,2})\s*~\s*(\d{1,2})\s*年(以内|内)?",
               lambda m: f"{int(m.group(1)) * 12}ヶ月~{int(m.group(2)) * 12}ヶ月" + ("以内" if m.group(3) else ""), w)
    w = re.sub(r"(?<!\d)(\d{1,2})年(以内|内)", lambda m: f"{int(m.group(1)) * 12}ヶ月以内", w)
    w = re.sub(r"(?<!\d)(\d)年", lambda m: f"{int(m.group(1)) * 12}ヶ月", w)

    def period(tok):
        ym = re.search(r"(20\d{2})", tok)
        if not ym:
            return None
        y = int(ym.group(1))
        mo = re.search(r"20\d{2}年?(\d{1,2})月", tok)
        if mo:
            m = int(mo.group(1)); return (y, m), (y, m)
        if "年内" in tok or "内" in tok.replace("以内", ""):
            return (R if y == R[0] else (y, 1)), (y, 12)
        if re.search(r"Q1|1Q|第1四半期", tok): return (y, 1), (y, 3)
        if re.search(r"Q2|2Q|第2四半期", tok): return (y, 4), (y, 6)
        if re.search(r"Q3|3Q|第3四半期", tok): return (y, 7), (y, 9)
        if re.search(r"Q4|4Q|第4四半期", tok): return (y, 10), (y, 12)
        if re.search(r"前半|上期|上半期|H1|春", tok): return (y, 1), (y, 6)
        if re.search(r"秋", tok): return (y, 9), (y, 11)
        if re.search(r"後半|下期|下半期|H2|冬|年末", tok): return (y, 7), (y, 12)
        return (y, 1), (y, 12)

    def parse(text):
        toks = [t.strip() for t in text.split("~")]
        units = []
        for t in toks:
            units.append("m" if "ヶ月" in t else None)
        rng = []
        for i, t in enumerate(toks):
            p = period(t)
            if p:
                rng.append(p); continue
            if "年内" in t:
                rng.append((R, (r.year, 12))); continue
            n = re.search(r"(\d+)\s*ヶ月", t)
            if not n and re.fullmatch(r"\d+", t) and any(units[i + 1:]):
                n = re.search(r"(\d+)", t)
            if n:
                k = int(n.group(1)); p = _ym_add(R[0], R[1], k); rng.append((p, p)); continue
            if re.search(r"来年|翌年", t):
                rng.append(((r.year + 1, 1), (r.year + 1, 12))); continue
            rng.append(None)
        got = [x for x in rng if x]
        if not got:
            return None
        s = rng[0][0] if rng[0] else R
        e = got[-1][1]
        if "以内" in text or "まで" in text:
            s = R
        if len(toks) == 1 and got and "ヶ月" in text and not ("以内" in text or "まで" in text) and not re.search(r"20\d{2}", text):
            s = got[0][0]
        return s, e

    main = re.sub(r"\(.*?\)", "", w)
    res = parse(main) or (parse(" ".join(re.findall(r"\((.*?)\)", w))) if "(" in w else None)
    if not res:
        return R, _ym_add(R[0], R[1], 12), True
    s, e = res
    if s < R:
        s = R
    if e < s:
        s = R
        e = max(e, s)
    return s, e, False


# ------------------------------------------------------------------ 共通CSS
CSS2 = """
<style>
.gb{display:inline-block;font-size:11px;font-weight:700;padding:1px 8px;border-radius:999px;white-space:nowrap}
.g0{background:rgba(124,92,255,.18);color:#a08cff}.g1{background:rgba(229,83,75,.18);color:#ff7b72}
.g2{background:rgba(240,136,62,.18);color:#f0883e}.g3{background:rgba(46,160,67,.18);color:#3fb950}
.g4{background:rgba(187,128,9,.2);color:#d29922}.g5{background:rgba(219,97,162,.18);color:#f778ba}
.g6{background:rgba(88,166,255,.18);color:#79c0ff}
@media (prefers-color-scheme: light){.g0{color:#6639ba}.g1{color:#c93c37}.g2{color:#bc4c00}.g3{color:#1a7f37}.g4{color:#9a6700}.g5{color:#bf3989}.g6{color:#0969da}}
.pghead{margin:4px 0 16px}
.pghead .eyebrow{font-size:11px;letter-spacing:.12em;color:var(--muted);font-weight:700}
.pghead h1{font-size:24px;margin:4px 0 6px}
.pghead p{margin:0;color:var(--muted);font-size:13px}
.fbar{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:0 0 14px}
.fbar .lbl{color:var(--muted);font-size:12px;min-width:52px}
.fb{border:1px solid var(--border);background:var(--panel);color:var(--text);border-radius:999px;padding:4px 12px;font-size:12px;cursor:pointer}
.fb.active{background:var(--accent);border-color:var(--accent);color:#fff}
.fb .n{color:var(--muted);font-size:11px;margin-left:4px}.fb.active .n{color:#fff}
.badge{display:inline-block;font-size:10.5px;font-weight:700;padding:1px 7px;border-radius:5px;border:1px solid var(--border);margin-right:4px;color:var(--muted)}
.badge.new{background:#e8583a;border-color:#e8583a;color:#fff}
.badge.k-q{border-color:#58a6ff;color:#58a6ff}.badge.k-d{border-color:#d29922;color:#d29922}.badge.k-p{border-color:#a371f7;color:#a371f7}
</style>
"""

DEEP_CSS = """
<style>
.tgrid{display:grid;gap:10px;grid-template-columns:repeat(auto-fill,minmax(270px,1fr));margin-bottom:22px}
.tcard{display:block;border:1px solid var(--border);background:var(--panel);border-radius:12px;padding:12px 14px;text-decoration:none;color:var(--text)}
.tcard:hover{border-color:var(--accent)}
.tcard .tn{font-weight:800;font-size:14px;line-height:1.45}
.tcard .tq{color:var(--muted);font-size:12px;margin:2px 0 6px}
.tcard .tv{font-size:12px;line-height:1.6;color:var(--text);display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.tcard .tm{font-size:11px;color:var(--muted);margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.theme{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:18px 20px;margin-bottom:16px;scroll-margin-top:12px}
.theme h2{margin:0;font-size:18px}
.theme .q{color:var(--muted);font-size:13px;margin:2px 0 8px}
.theme .meta{display:flex;gap:6px;flex-wrap:wrap;align-items:center;font-size:12px;color:var(--muted);margin-bottom:10px}
.view{border-left:4px solid var(--accent);background:rgba(88,166,255,.08);border-radius:0 8px 8px 0;padding:10px 14px;font-size:13.5px;line-height:1.75;margin-bottom:14px}
.view b{display:block;font-size:11.5px;color:var(--accent);margin-bottom:2px}
.tl{position:relative;margin-left:8px;padding-left:18px;border-left:2px solid var(--border)}
.tli{position:relative;padding:4px 0 12px}
.tli:before{content:"";position:absolute;left:-25px;top:10px;width:10px;height:10px;border-radius:50%;background:var(--panel);border:2px solid var(--accent)}
.tli.kd:before{border-color:#d29922}.tli.kp:before{border-color:#a371f7}
.tli .h{font-size:12px;color:var(--muted);display:flex;gap:6px;flex-wrap:wrap;align-items:center}
.tli .t{font-weight:700;font-size:13.5px;line-height:1.55;margin:2px 0}
.tli details summary{cursor:pointer;color:var(--accent);font-size:12px}
.tli .b{font-size:12.5px;line-height:1.7;margin-top:6px}
.tli .b dt{color:var(--muted);font-size:11.5px;margin-top:6px}.tli .b dd{margin:0}
.morebtn{border:1px solid var(--border);background:transparent;color:var(--accent);border-radius:999px;padding:4px 12px;font-size:12px;cursor:pointer}
</style>
"""


# ------------------------------------------------------------------ 問い・掘り下げページ
def render_deep_html(sheet_data, CSS, nav_html, latest):
    themes, items, preds = collect(sheet_data)
    pcount = Counter(p["tid"] for p in preds)
    try:
        ld = date.fromisoformat(latest)
    except Exception:
        ld = date.today()

    def is_new(d):
        try:
            return (ld - date.fromisoformat(d)).days <= 3
        except Exception:
            return False

    order = []
    for t in themes:
        its = items.get(t["id"], [])
        if not its:
            continue
        last = max(x["d"] for x in its)
        order.append((last, t, its))
    order.sort(key=lambda x: (x[1]["id"] == "T99", x[0]), reverse=False)
    order.sort(key=lambda x: (x[1]["id"] != "T99", x[0]), reverse=True)

    gtotal = Counter()
    cards, blocks = [], []
    for last, t, its in order:
        gs = Counter(x["g"] for x in its)
        gtotal.update(gs)
        nq = sum(1 for x in its if x["kind"] == "問い")
        nd = len(its) - nq
        first = min(x["d"] for x in its if x["d"]) if any(x["d"] for x in its) else ""
        newn = sum(1 for x in its if is_new(x["d"]))
        view = t["view"] or (its[0].get("imp") or its[0].get("body") or "")
        newb = f'<span class="badge new">NEW {newn}</span>' if newn else ""
        gattr = "|".join(g for g in gs)
        cards.append(f'<a class="tcard" href="#{t["id"]}" data-g="{gattr}"><div class="tn">{esc(t["name"])}</div>'
                     f'<div class="tq">{esc(t["q"])}</div><div class="tv">{esc(view)}</div>'
                     f'<div class="tm">{newb}<span>問い {nq}</span><span>掘り下げ {nd}</span><span>予測 {pcount.get(t["id"], 0)}</span>'
                     f'<span>最終 {last[5:]}</span></div></a>')
        tl = []
        for i, x in enumerate(its):
            hide = ' data-more="1" style="display:none"' if i >= 8 else ""
            if x["kind"] == "問い":
                body = "".join(f'<dt>{k}</dt><dd>{esc(v)}</dd>' for k, v in
                               [("気付き", x["body"]), ("発火元", x["src"]), ("他業界への接続", x["ext"]), ("関連", x["rel"])] if v)
                kb = '<span class="badge k-q">問い</span>'
                cls = ""
                st = f'<span class="badge">{esc(x["status"])}</span>' if x.get("status") else ""
            else:
                body = "".join(f'<dt>{k}</dt><dd>{esc(v)}</dd>' for k, v in
                               [("詳細", x["body"]), ("構造解説", x["kw"]), ("示唆", x["imp"]), ("関連", x["rel"])] if v)
                if x.get("link"):
                    body += f'<dt>参考</dt><dd><a href="{esc(x["link"])}" target="_blank" rel="noopener">{esc(short(x["link"], 60))}</a></dd>'
                kb = '<span class="badge k-d">掘り下げ</span>'
                cls = " kd"
                st = ""
            nb = '<span class="badge new">NEW</span>' if is_new(x["d"]) else ""
            tl.append(f'<div class="tli{cls}" data-g="{x["g"]}"{hide}><div class="h">{esc(x["d"])} {kb}{gbadge(x["g"])}{st}{nb}</div>'
                      f'<div class="t">{esc(x["title"])}</div>'
                      f'<details><summary>内容を開く</summary><dl class="b">{body}</dl></details></div>')
        more = (f'<button class="morebtn" onclick="this.parentNode.querySelectorAll(\'[data-more]\').forEach(e=>e.style.display=\'\');this.remove()">'
                f'さらに{len(its) - 8}件を表示</button>') if len(its) > 8 else ""
        vlabel = f'いまの見立て（{t["vdate"]}更新）' if t["view"] and t["vdate"] else "いまの見立て（最新エントリより）"
        blocks.append(f'<div class="theme" id="{t["id"]}" data-g="{gattr}"><h2>{esc(t["name"])} {newb}</h2><div class="q">{esc(t["q"])}</div>'
                      f'<div class="meta">{"".join(gbadge(g) for g, _ in gs.most_common())}<span>· 問い {nq} · 掘り下げ {nd} · '
                      f'予測 {pcount.get(t["id"], 0)} · {first}〜 · 最終更新 {last}</span></div>'
                      f'<div class="view"><b>{vlabel}</b>{esc(view)}</div>'
                      f'<div class="tl">{"".join(tl)}</div>{more}</div>')

    fbtn = ['<span class="lbl">ジャンル</span><button class="fb active" data-k="__all__">すべて</button>']
    fbtn += [f'<button class="fb" data-k="{g}">{g}<span class="n">{gtotal[g]}</span></button>' for g in GENRES if gtotal.get(g)]
    js = """<script>(function(){const sel=new Set();const bs=document.querySelectorAll('#gf .fb');
function ap(){document.querySelectorAll('.tcard,.theme').forEach(el=>{const gs=(el.dataset.g||'').split('|');
el.style.display=(!sel.size||gs.some(g=>sel.has(g)))?'':'none';});
document.querySelectorAll('.tli').forEach(el=>{if(el.dataset.more&&el.style.display==='none'&&!sel.size)return;
el.style.display=(!sel.size||sel.has(el.dataset.g))?(el.dataset.more&&!sel.size?'none':''):'none';});}
bs.forEach(b=>b.onclick=()=>{const k=b.dataset.k;if(k==='__all__'){sel.clear();bs.forEach(x=>x.classList.remove('active'));b.classList.add('active');}
else{sel.has(k)?(sel.delete(k),b.classList.remove('active')):(sel.add(k),b.classList.add('active'));
document.querySelector('#gf [data-k="__all__"]').classList.toggle('active',!sel.size);}ap();});})();</script>"""
    nq_all = sum(1 for v in items.values() for x in v if x["kind"] == "問い")
    nd_all = sum(1 for v in items.values() for x in v if x["kind"] == "掘り下げ")
    updated = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>問い・掘り下げ | エンタメ業界トラッカー</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{CSS}{CSS2}{DEEP_CSS}</head><body><div class="container">
{nav_html('18-deep')}
<div class="pghead"><div class="eyebrow">DEEP DIVES — 問いで追う論点</div><h1>問い・掘り下げ</h1>
<p>横断的問い {nq_all}件と掘り下げ {nd_all}件を、{len(order)}本の論点テーマに整理。各テーマに「いまの見立て」と論点の推移（新しい順）を載せています。最終更新 {updated}</p></div>
<div class="fbar" id="gf">{"".join(fbtn)}</div>
<div class="tgrid">{"".join(cards)}</div>
{"".join(blocks)}
<footer>テーマ分類は「論点テーマ」シートのキーワードで自動判定 · 見立ては同シートで更新</footer>
</div>{js}</body></html>""", themes, items, preds


# ------------------------------------------------------------------ 予測マップ
PRED_CSS = """
<style>
.pk{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:8px;margin-bottom:16px}
.pk div{border:1px solid var(--border);background:var(--panel);border-radius:10px;padding:10px 12px}
.pk b{display:block;font-size:24px;line-height:1.1}
.pk span{font-size:11.5px;color:var(--muted)}
.box{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;margin-bottom:16px}
.box h2{font-size:16px;margin:0 0 4px}.box .lead{color:var(--muted);font-size:12.5px;margin:0 0 12px}
.gantt{overflow-x:auto}
.gin{min-width:900px}
.gh,.gr{display:grid;grid-template-columns:330px 1fr;align-items:center}
.gh .axis{position:relative;height:24px;border-bottom:1px solid var(--border)}
.gh .axis span{position:absolute;top:2px;font-size:10.5px;color:var(--muted);transform:translateX(-50%);white-space:nowrap}
.ggroup{font-size:12px;font-weight:800;margin:14px 0 4px;padding-top:6px;border-top:1px dashed var(--border)}
.ggroup a{color:var(--text);text-decoration:none}
.gr{min-height:28px;border-bottom:1px solid rgba(127,127,127,.12);cursor:pointer}
.gr:hover{background:rgba(88,166,255,.06)}
.gr .lab{font-size:12px;line-height:1.4;padding:4px 10px 4px 0;display:flex;gap:6px;align-items:flex-start}
.gr .lab .tx{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.gr .track{position:relative;height:28px}
.gr .bar{position:absolute;top:8px;height:12px;border-radius:6px;opacity:.9}
.gr .bar.fz{background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.35) 0 4px,transparent 4px 8px)}
.gr .rec{position:absolute;top:10px;width:8px;height:8px;border-radius:50%;background:var(--muted);transform:translateX(-50%)}
.today{position:absolute;top:0;bottom:0;width:2px;background:#e8583a;z-index:1}
.c-hi{background:#2ea043}.c-mid{background:#58a6ff}.c-lo{background:#8b949e}
.st-ok{background:#2ea043}.st-ng{background:#f85149}
.gdet{display:none;grid-column:1/-1;font-size:12.5px;line-height:1.7;padding:8px 12px 12px;background:rgba(127,127,127,.06);border-radius:8px;margin:4px 0 8px}
.gdet.on{display:block}
.gdet dt{color:var(--muted);font-size:11.5px}.gdet dd{margin:0 0 4px}
.lgd{display:flex;gap:14px;flex-wrap:wrap;font-size:11.5px;color:var(--muted);margin-top:8px}
.lgd i{display:inline-block;width:18px;height:10px;border-radius:5px;margin-right:5px;vertical-align:middle}
.mx{display:grid;grid-template-columns:90px repeat(3,1fr);gap:6px;font-size:12px}
.mx .hd{color:var(--muted);font-weight:700;text-align:center;padding:4px}
.mx .rl{color:var(--muted);font-weight:700;display:flex;align-items:center}
.mx .cell{border-radius:8px;padding:8px;text-align:center;font-weight:800;font-size:18px;border:1px solid var(--border)}
.mx .cell small{display:block;font-size:10.5px;font-weight:400;color:var(--muted)}
.due{display:grid;gap:8px;grid-template-columns:repeat(auto-fill,minmax(300px,1fr))}
.due .c{border:1px solid var(--border);border-left:4px solid #e8583a;border-radius:8px;padding:10px 12px;font-size:12.5px;line-height:1.6}
.due .c .w{font-size:11px;color:var(--muted)}
@media (max-width:640px){.gin{min-width:720px}.gh,.gr{grid-template-columns:220px 1fr}}
</style>
"""


def _mi(ym):
    return ym[0] * 12 + ym[1] - 1


def horizon(p, ld):
    """短期(〜3ヶ月)/中期(〜1年)/長期(1年〜) を、記録日から終了時期までの長さで判定"""
    try:
        r = date.fromisoformat(p["d"])
    except Exception:
        r = ld
    span = _mi(p["e"]) - (r.year * 12 + r.month - 1)
    return "短期" if span <= 3 else ("中期" if span <= 12 else "長期")


def status(p, ld):
    v = p.get("verdict") or ""
    if v.startswith("的中") or v in ("当たり", "○"):
        return "ok"
    if v.startswith("外れ") or v in ("×",):
        return "ng"
    cur = ld.year * 12 + ld.month - 1
    if _mi(p["e"]) < cur:
        return "due"
    if _mi(p["s"]) <= cur:
        return "live"
    return "future"


def render_pred_html(sheet_data, CSS, nav_html, latest, themes=None, preds=None):
    if preds is None:
        themes, _, preds = collect(sheet_data)
    tname = {t["id"]: t["name"] for t in themes}
    try:
        ld = date.fromisoformat(latest)
    except Exception:
        ld = date.today()
    for p in preds:
        p["st"] = status(p, ld)
        p["hz"] = horizon(p, ld)
    stc = Counter(p["st"] for p in preds)
    # 軸の範囲
    lo = min(min(_mi(p["s"]) for p in preds), ld.year * 12 + ld.month - 1 - 2) if preds else 0
    hi = max(_mi(p["e"]) for p in preds) + 1 if preds else 12
    hi = min(hi, lo + 48)
    span = max(hi - lo, 1)
    pct = lambda m: max(0.0, min(100.0, (m - lo) / span * 100))
    ticks = []
    for m in range(lo, hi + 1):
        y, mo = divmod(m, 12)
        if mo % 3 == 0:
            ticks.append(f'<span style="left:{pct(m):.2f}%">{y}/{mo + 1:02d}</span>')
    today_x = pct(ld.year * 12 + ld.month - 1 + (ld.day - 1) / 31)
    # グループ：テーマ別
    groups = defaultdict(list)
    for p in preds:
        groups[p["tid"]].append(p)
    rows = []
    for tid in sorted(groups, key=lambda k: (-len(groups[k]), k)):
        ps = sorted(groups[tid], key=lambda p: (_mi(p["s"]), _mi(p["e"])))
        rows.append(f'<div class="ggroup" data-tid="{tid}"><a href="18-deep.html#{tid}">{esc(tname.get(tid, "その他"))}</a> '
                    f'<span style="color:var(--muted);font-weight:400">{len(ps)}件</span></div>')
        for i, p in enumerate(ps):
            col = "st-ok" if p["st"] == "ok" else ("st-ng" if p["st"] == "ng" else ("c-hi" if p["conf"] == "高" else ("c-lo" if p["conf"] == "低" else "c-mid")))
            fz = " fz" if p["fuzzy"] else ""
            x0, x1 = pct(_mi(p["s"])), pct(_mi(p["e"]) + 1)
            try:
                r = date.fromisoformat(p["d"]); rx = pct(r.year * 12 + r.month - 1 + (r.day - 1) / 31)
            except Exception:
                rx = None
            stl = {"ok": "的中", "ng": "外れ", "due": "検証待ち", "live": "期間中", "future": "これから"}[p["st"]]
            key = f"{tid}-{i}"
            det = "".join(f'<dt>{k}</dt><dd>{esc(v)}</dd>' for k, v in
                          [("予測", p["text"]), ("想定時期", p["when"]), ("根拠", p["src"]), ("検証メモ", p["check"]),
                           ("関連", p["rel"]), ("判定", p["verdict"] or "未判定")] if v)
            rows.append(f'<div class="gr" data-g="{p["g"]}" data-st="{p["st"]}" data-c="{esc(p["conf"])}" onclick="document.getElementById(\'d{key}\').classList.toggle(\'on\')">'
                        f'<div class="lab">{gbadge(p["g"])}<span class="tx">{esc(short(p["text"], 80))}</span></div>'
                        f'<div class="track"><div class="today" style="left:{today_x:.2f}%"></div>'
                        + (f'<div class="rec" style="left:{rx:.2f}%" title="記録日 {p["d"]}"></div>' if rx is not None else "")
                        + f'<div class="bar {col}{fz}" style="left:{x0:.2f}%;width:{max(x1 - x0, 1.2):.2f}%" title="{esc(p["when"])} · 確度{esc(p["conf"])} · {stl}"></div></div>'
                        f'<div class="gdet" id="d{key}"><span class="badge">確度 {esc(p["conf"])}</span><span class="badge">{p["hz"]}</span>'
                        f'<span class="badge">{stl}</span><span class="badge">記録 {p["d"]}</span><dl>{det}</dl></div></div>')
    # 確度×時間軸
    mx = Counter((p["conf"], p["hz"]) for p in preds)
    mxmax = max(mx.values()) if mx else 1
    cells = ['<div></div>'] + [f'<div class="hd">{h}<br><small style="font-weight:400">{d}</small></div>'
                               for h, d in [("短期", "〜3ヶ月"), ("中期", "〜1年"), ("長期", "1年〜")]]
    for c in ["高", "中", "低"]:
        cells.append(f'<div class="rl">確度 {c}</div>')
        for h in ["短期", "中期", "長期"]:
            n = mx.get((c, h), 0)
            a = 0.08 + 0.5 * n / mxmax if n else 0
            cells.append(f'<div class="cell" style="background:rgba(88,166,255,{a:.2f})">{n or "—"}<small>{c}×{h}</small></div>')
    # 検証待ち
    due = sorted([p for p in preds if p["st"] == "due"], key=lambda p: _mi(p["e"]))
    duec = "".join(f'<div class="c"><div class="w">期限 {p["e"][0]}/{p["e"][1]:02d} · 記録 {p["d"]} · 確度 {esc(p["conf"])} {gbadge(p["g"])}</div>'
                   f'{esc(short(p["text"], 120))}<div class="w">検証: {esc(short(p["check"], 90))}</div></div>' for p in due[:12])
    soon = sorted([p for p in preds if p["st"] == "live" and _mi(p["e"]) <= ld.year * 12 + ld.month - 1 + 2], key=lambda p: _mi(p["e"]))
    soonc = "".join(f'<div class="c" style="border-left-color:#d29922"><div class="w">期限 {p["e"][0]}/{p["e"][1]:02d} · 確度 {esc(p["conf"])} {gbadge(p["g"])}</div>'
                    f'{esc(short(p["text"], 120))}<div class="w">検証: {esc(short(p["check"], 90))}</div></div>' for p in soon[:12])
    gcnt = Counter(p["g"] for p in preds)
    fb = ['<span class="lbl">ジャンル</span><button class="fb active" data-k="__all__">すべて</button>']
    fb += [f'<button class="fb" data-k="{g}">{g}<span class="n">{gcnt[g]}</span></button>' for g in GENRES if gcnt.get(g)]
    sb = ['<span class="lbl">状態</span><button class="fb active" data-k="__all__">すべて</button>']
    sb += [f'<button class="fb" data-k="{k}">{l}<span class="n">{stc.get(k, 0)}</span></button>'
           for k, l in [("due", "検証待ち"), ("live", "期間中"), ("future", "これから"), ("ok", "的中"), ("ng", "外れ")] if stc.get(k)]
    js = """<script>(function(){const S={g:new Set(),s:new Set()};
function ap(){document.querySelectorAll('.gr').forEach(r=>{const ok=(!S.g.size||S.g.has(r.dataset.g))&&(!S.s.size||S.s.has(r.dataset.st));r.style.display=ok?'':'none';});
document.querySelectorAll('.ggroup').forEach(h=>{let n=h.nextElementSibling,v=0;while(n&&n.classList.contains('gr')){if(n.style.display!=='none')v++;n=n.nextElementSibling;}h.style.display=v?'':'none';});}
[['#fg','g'],['#fs','s']].forEach(([id,k])=>{const bs=document.querySelectorAll(id+' .fb');bs.forEach(b=>b.onclick=()=>{const v=b.dataset.k,set=S[k];
if(v==='__all__'){set.clear();bs.forEach(x=>x.classList.remove('active'));b.classList.add('active');}
else{set.has(v)?(set.delete(v),b.classList.remove('active')):(set.add(v),b.classList.add('active'));document.querySelector(id+' [data-k="__all__"]').classList.toggle('active',!set.size);}ap();});});})();</script>"""
    updated = datetime.now().strftime("%Y-%m-%d %H:%M")
    hi_n = sum(1 for p in preds if p["conf"] == "高")
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>予測マップ | エンタメ業界トラッカー</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{CSS}{CSS2}{PRED_CSS}</head><body><div class="container">
{nav_html('19-predictions')}
<div class="pghead"><div class="eyebrow">OUTLOOK MAP — 今後起きそうなこと</div><h1>予測マップ</h1>
<p>「今後起きそうなこと」{len(preds)}件を、想定時期のタイムラインで表示。行をクリックすると根拠と検証メモが開きます。最終更新 {updated}</p></div>
<div class="pk"><div><b>{len(preds)}</b><span>予測の総数（確度 高 {hi_n}）</span></div>
<div><b style="color:#e8583a">{stc.get("due", 0)}</b><span>検証待ち（期限を過ぎた）</span></div>
<div><b style="color:#d29922">{len(soon)}</b><span>2ヶ月以内に期限</span></div>
<div><b>{stc.get("live", 0)}</b><span>期間中</span></div><div><b>{stc.get("future", 0)}</b><span>これから始まる</span></div>
<div><b style="color:#2ea043">{stc.get("ok", 0)}</b><span>的中 / 外れ {stc.get("ng", 0)}</span></div></div>
<div class="box"><h2>タイムライン（テーマ別）</h2><p class="lead">バーは想定時期の範囲、灰色の点は記録日、赤い縦線は今日。斜線のバーは時期の記載があいまいなもの（記録日から1年で仮置き）。</p>
<div class="fbar" id="fg">{"".join(fb)}</div><div class="fbar" id="fs">{"".join(sb)}</div>
<div class="gantt"><div class="gin"><div class="gh"><div></div><div class="axis">{"".join(ticks)}<div class="today" style="left:{today_x:.2f}%"></div></div></div>
{"".join(rows)}</div></div>
<div class="lgd"><span><i class="c-hi"></i>確度 高</span><span><i class="c-mid"></i>確度 中</span><span><i class="c-lo"></i>確度 低</span>
<span><i class="st-ok"></i>的中</span><span><i class="st-ng"></i>外れ</span><span><i class="c-mid fz" style="background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.4) 0 3px,transparent 3px 6px)"></i>時期があいまい</span></div></div>
<div class="box"><h2>答え合わせが必要な予測</h2><p class="lead">想定時期を過ぎたもの（赤）と、2ヶ月以内に期限が来るもの（黄）。「今後起きそうなこと」シートの「判定」列に 的中／外れ を入れると、バーの色が変わります。</p>
<div class="due">{duec}{soonc or ""}</div>{'' if (duec or soonc) else '<div class="empty">該当なし</div>'}</div>
<div class="box"><h2>確度 × 時間軸</h2><p class="lead">件数の分布。時間軸は記録日から想定時期の終わりまでの長さで判定。</p><div class="mx">{"".join(cells)}</div></div>
<footer>想定時期の文字列（「6〜12ヶ月」「2027年前半」「年内」など）を期間に自動変換しています</footer>
</div>{js}</body></html>"""
