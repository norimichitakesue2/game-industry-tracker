"""トップページ（ダッシュボード）生成。
xlsx_to_mdhtml.main() から sheet_data を受け取って docs/index.html を書き出す。
各セクションは xlsx の値だけから組み立てるので、日次タスクで自動的に最新化される。
"""
import html
import re
from collections import Counter, OrderedDict
from datetime import date, datetime, timedelta

GENRES = ["ゲーム", "映画", "アニメ", "漫画", "小説", "音楽", "横断"]
WEEKDAYS = "月火水木金土日"
_DPAT = re.compile(r"(\d{4}-\d{2}-\d{2})")


def esc(s):
    return html.escape("" if s is None else str(s))


def dstr(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    m = _DPAT.search(str(v))
    return m.group(1) if m else ""


def todate(s):
    try:
        return date.fromisoformat(s)
    except Exception:
        return None


def jdate(s, with_year=False):
    d = todate(s)
    if not d:
        return esc(s)
    base = f"{d.year}年{d.month}月{d.day}日" if with_year else f"{d.month}/{d.day}"
    return f"{base}（{WEEKDAYS[d.weekday()]}）"


def rows_as_dicts(sheet_data, name):
    if name not in sheet_data:
        return []
    headers, rows = sheet_data[name]
    out = []
    for r in rows:
        out.append({h: r[i] for i, h in enumerate(headers) if h and i < len(r)})
    return out


def gbadge(g):
    g = g or "横断"
    gi = GENRES.index(g) if g in GENRES else 6
    return f'<span class="gb g{gi}">{esc(g)}</span>'


def first_line(s, n=90):
    s = (s or "").strip().replace("\n", " ")
    return s if len(s) <= n else s[:n] + "…"


DASH_CSS = """
<style>
.hero{padding:22px 24px;border-radius:12px;background:linear-gradient(135deg,var(--header-bg),#2a4f8c);color:#fff;margin-bottom:18px}
.hero .eyebrow{font-size:11px;letter-spacing:.12em;opacity:.75;font-weight:700}
.hero h1{font-size:24px;margin:4px 0 6px;color:#fff}
.hero p{margin:0;font-size:13px;opacity:.85}
.hero .cta{display:inline-block;margin-top:12px;background:#fff;color:#1f3864;font-weight:700;font-size:13px;
  padding:7px 14px;border-radius:999px;text-decoration:none}
.hero .cta:hover{opacity:.9}
.anchors{display:flex;flex-wrap:wrap;gap:6px;margin:0 0 18px}
.anchors a{font-size:12px;padding:5px 11px;border:1px solid var(--border);border-radius:999px;color:var(--text);
  text-decoration:none;background:var(--panel)}
.anchors a:hover{border-color:var(--accent);color:var(--accent)}
.dsec{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:18px 20px;margin-bottom:16px}
.dsec>h2{font-size:17px;margin:0 0 4px;display:flex;align-items:center;gap:8px}
.dsec>h2 .mk{display:inline-flex;width:24px;height:24px;border-radius:7px;background:var(--header-bg);color:#fff;
  font-size:12px;align-items:center;justify-content:center}
.dsec .lead{color:var(--muted);font-size:12.5px;margin:0 0 12px}
.sub-h{font-size:12px;font-weight:700;color:var(--muted);letter-spacing:.04em;margin:16px 0 8px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(112px,1fr));gap:8px}
.kpi{border:1px solid var(--border);border-radius:10px;padding:10px 12px;text-decoration:none;color:var(--text);display:block}
.kpi:hover{border-color:var(--accent)}
.kpi .n{font-size:22px;font-weight:800;line-height:1.1}
.kpi .l{font-size:11.5px;color:var(--muted)}
.kpi .d{font-size:11px;margin-top:2px}
.up{color:#3fb950}.down{color:#f85149}.flat{color:var(--muted)}
.gb{display:inline-block;font-size:11px;font-weight:700;padding:1px 8px;border-radius:999px;white-space:nowrap}
.g0{background:rgba(124,92,255,.18);color:#a08cff}.g1{background:rgba(229,83,75,.18);color:#ff7b72}
.g2{background:rgba(240,136,62,.18);color:#f0883e}.g3{background:rgba(46,160,67,.18);color:#3fb950}
.g4{background:rgba(187,128,9,.2);color:#d29922}.g5{background:rgba(219,97,162,.18);color:#f778ba}
.g6{background:rgba(88,166,255,.18);color:#79c0ff}
@media (prefers-color-scheme: light){
  .g0{color:#6639ba}.g1{color:#c93c37}.g2{color:#bc4c00}.g3{color:#1a7f37}.g4{color:#9a6700}.g5{color:#bf3989}.g6{color:#0969da}
  .up{color:#1a7f37}.down{color:#cf222e}
}
.picks{display:grid;gap:10px;grid-template-columns:repeat(auto-fit,minmax(260px,1fr))}
.pick{border:1px solid var(--border);border-radius:10px;padding:12px 14px;display:flex;flex-direction:column;gap:6px}
.pick .no{font-size:11px;color:var(--muted);font-weight:700}
.pick a.t{font-weight:700;font-size:14px;line-height:1.45;color:var(--text);text-decoration:none}
.pick a.t:hover{color:var(--accent)}
.pick .s{font-size:12.5px;color:var(--muted);line-height:1.55}
.glist{display:grid;gap:10px;grid-template-columns:repeat(auto-fit,minmax(300px,1fr))}
.gcol{border:1px solid var(--border);border-radius:10px;padding:10px 12px}
.gcol .gh{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.gcol .gh .c{font-size:11.5px;color:var(--muted)}
.gcol ul{margin:0;padding-left:16px;font-size:12.5px;line-height:1.55}
.gcol li{margin:3px 0}
.gcol li a{color:var(--text);text-decoration:none}
.gcol li a:hover{color:var(--accent)}
.gcol details summary{font-size:11.5px;color:var(--accent);cursor:pointer;margin-top:4px;list-style:none}
.trend{display:flex;flex-direction:column;gap:5px}
.trow{display:grid;grid-template-columns:92px 1fr 34px;align-items:center;gap:8px;font-size:11.5px}
.trow .bar{display:flex;height:14px;border-radius:4px;overflow:hidden;background:rgba(127,127,127,.12)}
.trow .bar span{display:block;height:100%}
.b0{background:#8b6cff}.b1{background:#e5534b}.b2{background:#f0883e}.b3{background:#2ea043}.b4{background:#d29922}.b5{background:#db61a2}.b6{background:#58a6ff}
.legend{display:flex;flex-wrap:wrap;gap:10px;font-size:11px;color:var(--muted);margin-top:6px}
.legend i{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:4px;vertical-align:middle}
.chiprow{display:flex;flex-wrap:wrap;gap:6px}
.chip{font-size:12px;border:1px solid var(--border);border-radius:999px;padding:3px 10px;color:var(--text);text-decoration:none;background:transparent}
.chip b{margin-left:4px}
.cards{display:grid;gap:10px;grid-template-columns:repeat(auto-fit,minmax(280px,1fr))}
.card{border:1px solid var(--border);border-radius:10px;padding:12px 14px;font-size:12.5px;line-height:1.6}
.card .ct{font-weight:700;font-size:13.5px;line-height:1.5;margin-bottom:6px}
.card .cm{color:var(--muted);font-size:11.5px;margin-top:6px}
.badge{display:inline-block;font-size:10.5px;font-weight:700;padding:1px 7px;border-radius:5px;border:1px solid var(--border);margin-right:4px}
.badge.hi{border-color:#3fb950;color:#3fb950}.badge.mid{border-color:#d29922;color:#d29922}
.ev{display:grid;grid-template-columns:96px 1fr;gap:10px;padding:8px 0;border-top:1px solid var(--border);font-size:12.5px}
.ev:first-child{border-top:none}
.ev .when{font-weight:700}
.ev .when small{display:block;font-weight:400;color:var(--muted);font-size:11px}
.ev .what a{color:var(--text);text-decoration:none;font-weight:600}
.ev .what a:hover{color:var(--accent)}
.ev .what .x{color:var(--muted);font-size:11.5px}
.now{color:#3fb950;font-weight:700;font-size:10.5px;margin-left:6px}
.lib{display:grid;gap:10px;grid-template-columns:repeat(auto-fit,minmax(200px,1fr))}
.libc{border:1px solid var(--border);border-radius:10px;padding:10px 12px}
.libc .lt{font-weight:700;font-size:13px;margin-bottom:6px}
.libc a{display:flex;justify-content:space-between;font-size:12.5px;color:var(--text);text-decoration:none;padding:2px 0}
.libc a:hover{color:var(--accent)}
.libc a span{color:var(--muted);font-variant-numeric:tabular-nums}
.libc .dot{color:#e8583a;font-size:9px;vertical-align:super}
.months details{border-top:1px solid var(--border);padding:8px 0}
.months details:first-child{border-top:none}
.months summary{cursor:pointer;font-weight:700;font-size:13.5px}
.months summary span{color:var(--muted);font-weight:400;font-size:12px;margin-left:6px}
.dayrow{display:grid;grid-template-columns:96px 52px 1fr;gap:8px;font-size:12.5px;padding:5px 0 5px 4px}
.dayrow a{color:var(--text);text-decoration:none}
.dayrow a:hover{color:var(--accent)}
.dayrow .cnt{color:var(--muted);text-align:right}
.tgrid2{display:grid;gap:10px;grid-template-columns:repeat(auto-fill,minmax(280px,1fr))}
.tcard2{display:block;border:1px solid var(--border);border-radius:10px;padding:12px 14px;text-decoration:none;color:var(--text)}
.tcard2:hover{border-color:var(--accent)}
.tcard2 .tn{font-weight:800;font-size:14px}
.tcard2 .tq{color:var(--muted);font-size:12px;margin:2px 0 6px}
.tcard2 .tv{font-size:12px;line-height:1.6;display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.tcard2 .tm{font-size:11px;color:var(--muted);margin-top:8px}
.newb{font-size:10px;background:#e8583a;color:#fff;border-radius:5px;padding:1px 6px;vertical-align:middle}
.plist{display:flex;flex-direction:column;gap:8px}
.prow{display:grid;grid-template-columns:1fr 160px 130px;gap:12px;align-items:center;font-size:12.5px}
.prow .pbar{height:8px;border-radius:4px;background:rgba(127,127,127,.18);overflow:hidden}
.prow .pbar span{display:block;height:100%}
.prow .pl{font-weight:700;font-size:12px}
.prow .pl small{display:block;font-weight:400;color:var(--muted);font-size:10.5px}
@media (max-width:640px){
  .prow{grid-template-columns:1fr;gap:4px}
  body{padding:14px}
  .hero{padding:18px}
  .dayrow{grid-template-columns:72px 40px 1fr}
  .ev{grid-template-columns:80px 1fr}
  .trow{grid-template-columns:70px 1fr 28px}
}
</style>
"""


def render_dashboard_html(sheet_data, stats, groups, slugs, short, recent_slugs, nav_html, css):
    news = rows_as_dicts(sheet_data, "日次ニュース")
    for r in news:
        r["_d"] = dstr(r.get("日付"))
    dates = sorted({r["_d"] for r in news if r["_d"]}, reverse=True)
    latest = dates[0] if dates else ""
    prev = dates[1] if len(dates) > 1 else ""
    today_rows = [r for r in news if r["_d"] == latest]
    prev_rows = [r for r in news if r["_d"] == prev]

    # ---- 今日のKPI（ジャンル別件数と前日比）
    cnt = Counter((r.get("ジャンル") or "横断") for r in today_rows)
    pcnt = Counter((r.get("ジャンル") or "横断") for r in prev_rows)
    kpis = [f'<a class="kpi" href="01-daily-news.html#d-{latest}"><div class="n">{len(today_rows)}</div>'
            f'<div class="l">本日の追加件数</div>{_delta(len(today_rows), len(prev_rows))}</a>']
    for g in GENRES:
        kpis.append(f'<a class="kpi" href="01-daily-news.html#d-{latest}"><div class="n">{cnt.get(g, 0)}</div>'
                    f'<div class="l">{gbadge(g)}</div>{_delta(cnt.get(g, 0), pcnt.get(g, 0))}</a>')

    # ---- 注目3件：備考に「★」があればそれを優先、なければタグで重み付け
    starred = [r for r in today_rows if "★" in str(r.get("備考") or "")]
    if len(starred) < 3:
        weight = {"M&A": 3, "規制": 3, "業績": 2, "適時開示": 2, "人事・組織": 2, "AI": 1, "出資": 2, "提携": 1}
        rest = [r for r in today_rows if r not in starred]
        seen_g = {r.get("ジャンル") for r in starred}
        rest.sort(key=lambda r: -sum(weight.get(t.strip(), 0) for t in str(r.get("タグ") or "").split(",")))
        for r in rest:
            if len(starred) >= 3:
                break
            if r.get("ジャンル") in seen_g:
                continue
            starred.append(r)
            seen_g.add(r.get("ジャンル"))
    picks = []
    for i, r in enumerate(starred[:3], 1):
        tags = "".join(f'<span class="tag">{esc(t.strip())}</span>' for t in str(r.get("タグ") or "").split(",") if t.strip())
        picks.append(f'<div class="pick"><div class="no">PICK {i} {gbadge(r.get("ジャンル"))}</div>'
                     f'<a class="t" href="{esc(r.get("URL"))}" target="_blank" rel="noopener">{esc(r.get("見出し"))}</a>'
                     f'<div class="s">{esc(first_line(r.get("今後の予測"), 110))}</div><div>{tags}</div></div>')

    # ---- ジャンル別ヘッドライン
    gcols = []
    for g in GENRES:
        items = [r for r in today_rows if (r.get("ジャンル") or "横断") == g]
        if not items:
            continue
        li = lambda r: f'<li><a href="{esc(r.get("URL"))}" target="_blank" rel="noopener">{esc(r.get("見出し"))}</a></li>'
        head = "".join(li(r) for r in items[:3])
        more = ""
        if len(items) > 3:
            more = (f'<details><summary>ほか{len(items) - 3}件を表示</summary><ul>'
                    + "".join(li(r) for r in items[3:]) + "</ul></details>")
        gcols.append(f'<div class="gcol"><div class="gh">{gbadge(g)}<span class="c">{len(items)}件</span></div>'
                     f'<ul>{head}</ul>{more}</div>')

    # ---- 直近7日のジャンル推移
    last7 = sorted(dates[:7])
    mx = max([sum(1 for r in news if r["_d"] == d) for d in last7] or [1])
    trows = []
    for d in last7:
        c = Counter((r.get("ジャンル") or "横断") for r in news if r["_d"] == d)
        tot = sum(c.values())
        segs = "".join(f'<span class="b{GENRES.index(g)}" style="width:{c.get(g, 0) / mx * 100:.1f}%" title="{g} {c.get(g, 0)}件"></span>'
                       for g in GENRES if c.get(g))
        trows.append(f'<div class="trow"><a href="01-daily-news.html#d-{d}" style="color:var(--text);text-decoration:none">{jdate(d)}</a>'
                     f'<div class="bar">{segs}</div><div style="text-align:right">{tot}</div></div>')
    legend = "".join(f'<span><i class="b{i}"></i>{g}</span>' for i, g in enumerate(GENRES))

    # ---- タグの増減（直近7日 vs その前7日）
    w1 = set(dates[:7]); w0 = set(dates[7:14])
    t1 = Counter(); t0 = Counter()
    for r in news:
        tags = [t.strip() for t in str(r.get("タグ") or "").split(",") if t.strip()]
        for t in tags:
            if t in ("国内", "海外"):
                continue
            if r["_d"] in w1:
                t1[t] += 1
            elif r["_d"] in w0:
                t0[t] += 1
    tagchips = []
    for t, n in t1.most_common(14):
        diff = n - t0.get(t, 0)
        cls = "up" if diff > 0 else ("down" if diff < 0 else "flat")
        sign = f"+{diff}" if diff > 0 else str(diff)
        tagchips.append(f'<span class="chip">{esc(t)}<b>{n}</b> <span class="{cls}">{sign}</span></span>')

    # ---- その他シートの「本日分」
    def today_of(name, col="記録日"):
        return [r for r in rows_as_dicts(sheet_data, name) if dstr(r.get(col)) == latest]

    qs = today_of("横断的問い")
    if not qs:
        qs = rows_as_dicts(sheet_data, "横断的問い")[:3]
    qcards = "".join(
        f'<div class="card"><div class="ct">{esc(first_line(q.get("問い"), 150))}</div>'
        f'<div>{gbadge(q.get("ジャンル"))} <span class="badge">{esc(q.get("ステータス") or "")}</span></div>'
        f'<div class="cm">発火元: {esc(first_line(q.get("発火元（ニュース等）"), 80))}</div></div>' for q in qs[:4])

    fs = rows_as_dicts(sheet_data, "今後起きそうなこと")[:3]
    fcards = "".join(
        f'<div class="card"><div class="ct">{esc(first_line(f.get("今後起きそうなこと"), 150))}</div>'
        f'<div><span class="badge {"hi" if f.get("確度") == "高" else "mid"}">確度 {esc(f.get("確度"))}</span>'
        f'<span class="badge">{esc(f.get("想定時期"))}</span>{gbadge(f.get("ジャンル"))}</div>'
        f'<div class="cm">記録日 {esc(dstr(f.get("記録日")))} · 検証: {esc(first_line(f.get("検証メモ"), 70))}</div></div>' for f in fs)

    # ---- 論点テーマ（問い・掘り下げ）と予測の進行状況
    deephtml, predhtml = "", ""
    try:
        import insights
        themes, items, preds = insights.collect(sheet_data)
        ld = todate(latest) or date.today()
        tl = []
        for t in themes:
            its = items.get(t["id"], [])
            if its and t["id"] != "T99":
                tl.append((max(x["d"] for x in its), t, its))
        tl.sort(key=lambda x: x[0], reverse=True)
        pc = Counter(p["tid"] for p in preds)
        tc = []
        for last, t, its in tl[:6]:
            newn = sum(1 for x in its if x["d"] and (ld - todate(x["d"])).days <= 3)
            nq = sum(1 for x in its if x["kind"] == "問い")
            tc.append(f'<a class="tcard2" href="18-deep.html#{t["id"]}"><div class="tn">{esc(t["name"])}'
                      + (f' <span class="newb">NEW {newn}</span>' if newn else "") + "</div>"
                      f'<div class="tq">{esc(t["q"])}</div><div class="tv">{esc(t["view"])}</div>'
                      f'<div class="tm">問い {nq} · 掘り下げ {len(its) - nq} · 予測 {pc.get(t["id"], 0)} · 最終 {last[5:]}</div></a>')
        deephtml = f'<div class="tgrid2">{"".join(tc)}</div>'
        cur = ld.year * 12 + ld.month - 1
        st = Counter()
        rowsp = []
        for p in preds:
            s0 = p["s"][0] * 12 + p["s"][1] - 1; e0 = p["e"][0] * 12 + p["e"][1] - 1
            v = p.get("verdict", "")
            k = "ok" if v.startswith("的中") else ("ng" if v.startswith("外れ") else ("due" if e0 < cur else ("live" if s0 <= cur else "future")))
            st[k] += 1
            p["_k"], p["_e0"], p["_s0"] = k, e0, s0
        near = sorted([p for p in preds if p["_k"] in ("due", "live")], key=lambda p: p["_e0"])
        due = [p for p in near if p["_k"] == "due"][:3]
        live = [p for p in near if p["_k"] == "live"][:5]
        for p in due + live:
            r = todate(p["d"]) or ld
            r0 = r.year * 12 + r.month - 1
            tot = max(p["_e0"] + 1 - r0, 1)
            el = min(max(cur - r0 + ld.day / 31, 0), tot)
            pctv = el / tot * 100
            left = p["_e0"] - cur
            lab = ("期限超過・要判定" if p["_k"] == "due" else (f"残り約{left}ヶ月" if left > 0 else "今月が期限"))
            col = "#e8583a" if p["_k"] == "due" else ("#d29922" if left <= 2 else "#58a6ff")
            rowsp.append(f'<div class="prow"><div class="pt">{gbadge(p["g"])} {esc(first_line(p["text"], 70))}</div>'
                         f'<div class="pbar"><span style="width:{pctv:.0f}%;background:{col}"></span></div>'
                         f'<div class="pl" style="color:{col}">{lab}<small>{esc(p["when"])} · 確度{esc(p["conf"])}</small></div></div>')
        predhtml = (f'<div class="kpis"><div class="kpi"><div class="n">{len(preds)}</div><div class="l">予測の総数</div></div>'
                    f'<div class="kpi"><div class="n" style="color:#e8583a">{st["due"]}</div><div class="l">検証待ち（期限超過）</div></div>'
                    f'<div class="kpi"><div class="n">{st["live"]}</div><div class="l">期間中</div></div>'
                    f'<div class="kpi"><div class="n">{st["future"]}</div><div class="l">これから</div></div>'
                    f'<div class="kpi"><div class="n" style="color:#3fb950">{st["ok"]}</div><div class="l">的中（外れ {st["ng"]}）</div></div></div>'
                    f'<div class="sub-h">期限が近い予測（バーは記録日から期限までの経過）</div><div class="plist">{"".join(rowsp)}</div>')
    except Exception as e:
        import sys, traceback; traceback.print_exc()

    mm = today_of("メディアミックス追跡")
    mmcards = "".join(
        f'<div class="card"><div class="ct">{esc(m.get("原作名"))}</div>'
        f'<div>{gbadge(m.get("原作ジャンル"))} → {gbadge(m.get("展開先ジャンル"))} <span class="badge">{esc(m.get("ステータス"))}</span></div>'
        f'<div style="margin-top:6px">{esc(first_line(m.get("展開内容"), 90))}</div>'
        f'<div class="cm">公開/開始予定: {esc(m.get("公開/開始予定") or "未定")}</div></div>' for m in mm)

    gd = today_of("ゲームデザイン・トレンド")
    gdcards = "".join(
        f'<div class="card"><div class="ct">{esc(x.get("デザイン名/パターン"))}</div>'
        f'<div><span class="badge">{esc(x.get("分類"))}</span><span class="badge">{esc(x.get("ステータス"))}</span>'
        f'<span class="badge">注目度 {esc(x.get("注目度(1-5)"))}</span></div>'
        f'<div class="cm">発火元: {esc(first_line(x.get("発火元ニュース"), 80))}</div></div>' for x in gd)

    # ---- 直近イベント（今日〜14日後＋開催中）
    base = todate(latest) or date.today()
    evs = []
    for e in rows_as_dicts(sheet_data, "イベントカレンダー"):
        s = todate(dstr(e.get("開催日")))
        en = todate(dstr(e.get("終了日"))) or s
        if not s or e.get("ステータス") == "中止":
            continue
        if en >= base and s <= base + timedelta(days=14):
            evs.append((s, en, e))
    evs.sort(key=lambda x: x[0])
    evhtml = []
    for s, en, e in evs[:14]:
        ongoing = s <= base <= en
        rng = f"〜{en.month}/{en.day}" if en and en != s else ""
        url = e.get("公式/ソースURL") or e.get("公式・ソースURL") or ""
        name = esc(e.get("イベント名"))
        name = f'<a href="{esc(url)}" target="_blank" rel="noopener">{name}</a>' if url else name
        evhtml.append(f'<div class="ev"><div class="when">{s.month}/{s.day}（{WEEKDAYS[s.weekday()]}）<small>{rng}</small></div>'
                      f'<div class="what">{name}{"<span class=now>開催中</span>" if ongoing else ""}'
                      f'<div class="x">{esc(e.get("種別"))} · {esc(first_line(e.get("概要・注目点"), 70))}</div></div></div>')

    # ---- ライブラリ
    libs = []
    for gname, sheets in groups:
        links = []
        for sn in sheets:
            slug = slugs.get(sn)
            if not slug:
                continue
            dot = ' <span class="dot">●</span>' if slug in recent_slugs else ""
            links.append(f'<a href="{slug}.html">{esc(short(sn))}{dot}<span>{stats.get(sn, 0):,}</span></a>')
        libs.append(f'<div class="libc"><div class="lt">{esc(gname)}</div>{"".join(links)}</div>')
    libs.append('<div class="libc"><div class="lt">マップ</div><a href="industry-map.html">全体連動マップ<span>→</span></a>'
                '<a href="design-map.html">デザイン連動マップ<span>→</span></a></div>')

    # ---- 日次一覧（月アコーディオン）
    bymonth = OrderedDict()
    for d in dates:
        bymonth.setdefault(d[:7], []).append(d)
    months = []
    for i, (ym, ds) in enumerate(bymonth.items()):
        total = sum(1 for r in news if r["_d"][:7] == ym)
        rowsh = []
        for d in ds:
            dr = [r for r in news if r["_d"] == d]
            lead = dr[0].get("見出し") if dr else ""
            rowsh.append(f'<div class="dayrow"><a href="01-daily-news.html#d-{d}">{jdate(d)}</a>'
                         f'<span class="cnt">{len(dr)}本</span><a href="01-daily-news.html#d-{d}">{esc(first_line(lead, 60))}</a></div>')
        y, m = ym.split("-")
        months.append(f'<details{" open" if i == 0 else ""}><summary>{int(y)}年{int(m)}月<span>{len(ds)}日 · {total}本</span></summary>'
                      + "".join(rowsh) + "</details>")

    updated = datetime.now().strftime("%Y-%m-%d %H:%M")
    total_rows = sum(stats.values())

    def sec(anchor, mk, title, lead, inner):
        return (f'<div class="dsec" id="{anchor}"><h2><span class="mk">{mk}</span>{title}</h2>'
                + (f'<p class="lead">{lead}</p>' if lead else "") + inner + "</div>")

    other = ""
    if qcards:
        other += f'<div class="sub-h">新しい横断的問い（{len(qs[:4])}）</div><div class="cards">{qcards}</div>'
    if mmcards:
        other += f'<div class="sub-h">メディアミックスの動き（{len(mm)}）</div><div class="cards">{mmcards}</div>'
    if gdcards:
        other += f'<div class="sub-h">ゲームデザインの型（{len(gd)}）</div><div class="cards">{gdcards}</div>'

    body = [
        f'''<div class="hero"><div class="eyebrow">ENTERTAINMENT INDUSTRY TRACKER</div>
<h1>エンタメ業界・時事ダッシュボード</h1>
<p>ゲーム・映画・アニメ・漫画・小説・音楽の6ジャンルを毎朝集約 · 最終更新 {updated} · 全{total_rows:,}行</p>
<a class="cta" href="01-daily-news.html#d-{latest}">最新の日次ニュース {jdate(latest)} を読む →</a></div>''',
        '<div class="anchors"><a href="#today">★ 今日の変化</a><a href="#deep">問い・掘り下げ</a><a href="#future">予測</a><a href="#trend">推移</a>'
        '<a href="#events">イベント</a><a href="#library">ライブラリ</a><a href="#daily">日次一覧</a></div>',
        sec("today", "★", f"今日の変化 · {jdate(latest)}",
            f"前回 {jdate(prev)} からの追加分。件数をクリックするとその日のニュースへ移動します。",
            f'<div class="kpis">{"".join(kpis)}</div>'
            f'<div class="sub-h">注目の3件</div><div class="picks">{"".join(picks)}</div>'
            f'<div class="sub-h">ジャンル別ヘッドライン</div><div class="glist">{"".join(gcols)}</div>'
            + other),
        sec("deep", "?", "問いで追う論点", '横断的問いと掘り下げを論点テーマごとに整理（最近動いた順）。<a href="18-deep.html">すべてのテーマ →</a>',
            deephtml),
        sec("future", "🔮", "予測の進行状況", '<a href="19-predictions.html">予測マップ（タイムライン）を開く →</a>',
            predhtml + f'<div class="sub-h">最新の予測</div><div class="cards">{fcards}</div>'),
        sec("trend", "📈", "直近7日の推移", "日次ニュースのジャンル構成と、よく付いたタグ（直近7日、カッコ内は前の7日との差）。",
            f'<div class="trend">{"".join(trows)}</div><div class="legend">{legend}</div>'
            f'<div class="sub-h">タグの増減</div><div class="chiprow">{"".join(tagchips)}</div>'),
        sec("events", "📅", "直近2週間のイベント", '<a href="14-events.html">イベントカレンダー全体 →</a>',
            "".join(evhtml) or '<div class="empty">予定なし</div>'),
        sec("library", "◎", "ライブラリ", "シート別の蓄積件数。● は直近3日以内に更新。", f'<div class="lib">{"".join(libs)}</div>'),
        sec("daily", "D", "日次ニュース一覧", f"全{len(dates)}日 · 新しい順", f'<div class="months">{"".join(months)}</div>'),
    ]
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>エンタメ業界・時事ダッシュボード</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
{css}{DASH_CSS}</head><body><div class="container">
{nav_html('00-readme')}
{body[0]}
{''.join(body[1:])}
<footer>自動生成 · GitHub Pages · 毎朝の日次タスクで更新</footer>
</div></body></html>"""


def _delta(now, before):
    d = now - before
    if d > 0:
        return f'<div class="d up">▲{d} 前日比</div>'
    if d < 0:
        return f'<div class="d down">▼{-d} 前日比</div>'
    return '<div class="d flat">± 0</div>'
