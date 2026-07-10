#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gen_maps.py — xlsx の各シートから 2 つの「連動マップ」ページを自動生成する。
  docs/design-map.html    ... ゲームデザイン: 連動マップ + ライフサイクル俯瞰
  docs/industry-map.html  ... 業界全体: 全体連動マップ + 予測マップ

- ノードは各シートから自動生成（新しい行を足せば翌日から自動で反映）。
- エッジ（つながり）は判断が要るため、DRIVER→対象のキーワード規則で半自動生成。
  新しい型を配線したい時は、下部の *_EDGES にルールを1行足すだけ。
- 座標は「レーンごとに均等割り付け」で自動計算（手打ち座標なし）。
実行: リポジトリルートで `python scripts/gen_maps.py`（xlsx_to_mdhtml から自動呼び出しされる）
"""
import html, re
from pathlib import Path
from openpyxl import load_workbook

REPO = Path(__file__).resolve().parents[1]
XLSX = REPO / "master" / "ゲーム業界・時事情報収集.xlsx"
DOCS = REPO / "docs"

# 既存サイトと同じ CSS / nav を流用（実行時 import で循環回避）
def _shared():
    from xlsx_to_mdhtml import CSS, nav_html
    return CSS, nav_html

def esc(v):
    return html.escape("" if v is None else str(v))

def rows(wb, sn):
    ws = wb[sn]; h = [c.value for c in ws[1]]
    while h and h[-1] is None: h.pop()
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        r = list(r[:len(h)])
        if any(v not in (None, "") for v in r):
            out.append(dict(zip(h, r)))
    return out

# ---- 状態→色（ライフサイクル） ----
STATUS_COLORS = {
    "up":     ("#059669", "#ecfdf5", "#065f46", "台頭"),
    "hot":    ("#2563eb", "#eff6ff", "#1e40af", "流行中"),
    "mature": ("#d97706", "#fff7ed", "#9a3412", "成熟/逆風"),
    "decline":("#94a3b8", "#f1f5f9", "#475569", "衰退"),
    "reeval": ("#7c3aed", "#f5f3ff", "#5b21b6", "再評価"),
}
def status_group(s):
    s = str(s or "")
    if "衰退" in s: return "decline"
    if "再評価" in s or "復権" in s or "過熱" in s: return "reeval"
    if "逆風" in s or "規制" in s or ("成熟" in s and "流行" not in s): return "mature"
    if "流行" in s: return "hot"
    if "台頭" in s: return "up"
    if "成熟" in s: return "mature"
    return "hot"

# 状態→ライフサイクルX位置（1=台頭 .. 4=衰退, 再評価は復路3.4）
STATUS_X = {"up":1.0, "hot":2.0, "mature":3.0, "decline":4.0, "reeval":3.4}

EDGE_COLORS = {"+":"#059669", "-":"#dc2626", "±":"#94a3b8"}

def bezier(x1,y1,x2,y2):
    xm=(x1+x2)/2
    return f'M{x1:.0f} {y1:.0f} C{xm:.0f} {y1:.0f} {xm:.0f} {y2:.0f} {x2:.0f} {y2:.0f}'

def node_box(x,y,w,h,fill,stroke,title,sub,key,tw=None):
    tw = tw or (w-24)
    t = esc(title)
    s = f'<text x="{x+12}" y="{y+h-13}" font-size="10.5" font-weight="400" fill="{stroke}">{esc(sub)}</text>' if sub else ""
    return (f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')">'
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="9" fill="{fill}" stroke="{stroke}" stroke-width="1.6"/>'
            f'<text x="{x+12}" y="{y+22}" font-size="12" font-weight="700" fill="#0f172a" textLength="{tw}" lengthAdjust="spacingAndGlyphs">{t}</text>'
            f'{s}</g>')

def driver_box(x,y,w,h,title,sub,key):
    return (f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')">'
            f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="9" fill="#fff" stroke="#cbd5e1" stroke-width="1.5"/>'
            f'<text x="{x+12}" y="{y+22}" font-size="12.5" font-weight="800" fill="#0f172a">{esc(title)}</text>'
            f'<text x="{x+12}" y="{y+h-12}" font-size="10.5" fill="#64748b">{esc(sub)}</text></g>')

def legend_life(x,y):
    out=[f'<text x="{x}" y="{y}" font-size="11.5" font-weight="700" fill="#475569">状態:</text>']
    xx=x+42
    for k in ["up","hot","mature","decline","reeval"]:
        c,bg,_,lab=STATUS_COLORS[k]
        out.append(f'<rect x="{xx}" y="{y-11}" width="12" height="12" rx="3" fill="{bg}" stroke="{c}"/><text x="{xx+16}" y="{y}" font-size="11.5" fill="#475569">{esc(lab)}</text>')
        xx+=len(lab)*13+34
    return "".join(out)

def legend_edge(x,y):
    out=[f'<text x="{x}" y="{y}" font-size="11.5" font-weight="700" fill="#475569">影響:</text>']
    xx=x+42
    for sign,lab in [("+","追い風(＋)"),("-","逆風(−)"),("±","反動/両面(±)")]:
        c=EDGE_COLORS[sign]; dash=' stroke-dasharray="5 4"' if sign=="±" else ""
        out.append(f'<line x1="{xx}" y1="{y-4}" x2="{xx+22}" y2="{y-4}" stroke="{c}" stroke-width="3"{dash}/><text x="{xx+28}" y="{y}" font-size="11.5" fill="#475569">{esc(lab)}</text>')
        xx+=len(lab)*13+40
    return "".join(out)


# ============================================================
# ページシェル（既存サイトのCSS/navを流用）
# ============================================================
DETAIL_JS = """
<script>
var MDATA = __MDATA__;
function mShow(k){
  var d = MDATA[k]; if(!d) return;
  var p = document.getElementById('mdetail');
  var rows = (d.rows||[]).map(function(r){return '<div class="mdrow"><span>'+r[0]+'</span><b>'+r[1]+'</b></div>';}).join('');
  p.innerHTML = '<button class="mdx" onclick="mHide()">×</button><div class="mdt">'+d.t+'</div>'+(d.s?'<div class="mds">'+d.s+'</div>':'')+rows;
  p.classList.add('on');
}
function mHide(){document.getElementById('mdetail').classList.remove('on');}
</script>
"""

PAGE_CSS = """
<style>
.mapwrap{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:8px;overflow-x:auto}
.maplead{color:var(--muted);font-size:13.5px;margin:2px 0 14px}
.mnode:hover rect{stroke-width:2.4}
#mdetail{position:fixed;right:18px;bottom:18px;width:min(360px,90vw);background:var(--panel);border:1px solid var(--border);
  border-radius:12px;box-shadow:0 16px 40px rgba(0,0,0,.28);padding:16px 16px 14px;opacity:0;pointer-events:none;transform:translateY(8px);transition:.15s;z-index:50}
#mdetail.on{opacity:1;pointer-events:auto;transform:none}
.mdx{position:absolute;top:9px;right:11px;border:none;background:var(--border);color:var(--text);width:26px;height:26px;border-radius:50%;font-size:16px;cursor:pointer}
.mdt{font-weight:800;font-size:14.5px;padding-right:26px;line-height:1.4}
.mds{font-size:12.5px;color:var(--muted);margin:6px 0 8px;line-height:1.55}
.mdrow{display:grid;grid-template-columns:88px 1fr;gap:8px;font-size:12.5px;padding:4px 0;border-top:1px solid var(--border)}
.mdrow span{color:var(--muted)}
.viewtabs{display:flex;gap:8px;margin:6px 0 14px;flex-wrap:wrap}
</style>
"""

def page(title, subtitle, active_nav, sections_html, mdata_json):
    CSS, nav_html = _shared()
    js = DETAIL_JS.replace("__MDATA__", mdata_json)
    body = []
    for sec in sections_html:
        body.append(sec)
    return f"""<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)}</title>
{CSS}{PAGE_CSS}</head><body><div class="container">
<header><h1>🎮 {esc(title)}</h1><span class="subtle" style="color:var(--muted);font-size:13px">{esc(subtitle)}</span></header>
{nav_html()}
{''.join(body)}
<div id="mdetail"></div>
</div>{js}</body></html>"""

def section(num, heading, lead, svg):
    return (f'<div class="section"><h2><span class="num" style="display:inline-flex;width:24px;height:24px;border-radius:7px;'
            f'background:var(--header-bg);color:#fff;font-size:13px;align-items:center;justify-content:center;margin-right:8px">{num}</span>{esc(heading)}</h2>'
            f'<p class="maplead">{lead}</p><div class="mapwrap">{svg}</div></div>')

# ============================================================
# 1) ゲームデザイン ページ
# ============================================================
DESIGN_DRIVERS = [
    ("生成AIの実用化", "低コスト化・AI支援"),
    ("GaaS逆風・選別", "ライブサービスの淘汰"),
    ("課金・年齢規制", "ルートボックス/DFA"),
    ("再編・レイオフ", "プラットフォーマー"),
    ("コスト高・値上げ", "部材・人件費"),
    ("ユーザーの信頼志向", "オーセンティシティ"),
]
DESIGN_EDGES = [
    (0,"AI NPC","+"),(0,"少人数","+"),(0,"オーセンティシティ","±"),
    (1,"エバーグリーン","+"),(1,"ライブサービス","-"),(1,"シングルプレイ大作","+"),(1,"サバイバルクラフト","+"),
    (2,"有料ランダム","-"),(2,"エバーグリーン","+"),
    (3,"シングルプレイ大作","+"),(3,"ライブサービス","-"),
    (4,"エバーグリーン","+"),(4,"少人数","+"),(4,"復刻","+"),
    (5,"オーセンティシティ","+"),(5,"メタバース","-"),(5,"有料ランダム","-"),
]

def build_design(wb, mdata):
    ds = rows(wb, "ゲームデザイン・トレンド")
    # 状態グループ→表示順
    order = {"up":0,"hot":1,"reeval":2,"mature":3,"decline":4}
    for d in ds:
        d["_g"] = status_group(d.get("ステータス"))
        try: d["_att"] = int(d.get("注目度(1-5)") or 0)
        except: d["_att"] = 0
    ds.sort(key=lambda d:(order.get(d["_g"],9), -d["_att"]))

    # ---- View A: 連動マップ ----
    n = len(ds); top=64; step=56; dh=50
    dW, dX = 322, 458
    drvX, drvW, drvH = 20, 176, 46
    height = top + (n-1)*step + dh + 60
    ycent = [top + i*step + dh/2 for i in range(n)]
    # driver centers spread over same span
    dn = len(DESIGN_DRIVERS)
    dspan_top, dspan_bot = ycent[0], ycent[-1]
    drv_c = [dspan_top + (dspan_bot-dspan_top)*i/(dn-1) for i in range(dn)]

    svg=[f'<svg viewBox="0 0 {dX+dW+120} {height}" xmlns="http://www.w3.org/2000/svg" role="img" style="width:100%;height:auto;min-width:820px" font-family="-apple-system,\'Hiragino Sans\',\'Noto Sans JP\',sans-serif">']
    svg.append('<title>ゲームデザイン連動マップ</title>')
    svg.append(f'<text x="{drvX+8}" y="46" font-size="12" font-weight="700" fill="#64748b">外部ドライバー</text>')
    svg.append(f'<text x="{dX+8}" y="46" font-size="12" font-weight="700" fill="#64748b">デザインの型（色＝状態 / 数字＝注目度）</text>')
    # edges
    def find_idx(sub):
        for i,d in enumerate(ds):
            if sub in str(d.get("デザイン名/パターン","")): return i
        return None
    epaths=[]
    for di,sub,sign in DESIGN_EDGES:
        ti=find_idx(sub)
        if ti is None or di>=dn: continue
        c=EDGE_COLORS[sign]; dash=' stroke-dasharray="5 4"' if sign=="±" else ""
        epaths.append(f'<path d="{bezier(drvX+drvW, drv_c[di], dX, ycent[ti])}" fill="none" stroke="{c}" stroke-width="2.4"{dash} opacity="0.85"/>')
    svg.append('<g>'+''.join(epaths)+'</g>')
    # driver nodes
    for i,(t,s) in enumerate(DESIGN_DRIVERS):
        y=drv_c[i]-drvH/2
        key=f"gd{i}"
        svg.append(driver_box(drvX,y,drvW,drvH,t,s,key))
        mdata[key]={"t":esc(t),"s":esc(s),"rows":[["役割","デザイン潮流を動かす外部要因"]]}
    # design nodes
    for i,d in enumerate(ds):
        g=d["_g"]; c,bg,tc,lab=STATUS_COLORS[g]
        y=top+i*step
        name=str(d.get("デザイン名/パターン",""))
        short=name if len(name)<=22 else name[:21]+"…"
        att=d["_att"]; key=f"d{i}"
        rep=str(d.get("代表タイトル") or "").replace("、",",").replace("，",",").split(",")[0].strip()
        rep=rep if len(rep)<=22 else rep[:21]+"…"
        svg.append(f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')">'
                   f'<rect x="{dX}" y="{y}" width="{dW}" height="{dh}" rx="9" fill="{bg}" stroke="{c}" stroke-width="1.6"/>'
                   f'<text x="{dX+12}" y="{y+18}" font-size="11.5" font-weight="700" fill="{tc}">{esc(short)}</text>'
                   f'<text x="{dX+12}" y="{y+30}" font-size="10" fill="{c}">{esc(lab)}</text>'
                   f'<text x="{dX+12}" y="{y+44}" font-size="10.5" fill="#64748b">代表: {esc(rep)}</text>'
                   f'<circle cx="{dX+dW-20}" cy="{y+dh/2}" r="12" fill="{c}"/>'
                   f'<text x="{dX+dW-20}" y="{y+dh/2+4}" font-size="12" font-weight="800" fill="#fff" text-anchor="middle">{att}</text></g>')
        mdata[key]={"t":esc(name),"s":esc(d.get("概要・特徴","")),
            "rows":[["状態",esc(d.get("ステータス",""))],["分類",esc(d.get("分類",""))],
                    ["注目度",str(att)+" / 5"],["代表",esc(d.get("代表タイトル",""))],
                    ["見立て",esc(d.get("今後の見立て",""))],["発火元",esc(d.get("発火元ニュース",""))]]}
    svg.append(legend_life(drvX, height-18))
    svg.append(legend_edge(dX, height-18))
    svg.append('</svg>')
    viewA="".join(svg)

    # ---- View B: ライフサイクル俯瞰 ----
    W,Hc=980,430; padL,padR,padT,padB=70,30,40,64
    plotW=W-padL-padR; plotH=Hc-padT-padB
    def px(xv): return padL + (xv-0.6)/(4.6-0.6)*plotW
    def py(att): return padT + (5-att)/(5-0.5)*plotH
    s2=[f'<svg viewBox="0 0 {W} {Hc}" xmlns="http://www.w3.org/2000/svg" role="img" style="width:100%;height:auto;min-width:760px" font-family="-apple-system,\'Hiragino Sans\',\'Noto Sans JP\',sans-serif">']
    s2.append('<title>デザイン・ライフサイクル俯瞰</title>')
    # axis bg bands
    bands=[("台頭",1,"#ecfdf5"),("流行中",2,"#eff6ff"),("成熟",3,"#fff7ed"),("衰退",4,"#f8fafc")]
    for lab,xv,col in bands:
        x0=px(xv-0.5); x1=px(xv+0.5)
        s2.append(f'<rect x="{x0:.0f}" y="{padT}" width="{x1-x0:.0f}" height="{plotH}" fill="{col}" opacity="0.7"/>')
        s2.append(f'<text x="{px(xv):.0f}" y="{padT+plotH+20}" font-size="12" font-weight="700" fill="#64748b" text-anchor="middle">{lab}</text>')
    s2.append(f'<text x="{px(3.4):.0f}" y="{padT+plotH+38}" font-size="11" fill="#7c3aed" text-anchor="middle">← 再評価/復権</text>')
    # y label
    s2.append(f'<text x="18" y="{padT+plotH/2}" font-size="11.5" fill="#64748b" transform="rotate(-90 18 {padT+plotH/2:.0f})" text-anchor="middle">注目度 高 →</text>')
    # points (jitter overlapping by small offset)
    used={}
    for i,d in enumerate(ds):
        g=d["_g"]; c,bg,tc,lab=STATUS_COLORS[g]
        xv=STATUS_X[g]; att=max(1,d["_att"])
        keyxy=(round(xv,1),att); off=used.get(keyxy,0); used[keyxy]=off+1
        X=px(xv)+ (off%2)*0 ; Y=py(att)+ off*16 - (0 if off==0 else 0)
        Y=py(att)+off*15
        name=str(d.get("デザイン名/パターン",""))
        short=name.split("(")[0].split("／")[0]
        short=short if len(short)<=16 else short[:15]+"…"
        key=f"d{i}"
        s2.append(f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')"><circle cx="{X:.0f}" cy="{Y:.0f}" r="7" fill="{c}"/>'
                  f'<text x="{X+11:.0f}" y="{Y+4:.0f}" font-size="11" fill="var(--text)">{esc(short)}</text></g>')
    s2.append(legend_life(padL, Hc-16))
    s2.append('</svg>')
    viewB="".join(s2)

    leadA=('外部の力（生成AI・GaaS逆風・規制・再編・コスト・ユーザー志向）が、どのデザインの型を'
           '<b>押し上げ／押し下げ</b>ているかの因果マップ。ノードをクリックで詳細。')
    leadB='横軸=成熟度（台頭→流行→成熟→衰退、右下は再評価/復権）、縦軸=注目度。<b>今どの型が伸び盛りで、どれが折り返しつつあるか</b>を一望。'
    secC = build_lineage_section(wb, mdata)
    return [section("A","デザイン連動マップ", leadA, viewA),
            section("B","デザイン・ライフサイクル俯瞰", leadB, viewB), secC]

# ============================================================
# 2) 業界全体 ページ
# ============================================================
IND_DRIVERS = [
    ("マクロ・為替/金利", "景気・ドル円・利上げ"),
    ("生成AI・技術", "AI/クラウド/XR"),
    ("規制強化", "課金・年齢・独禁"),
    ("資本・M&A再編", "買収・出資・レイオフ"),
    ("プラットフォーム戦略", "ハード/サブスク/PC"),
]
CAT_COLORS = {
    "tech":("#7c3aed","#f5f3ff","#5b21b6","技術"),
    "reg":("#b91c1c","#fef2f2","#991b1b","規制"),
    "model":("#0891b2","#ecfeff","#155e75","ビジネスモデル"),
    "hw":("#059669","#ecfdf5","#065f46","ハード"),
    "genre":("#d97706","#fff7ed","#9a3412","ジャンル"),
    "pf":("#2563eb","#eff6ff","#1e40af","PF"),
    "ad":("#db2777","#fdf2f8","#9d174d","広告"),
    "capital":("#475569","#f1f5f9","#334155","M&A/資本"),
}
def theme_cat(c):
    c=str(c or "")
    if "M&A" in c or "資本" in c: return "capital"
    if "規制" in c: return "reg"
    if "ビジネス" in c: return "model"
    if "ハード" in c: return "hw"
    if "ジャンル" in c: return "genre"
    if "広告" in c: return "ad"
    if "技術" in c or "B2B" in c: return "tech"
    if "PF" in c: return "pf"
    return "tech"

def theme_drivers(t):
    t=str(t); out=set()
    if any(k in t for k in ["生成AI","クラウド","Web3","デュアルユース","XR","VR"]): out.add(1)
    if "規制" in t or "ガチャ" in t or "Free-to-Play" in t: out.add(2)
    if "サブスク" in t: out.update([4,0])
    if any(k in t for k in ["Switch","Steam","VR/XR"]): out.add(4)
    if "e-Sports" in t or "eスポーツ" in t: out.add(4)
    if "資本再編" in t: out.add(3)
    if "モバイル広告" in t: out.add(0)
    if not out: out.add(4)
    return out

IND_BUCKETS = [
    ("金融・決済", ["金融","決済","フィンテック","保険","KYC","与信"]),
    ("小売・EC", ["小売","EC","百貨店","流通"]),
    ("広告・マーケ", ["広告"]),
    ("動画配信・メディア", ["動画配信","メディア","放送"]),
    ("通信", ["通信","5G","回線"]),
    ("音楽", ["音楽"]),
    ("映画・映像", ["映画","映像","VFX"]),
    ("スポーツ", ["スポーツ","リーグ"]),
    ("自動車", ["自動車","車載","コネクテッド"]),
    ("不動産・MICE", ["不動産","MICE","展示"]),
    ("教育", ["教育"]),
    ("製薬・ヘルスケア", ["製薬","医療","ヘルスケア","バイオ"]),
]
THEME_IND_EDGES = {
    "生成AI": ["広告・マーケ","映画・映像","教育"],
    "クラウド": ["通信","動画配信・メディア"],
    "サブスク": ["動画配信・メディア","音楽"],
    "e-Sports": ["スポーツ","広告・マーケ"],
    "Free-to-Play": ["金融・決済","広告・マーケ"],
    "EU/米国": ["金融・決済","動画配信・メディア"],
    "資本再編": ["金融・決済","動画配信・メディア"],
    "デュアルユース": ["自動車","不動産・MICE"],
    "Steam": ["小売・EC"],
    "モバイル広告": ["広告・マーケ"],
}

def build_industry(wb, mdata):
    themes = rows(wb, "業界トレンド・技術")
    def dkey(d):
        return str(d.get("更新日") or "")
    themes.sort(key=dkey, reverse=True)
    cq = rows(wb, "横断的問い")

    # L3 industry counts
    counts={lab:0 for lab,_ in IND_BUCKETS}
    for r in cq:
        txt=str(r.get("他業界への接続候補") or "")
        for lab,keys in IND_BUCKETS:
            if any(k in txt for k in keys): counts[lab]+=1
    inds=[(lab,counts[lab]) for lab,_ in IND_BUCKETS if counts[lab]>0]
    inds.sort(key=lambda x:-x[1]); inds=inds[:10]
    ind_pos={}

    # layout
    c1x,c1w,c1h=20,178,46
    c2x,c2w,c2h=336,252,34
    c3x,c3w,c3h=710,178,38
    n2=len(themes); top=64; step=40
    height=top+(n2-1)*step+c2h+64
    th_c=[top+i*step+c2h/2 for i in range(n2)]
    def spread(n,cx_top,cx_bot):
        return [cx_top+(cx_bot-cx_top)*i/(max(1,n-1)) for i in range(n)]
    drv_c=spread(len(IND_DRIVERS), th_c[0], th_c[-1])
    ind_c=spread(len(inds), th_c[0], th_c[-1])

    W=c3x+c3w+30
    svg=[f'<svg viewBox="0 0 {W} {height}" xmlns="http://www.w3.org/2000/svg" role="img" style="width:100%;height:auto;min-width:900px" font-family="-apple-system,\'Hiragino Sans\',\'Noto Sans JP\',sans-serif">']
    svg.append('<title>ゲーム業界 全体連動マップ</title>')
    svg.append(f'<text x="{c1x+6}" y="46" font-size="12" font-weight="700" fill="#64748b">外部ドライバー</text>')
    svg.append(f'<text x="{c2x+6}" y="46" font-size="12" font-weight="700" fill="#64748b">業界トレンド・技術（色＝カテゴリ）</text>')
    svg.append(f'<text x="{c3x+6}" y="46" font-size="12" font-weight="700" fill="#64748b">他業界への波及</text>')

    # edges L1->L2
    ep=[]
    for i,d in enumerate(themes):
        for di in theme_drivers(d.get("テーマ")):
            if di<len(drv_c):
                ep.append(f'<path d="{bezier(c1x+c1w, drv_c[di], c2x, th_c[i])}" fill="none" stroke="#94a3b8" stroke-width="1.8" opacity="0.55"/>')
    # edges L2->L3
    ind_y={lab:ind_c[i] for i,(lab,_) in enumerate(inds)}
    for i,d in enumerate(themes):
        tname=str(d.get("テーマ",""))
        for kw,buckets in THEME_IND_EDGES.items():
            if kw in tname:
                for b in buckets:
                    if b in ind_y:
                        ep.append(f'<path d="{bezier(c2x+c2w, th_c[i], c3x, ind_y[b])}" fill="none" stroke="#cbd5e1" stroke-width="1.5" opacity="0.5"/>')
    svg.append('<g>'+''.join(ep)+'</g>')

    # L1 drivers
    for i,(t,s) in enumerate(IND_DRIVERS):
        y=drv_c[i]-c1h/2; key=f"gi{i}"
        svg.append(driver_box(c1x,y,c1w,c1h,t,s,key))
        mdata[key]={"t":esc(t),"s":esc(s),"rows":[["役割","業界全体を動かす外部要因"]]}
    # L2 themes
    for i,d in enumerate(themes):
        g=theme_cat(d.get("カテゴリ")); c,bg,tc,lab=CAT_COLORS[g]
        y=top+i*step; name=str(d.get("テーマ","")); key=f"it{i}"
        short=name if len(name)<=17 else name[:16]+"…"
        svg.append(f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')">'
                   f'<rect x="{c2x}" y="{y}" width="{c2w}" height="{c2h}" rx="8" fill="{bg}" stroke="{c}" stroke-width="1.5"/>'
                   f'<text x="{c2x+11}" y="{y+15}" font-size="11.5" font-weight="700" fill="{tc}">{esc(short)}</text>'
                   f'<text x="{c2x+11}" y="{y+28}" font-size="9.5" fill="{c}">{esc(lab)}</text></g>')
        mdata[key]={"t":esc(name),"s":esc(d.get("現状","")),
            "rows":[["カテゴリ",esc(d.get("カテゴリ",""))],["重要動向",esc(d.get("重要動向",""))],
                    ["今後の予測",esc(d.get("今後の予測",""))],["関連企業",esc(d.get("関連企業",""))],["更新",esc(d.get("更新日",""))]]}
    # L3 industries
    for i,(lab,cnt) in enumerate(inds):
        y=ind_c[i]-c3h/2; key=f"ind{i}"
        r=min(1.0,cnt/12.0)
        svg.append(f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')">'
                   f'<rect x="{c3x}" y="{y}" width="{c3w}" height="{c3h}" rx="8" fill="#fff" stroke="#94a3b8" stroke-width="1.4"/>'
                   f'<text x="{c3x+11}" y="{y+17}" font-size="11.5" font-weight="700" fill="#334155">{esc(lab)}</text>'
                   f'<text x="{c3x+c3w-12}" y="{y+25}" font-size="11" font-weight="800" fill="#0891b2" text-anchor="end">{cnt}件</text></g>')
        mdata[key]={"t":esc(lab),"s":f"「横断的問い」で {cnt} 件の接続候補として言及","rows":[["役割","ゲーム業界の動きが波及・接続しうる隣接業界"]]}
    svg.append('</svg>')
    viewA="".join(svg)

    # ---- View B: 予測マップ ----
    preds=rows(wb,"今後起きそうなこと")
    CAT2={"reg":("#dc2626","規制"),"tech":("#7c3aed","AI・技術"),"capital":("#475569","M&A・資本"),
          "cost":("#d97706","価格・コスト"),"design":("#2563eb","デザイン/GaaS"),"other":("#0891b2","その他")}
    def pcat(t):
        t=str(t)
        if any(k in t for k in ["規制","年齢","ルートボックス","PEGI","確認","法規制"]): return "reg"
        if any(k in t for k in ["AI","生成","Agent"]): return "tech"
        if any(k in t for k in ["M&A","買収","LBO","出資","CVC","スタジオ","非公開化","統合"]): return "capital"
        if any(k in t for k in ["値上げ","メモリ","DRAM","価格","コスト","BNPL","中古"]): return "cost"
        if any(k in t for k in ["GaaS","ライブサービス","買い切り","シングル","物理"]): return "design"
        return "other"
    def tb(t):
        t=str(t)
        if any(k in t for k in ["6〜12","6-12","12ヶ月","半年"]): return 2
        if any(k in t for k in ["1〜2年","1-2年","2027","18ヶ月","次世代","次の","1年"]): return 3
        if any(k in t for k in ["数ヶ月","3〜6","3-6","年内","9月","7月"]): return 1
        return 2
    def cb(t):
        t=str(t)
        return 3 if "高" in t else (1 if "低" in t else 2)

    W2,H2=980,470; padL,padR,padT,padB=60,140,44,60
    plotW=W2-padL-padR; plotH=H2-padT-padB
    def qx(b): return padL + (b-1)/2.0*plotW
    def qy(c): return padT + (3-c)/2.0*plotH
    s2=[f'<svg viewBox="0 0 {W2} {H2}" xmlns="http://www.w3.org/2000/svg" role="img" style="width:100%;height:auto;min-width:820px" font-family="-apple-system,\'Hiragino Sans\',\'Noto Sans JP\',sans-serif">']
    s2.append('<title>今後の予測マップ</title>')
    for b,lab in [(1,"近い（〜半年/年内）"),(2,"中（6〜12ヶ月）"),(3,"遠い（1年〜/2027）")]:
        s2.append(f'<line x1="{qx(b):.0f}" y1="{padT}" x2="{qx(b):.0f}" y2="{padT+plotH}" stroke="var(--border)" stroke-width="1" opacity="0.5"/>')
        s2.append(f'<text x="{qx(b):.0f}" y="{padT+plotH+22}" font-size="11.5" font-weight="700" fill="#64748b" text-anchor="middle">{lab}</text>')
    for c,lab in [(3,"確度 高"),(2,"中"),(1,"低")]:
        s2.append(f'<text x="{padL-12}" y="{qy(c)+4:.0f}" font-size="11" fill="#64748b" text-anchor="end">{lab}</text>')
    s2.append(f'<text x="{padL+plotW/2:.0f}" y="{padT+plotH+44}" font-size="11.5" fill="#64748b" text-anchor="middle">← 想定時期 →</text>')
    used={}
    for i,d in enumerate(preds):
        cat=pcat(d.get("今後起きそうなこと")); col=CAT2[cat][0]
        b=tb(d.get("想定時期")); cf=cb(d.get("確度"))
        cell=(b,cf); off=used.get(cell,0); used[cell]=off+1
        X=qx(b)-40+ (off%3)*30; Y=qy(cf)-24 + (off//3)*16 + (off%3)*8
        key=f"p{i}"
        s2.append(f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')"><circle cx="{X:.0f}" cy="{Y:.0f}" r="6.5" fill="{col}"/></g>')
        mdata[key]={"t":esc(str(d.get("今後起きそうなこと",""))[:80]),"s":"",
            "rows":[["確度",esc(d.get("確度",""))],["想定時期",esc(d.get("想定時期",""))],
                    ["関連",esc(d.get("関連企業/IP",""))],["根拠",esc(d.get("根拠となるニュース",""))],
                    ["検証",esc(d.get("検証メモ",""))]]}
    # legend
    lx=W2-128; ly=padT+6
    s2.append(f'<text x="{lx}" y="{ly}" font-size="11.5" font-weight="700" fill="#475569">カテゴリ</text>')
    for j,(k,(col,lab)) in enumerate(CAT2.items()):
        yy=ly+20+j*20
        s2.append(f'<circle cx="{lx+6}" cy="{yy-4}" r="6" fill="{col}"/><text x="{lx+18}" y="{yy}" font-size="11" fill="#475569">{esc(lab)}</text>')
    s2.append('</svg>')
    viewB="".join(s2)

    leadA=('外部ドライバー→<b>ゲーム業界のトレンド/技術</b>→他業界への波及、という三層で全体の連動を俯瞰。'
           '右端の件数は「横断的問い」での接続言及の多さ。ノードをクリックで詳細。')
    leadB=('「今後起きそうなこと」を<b>想定時期（横）×確度（縦）</b>で配置した予測マップ。'
           '右上ほど“確度が高く近い”＝今すぐ追うべき論点。色はテーマ。')
    return [section("A","ゲーム業界 全体連動マップ", leadA, viewA),
            section("B","今後の予測マップ", leadB, viewB)]


# ============================================================
# 3) デザイン派生系統（時系列・統合系統樹）
# ============================================================
LINEAGE_LANES = ["対話型AI系","ローグ系","抽出/PvPvE系","MMO→GaaS系",
                 "買い切りエバーグリーン系","シネマティック物語系","サバイバル×収集系"]
LANE_ACCENT = {
    "対話型AI系":"#7c3aed","ローグ系":"#0891b2","抽出/PvPvE系":"#dc2626",
    "MMO→GaaS系":"#d97706","買い切りエバーグリーン系":"#059669",
    "シネマティック物語系":"#2563eb","サバイバル×収集系":"#db2777",
}
KIND_FILL = {"祖先":"#cbd5e1","発展":"#94a3b8","確立":"#64748b"}

def build_lineage_section(wb, mdata):
    ln = rows(wb, "デザイン派生系統")
    # 現行型の状態色（ゲームデザイン・トレンドから引く）
    ds = rows(wb, "ゲームデザイン・トレンド")
    def pattern_color(pat):
        for d in ds:
            nm=str(d.get("デザイン名/パターン",""))
            if pat and (pat in nm or nm in pat):
                return STATUS_COLORS[status_group(d.get("ステータス"))]
        return STATUS_COLORS["hot"]
    if not ln:
        return section("C","デザイン派生系統（時系列）","データ準備中。","<svg viewBox=\"0 0 100 20\"></svg>")

    # index nodes
    nodes={}
    for r in ln:
        nm=str(r.get("ノード(タイトル/技術)","")).strip()
        try: yr=int(r.get("年"))
        except: continue
        nodes[nm]={"lane":str(r.get("系統名","")),"year":yr,"kind":str(r.get("種別","")),
                   "parents":[x.strip() for x in str(r.get("派生元") or "").replace("、",",").split(",") if x.strip()],
                   "pattern":str(r.get("現行型") or ""),"memo":str(r.get("メモ") or "")}
    years=[n["year"] for n in nodes.values()]
    minY=min(years); maxY=2027
    lanes=[l for l in LINEAGE_LANES if any(n["lane"]==l for n in nodes.values())]

    W=1480; padL=132; padR=44; padT=58; padB=70
    bandH=(848-padT-padB)/len(lanes); H=int(padT+padB+bandH*len(lanes))
    def xof(y): return padL+(y-minY)/(maxY-minY)*(W-padL-padR)
    laneY={l:padT+(i+0.5)*bandH for i,l in enumerate(lanes)}
    # stagger within lane by year order
    yoff={}
    for l in lanes:
        seq=sorted([nm for nm,n in nodes.items() if n["lane"]==l], key=lambda nm:nodes[nm]["year"])
        for j,nm in enumerate(seq):
            yoff[nm]=(-20 if j%2==0 else 18)
    def pos(nm):
        n=nodes[nm]; return xof(n["year"]), laneY[n["lane"]]+yoff.get(nm,0)

    svg=[f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg" role="img" '
         f'style="width:100%;height:auto;min-width:1200px" font-family="-apple-system,\'Hiragino Sans\',\'Noto Sans JP\',sans-serif">']
    svg.append('<title>デザイン派生系統（統合系統樹）</title>')
    # year gridlines
    for yr in [1980,1990,2000,2010,2020,2026]:
        if yr<minY: continue
        x=xof(yr)
        svg.append(f'<line x1="{x:.0f}" y1="{padT-6}" x2="{x:.0f}" y2="{H-padB+6}" stroke="#e2e8f0" stroke-width="1" opacity="0.5"/>')
        svg.append(f'<text x="{x:.0f}" y="{padT-14}" font-size="11" fill="#94a3b8" text-anchor="middle">{yr}</text>')
        svg.append(f'<text x="{x:.0f}" y="{H-padB+22}" font-size="11" fill="#94a3b8" text-anchor="middle">{yr}</text>')
    # lane labels + separators
    for i,l in enumerate(lanes):
        yb=padT+i*bandH
        svg.append(f'<line x1="{padL-8}" y1="{yb:.0f}" x2="{W-padR}" y2="{yb:.0f}" stroke="#eef2f7" stroke-width="1"/>')
        svg.append(f'<text x="10" y="{laneY[l]+4:.0f}" font-size="11.5" font-weight="700" fill="{LANE_ACCENT.get(l,"#475569")}">{esc(l)}</text>')
    # edges
    ep=[]
    for nm,n in nodes.items():
        x2,y2=pos(nm)
        for pnm in n["parents"]:
            if pnm in nodes:
                x1,y1=pos(pnm)
                cross = nodes[pnm]["lane"]!=n["lane"]
                col = "#f97316" if cross else "#cbd5e1"
                dash=' stroke-dasharray="5 4"' if cross else ""
                ep.append(f'<path d="{bezier(x1,y1,x2,y2)}" fill="none" stroke="{col}" stroke-width="{2.2 if cross else 1.6}"{dash} opacity="0.8"/>')
    svg.append('<g>'+''.join(ep)+'</g>')
    # nodes
    idx=0
    for nm,n in sorted(nodes.items(), key=lambda kv:kv[1]["year"]):
        x,y=pos(nm); key=f"ln{idx}"; idx+=1
        mdata[key]={"t":esc(nm),"s":esc(n["memo"]),
            "rows":[["年",str(n["year"])],["系統",esc(n["lane"])],["種別",esc(n["kind"])],
                    ["派生元",esc("、".join(n["parents"]) or "—")]+([] if False else []),
                    ["現行型",esc(n["pattern"]) if n["pattern"] else "—"]]}
        if n["kind"].startswith("現在"):
            c,bg,tc,lab=pattern_color(n["pattern"])
            label=nm if len(nm)<=20 else nm[:19]+"…"
            bw=min(len(label)*12+22, 260); bx=min(x, W-bw-8); by=y-15
            svg.append(f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')">'
                       f'<rect x="{bx:.0f}" y="{by:.0f}" width="{bw}" height="30" rx="8" fill="{bg}" stroke="{c}" stroke-width="2"/>'
                       f'<text x="{bx+10:.0f}" y="{by+19:.0f}" font-size="11" font-weight="800" fill="{tc}">{esc(label)}</text></g>')
        else:
            fill=KIND_FILL.get(n["kind"],"#94a3b8")
            label=nm if len(nm)<=22 else nm[:21]+"…"
            # place label right if room else left
            right = x < W-260
            tx = x+11 if right else x-11
            anc = "start" if right else "end"
            svg.append(f'<g class="mnode" style="cursor:pointer" onclick="mShow(\'{key}\')">'
                       f'<circle cx="{x:.0f}" cy="{y:.0f}" r="6" fill="{fill}" stroke="#fff" stroke-width="1.2"/>'
                       f'<text x="{tx:.0f}" y="{y+4:.0f}" font-size="10.5" fill="var(--text)" text-anchor="{anc}">{esc(label)}</text></g>')
    # legend
    ly=H-22
    svg.append(f'<text x="{padL}" y="{ly}" font-size="11.5" font-weight="700" fill="#475569">種別:</text>')
    xx=padL+42
    for k in ["祖先","発展","確立"]:
        svg.append(f'<circle cx="{xx}" cy="{ly-4}" r="6" fill="{KIND_FILL[k]}"/><text x="{xx+12}" y="{ly}" font-size="11.5" fill="#475569">{k}</text>')
        xx+=len(k)*13+34
    svg.append(f'<rect x="{xx}" y="{ly-11}" width="14" height="12" rx="3" fill="#eff6ff" stroke="#2563eb"/><text x="{xx+18}" y="{ly}" font-size="11.5" fill="#475569">現在の型</text>')
    xx+=120
    svg.append(f'<line x1="{xx}" y1="{ly-4}" x2="{xx+22}" y2="{ly-4}" stroke="#cbd5e1" stroke-width="3"/><text x="{xx+28}" y="{ly}" font-size="11.5" fill="#475569">同系統</text>')
    xx+=110
    svg.append(f'<line x1="{xx}" y1="{ly-4}" x2="{xx+22}" y2="{ly-4}" stroke="#f97316" stroke-width="3" stroke-dasharray="5 4"/><text x="{xx+28}" y="{ly}" font-size="11.5" fill="#475569">他系統からの影響</text>')
    svg.append('</svg>')

    lead=('主要な型が<b>いつ・何から派生したか</b>を1本の時間軸に統合した系統樹。丸=節目タイトル、'
          '囲み=現在の型、<span style="color:#f97316">橙の点線</span>は他系統からの影響（例: GaaS逆風→買い切りエバーグリーン/シングル大作復権）。ノードをクリックで詳細。')
    return section("C","デザイン派生系統（時系列・統合系統樹）", lead, "".join(svg))

# ============================================================
def generate():
    import json
    wb = load_workbook(XLSX, data_only=True)
    # design page
    md1={}
    secs1 = build_design(wb, md1)
    (DOCS/"design-map.html").write_text(
        page("ゲームデザイン連動マップ","デザインの型を「外部要因×ライフサイクル」で構造化",
             "design-map", secs1, json.dumps(md1, ensure_ascii=False)), encoding="utf-8")
    # industry page
    md2={}
    secs2 = build_industry(wb, md2)
    (DOCS/"industry-map.html").write_text(
        page("ゲーム業界 全体連動マップ","マクロ→業界→他業界の波及と、今後の予測マップ",
             "industry-map", secs2, json.dumps(md2, ensure_ascii=False)), encoding="utf-8")
    print("maps: design-map.html / industry-map.html generated")

if __name__ == "__main__":
    generate()
