#!/usr/bin/env python3
"""
xlsx_to_mdhtml.py を エンタメ6ジャンル対応に拡張するパッチ（冪等）
リポジトリルートで実行: python3 patch_converter.py
既存の機能（タグチップ・フィルタUI・デザイン連動マップ等）は保持する
"""
import re
import sys
from pathlib import Path

TARGET = Path(__file__).resolve().parent / "scripts" / "xlsx_to_mdhtml.py"

NEW_SLUGS_BLOCK = '''    "デザイン派生系統": "16-design-lineage",
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
}'''

GROUPS_BLOCK = '''

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
'''

NEW_NAV = '''def nav_html(active_slug=None):
    parts = ['<nav class="grouped">']
    parts.append('<div class="navrow"><span class="navlabel">TOP</span>'
                 '<a href="index.html">ダッシュボード</a>'
                 '<a href="design-map.html">デザイン連動マップ</a>'
                 '<a href="industry-map.html">全体連動マップ</a></div>')
    for gname, sheets in SHEET_GROUPS:
        links = []
        for sn in sheets:
            slug = SHEET_SLUGS.get(sn)
            if not slug:
                continue
            cls = ' class="active"' if slug == active_slug else ''
            links.append(f'<a href="{slug}.html"{cls}>{html.escape(_short(sn))}</a>')
        if links:
            parts.append(f'<div class="navrow"><span class="navlabel">{html.escape(gname)}</span>'
                         + ''.join(links) + '</div>')
    parts.append('</nav>')
    return ''.join(parts)
'''

NAV_CSS = '''
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
.genre-h{margin:22px 0 6px; font-size:13px; color:var(--muted); font-weight:700;
  letter-spacing:.04em; border-bottom:1px solid var(--border); padding-bottom:5px}
</style>'''

NEW_INDEX_BODY = '''def render_index_html(stats):
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
'''


def main():
    if not TARGET.exists():
        print(f"ERROR: not found {TARGET}", file=sys.stderr)
        sys.exit(1)
    src = TARGET.read_text(encoding="utf-8")
    orig = src
    changes = []

    # 1) SHEET_SLUGS を拡張
    if '"メディアミックス追跡"' not in src:
        m = re.search(r'(SHEET_SLUGS\s*=\s*\{.*?)\n\}', src, re.S)
        if not m:
            print("ERROR: SHEET_SLUGS が見つかりません", file=sys.stderr)
            sys.exit(1)
        body = m.group(1)
        add = NEW_SLUGS_BLOCK
        if '"デザイン派生系統"' in body:
            add = "\n".join(l for l in NEW_SLUGS_BLOCK.split("\n")
                            if '"デザイン派生系統"' not in l)
        src = src[:m.start()] + body + "\n" + add + src[m.end():]
        changes.append("SHEET_SLUGS を拡張")
    else:
        changes.append("SHEET_SLUGS は拡張済 (skip)")

    # 2) SHEET_GROUPS を追加
    if "SHEET_GROUPS" not in src:
        idx = src.find("SECTION_SHEETS")
        src = src[:idx] + GROUPS_BLOCK.strip() + "\n\n" + src[idx:]
        changes.append("SHEET_GROUPS / _short を追加")
    else:
        changes.append("SHEET_GROUPS は追加済 (skip)")

    # 3) nav_html を差し替え
    if 'nav class="grouped"' not in src:
        m = re.search(r'def nav_html\(active_slug=None\):.*?\n(?=def )', src, re.S)
        if m:
            src = src[:m.start()] + NEW_NAV + "\n" + src[m.end():]
            changes.append("nav_html をグループ化版に差し替え")
        else:
            print("WARN: nav_html を差し替えられませんでした")
    else:
        changes.append("nav_html は差し替え済 (skip)")

    # 4) nav 用CSS を追加
    if "nav.grouped" not in src:
        src = src.replace("</style>", NAV_CSS, 1)
        changes.append("ナビ用CSSを追加")
    else:
        changes.append("ナビ用CSSは追加済 (skip)")

    # 5) render_index_html を差し替え
    if "エンタメ業界・時事情報収集" not in src:
        m = re.search(r'def render_index_html\(stats\):.*?\n(?=def )', src, re.S)
        if m:
            src = src[:m.start()] + NEW_INDEX_BODY + "\n" + src[m.end():]
            changes.append("render_index_html をジャンル別表示に差し替え")
        else:
            print("WARN: render_index_html を差し替えられませんでした")
    else:
        changes.append("render_index_html は差し替え済 (skip)")

    # 6) README.md 生成タイトルも変更
    src = src.replace('"# 🎮 ゲーム業界・時事情報収集"', '"# エンタメ業界・時事情報収集"')

    if src != orig:
        TARGET.write_text(src, encoding="utf-8")

    print("パッチ結果:")
    for c in changes:
        print(f"  - {c}")
    print("\n構文チェック:")
    import py_compile
    py_compile.compile(str(TARGET), doraise=True)
    print("  OK")


if __name__ == "__main__":
    main()
