# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, Alignment
PATH='master/ゲーム業界・時事情報収集.xlsx'
TODAY='2026-08-26'
wb=openpyxl.load_workbook(PATH)
FONT=Font(name='Yu Gothic'); ALIGN=Alignment(wrap_text=True, vertical='top')
def style_row(ws,r):
    for c in ws[r]:
        c.font=FONT; c.alignment=ALIGN
NL="\n"
def P(*x): return NL.join(x)

ws=wb['日次ニュース']
ex=set(str(r[3]).strip() for r in ws.iter_rows(min_row=2, values_only=True) if r[3])
rows=[
 ["映画","ヴェネツィア映画祭2026ラインナップ発表、ダニー・ボイル『Ink』が開幕作",
 "https://www.screendaily.com/news/venice-film-festival-reveals-2026-lineup/5218820.article",
 "ヴェネツィア国際映画祭（第82回、9/2〜12）は秋の賞レースの号砲。コンペ入り作品はアカデミー賞シーズンの前哨戦として世界の配給・配信の注目を集める。",
 "9/2開幕作はダニー・ボイル監督『Ink』（Jack O'Connell、Guy Pearce、Claire Foy）でコンペ上映。マーティン・マクドナー、フロリアン・ゼレール、ヴェルナー・ヘルツォーク、イ・チャンドンら巨匠作が並び、審査委員長はマギー・ジレンホール。",
 P("・9/2開幕作はボイル『Ink』でコンペ入り","・巨匠監督作が多数コンペに集結","・秋の賞レースと配信獲得競争の起点"),
 "受賞・話題作は劇場配給と配信プラットフォームの獲得合戦に直結。Netflix/Amazon等の配信勢と劇場配給の綱引きが続き、賞レースを通じた作品の“ブランド化”が価値を左右する。",
 "映画祭とアワードシーズンの配給戦略",
 "Netflix, Amazon MGM","Ink","受賞,海外,トレンド","映画祭"],
 ["映画","中国の2026年興行が35.7億ドルで前年比31.8%減、Hollywood大作が下支え",
 "https://www.hollywoodreporter.com/movies/movie-news/china-box-office-spider-man-christopher-nolan-the-odyssey-1236669960/",
 "世界2位の映画市場・中国は自国作の不振と消費減速で低迷が続く。北米が回復する一方、中国の弱さがグローバル興行の重石となっている。",
 "THRによると中国の2026年累計興行は35.7億ドルで前年同期比31.8%減。週末は約9730万ドルにとどまる中、『スパイダーマン』が首位を維持し、ノーラン『The Odyssey』のIMAX先行が6.9百万ドルを記録するなどHollywood大作が下支えした。",
 P("・中国2026年興行は前年比31.8%減と大幅減","・自国作不振をHollywood大作が補う","・北米回復と中国低迷の二極化"),
 "中国市場の弱含みが続けば、Hollywoodの世界戦略は中国依存を下げ北米・その他地域へ再配分。中国製大作の復調が鍵で、当局の公開許可や消費回復が変数となる。",
 "中国映画市場の低迷とグローバル再配分",
 "Sony Pictures","Spider-Man: Brand New Day, The Odyssey","興行,海外,トレンド","国際興行"],
]
added=0
for r in rows:
    if r[2].strip() in ex: 
        print("skip dup", r[1][:20]); continue
    ws.append(r); style_row(ws, ws.max_row); ex.add(r[2].strip()); added+=1
print("movie added", added)

# Venice event
wse=wb['イベントカレンダー']
names=set(str(x[2].value).strip() for x in wse.iter_rows(min_row=2) if x[2].value)
vname="第82回 ヴェネツィア国際映画祭 2026"
if not any("ヴェネツィア" in n for n in names):
    wse.append(["2026-09-02","2026-09-12",vname,"映画祭・賞","ヴェネツィア(伊)","リアル",
      "秋の賞レース号砲。開幕作はダニー・ボイル『Ink』、巨匠作がコンペに集結","https://www.screendaily.com/news/venice-film-festival-reveals-2026-lineup/5218820.article","開催予定",TODAY])
    style_row(wse, wse.max_row); print("venice event added")
else:
    print("venice already exists")

wb.save(PATH)
print("SAVED")
