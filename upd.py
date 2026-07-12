# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, Alignment
from copy import copy
wb = openpyxl.load_workbook("master/ゲーム業界・時事情報収集.xlsx")
TODAY = "2026-07-12"
YG = Font(name="Yu Gothic")
AL = Alignment(wrap_text=True, vertical="top")

def style(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c); cell.font = copy(YG); cell.alignment = copy(AL)

wn = wb["日次ニュース"]
ex = {str(wn.cell(r, 3).value or "").strip() for r in range(2, wn.max_row + 1)}
news = [
 [TODAY,
  "中国最大級のサブカル展「bilibili World 2026」が上海で開幕（7/10〜12）——Black Myth陣営やTotal War新作が出展",
  "https://www.invenglobal.com/articles/23379/bilibili-world-2026-a-subculture-hub-featuring-kevin-feige",
  "E3消滅後、アジアではbilibili WorldがサブカルとゲームのハブとしてTGS・gamescomに並ぶ規模に成長。中国のゲーム消費と海外パブリッシャーの中国市場攻略の重要接点になっている。",
  "上海の国家会展中心で7/10〜12開催。170超の出展者・700超のブースに世界130社のゲーム開発者が集結。原神・ブルアカ・鳴潮・アークナイツ：エンドフィールド等が出展し、Game Science（黒神話）やCreative AssemblyのTotal War新作もアピールした。",
  "・7/10〜12・上海、世界130社が参加する中国最大級のサブカル/ゲーム展\n・Game Science（黒神話：鍾馗の続報に注目）が『Black Myth』名義で出展\n・Creative Assembly『Total War: WARHAMMER 40,000』が新エンジンWarCoreで初のオフライン出展\n・マーベル・スタジオのKevin Feige氏が登壇\n・『Love and Deepspace』は炎上を受け出展を辞退",
  "中国国内の版号回復と大型IPの海外展開を背景に、bilibili Worldは中国市場攻略の実質的な商談・PRの場として比重が高まる。海外AAAの中国オフライン出展が定着すれば、gamescom/TGSと並ぶ第3極の見本市化が進む。",
  "中国ゲーム市場と見本市エコシステム",
  "bilibili, Game Science, Creative Assembly, SEGA",
  "黒神話, Total War, 原神, ブルーアーカイブ",
  "海外,新作,eスポーツ", ""],
 [TODAY,
  "eスポーツ・ワールドカップ2026、VALORANT部門が7/12にグランドファイナル——賞金総額200万ドル、パリで開催",
  "https://dotesports.com/valorant/guides/ewc-2026-valorant",
  "EWCは中東（リヤド）で2回開催後、2026年は情勢を受けパリに会場を移転。7/6〜8/23の長期会期で24競技25大会を実施する世界最大級のeスポーツ祭典で、クラブの複数タイトル横断ポイント制が特徴。",
  "パリのParis Expo Porte de Versaillesで開催中のEWC2026で、VALORANT部門（16チーム・賞金200万ドル）が7/2〜12の日程を終え、7/12に3位決定戦とグランドファイナルを実施。EWC全体は8/23まで継続する。",
  "・EWC2026はパリ開催、会期7/6〜8/23の過去最大規模\n・VALORANT部門の決勝が7/12、賞金プールは200万ドル\n・今年はTrackmaniaが新規参入、Fortniteがリロードモードで復帰\n・クラブ横断のポイントで年間王者を決定",
  "パリ移転で欧州の観客動員と放映がどこまで伸びるかが今後の開催地戦略を左右する。会期の長期化・多タイトル化により、単発大会からクラブ経営（横断ポイント×スポンサー）モデルへの移行がさらに進む。",
  "eスポーツの興行・クラブ経営モデル",
  "Esports World Cup Foundation, Riot Games",
  "VALORANT, Trackmania, Fortnite",
  "eスポーツ,海外", ""],
 [TODAY,
  "GameSpotで再びレイオフ——親会社FandomがコマースチームとゲームVPを整理、購入ガイド部門が消滅",
  "https://aftermath.site/gamespot-layoffs-july-2026/",
  "2023年にFandom傘下入りしたGameSpotは断続的な人員削減が続く。ゲームメディアはレビュー広告からアフィリエイト/コマース（セール・予約ガイド）収益へ依存を強めてきたが、その中核部門にメスが入った。",
  "Fandomが木曜にGameSpotのコマースチーム（正社員4名＋フリー5名）とゲーム・エンタメ担当VPのChris Grant氏を整理。セール記事や予約・購入ガイドを担っていた部門が実質消滅した。2026年に入って2度目の削減となる。",
  "・Fandomがコマースチーム全体とゲーム担当VPを削減\n・対象は購入/予約ガイド・セール記事の中核部門\n・2026年に入り2度目、Fandom買収後は繰り返しの削減\n・ゲームメディアの購入導線マネタイズモデルへの逆風を象徴",
  "プラットフォーム（Steam/ストア）が購入導線とキュレーションを内製化するなか、第三者メディアの送客マージンは縮小。メディアはB2B（データ・コンサル）や直接課金（会員制）への再編を迫られる。",
  "ゲームメディアの収益構造とアフィリエイト依存",
  "Fandom, GameSpot", "",
  "海外", ""],
 [TODAY,
  "EU AI Actの透明性義務が8/2適用開始へ——AI生成コンテンツの識別・表示がゲーム開発にも波及",
  "https://digital-strategy.ec.europa.eu/en/policies/regulatory-framework-ai",
  "EU AI Actは段階施行中で、生成AIの透明性ルールが2026年8月に本格適用。ディープフェイクやAI生成物の識別可能性・表示を求める。ゲームは最小リスク区分だが、Steam/Apple/GoogleがAI利用開示ポリシーを整備しつつある。",
  "EU AI Actの透明性条項が2026年8月2日から適用開始。AI生成コンテンツは機械可読な識別を付し、ディープフェイク等は明示的にラベリングする必要がある。ストアプラットフォーム側のAIアセット開示要求とも連動し、開発工程の記録・表示対応が求められる。",
  "・生成AIの透明性ルールが8/2から適用開始\n・AI生成物の識別可能性とディープフェイクの明示ラベルが必要\n・ゲーム自体は最小リスク区分だが表示・開示の実務対応は発生\n・Steam/Apple/GoogleもAIアセット開示ポリシーを整備中",
  "EUでの表示義務が事実上のグローバル標準化圧力となり、ストアページでのAI利用開示が一般化。開発側は生成物のトレーサビリティ管理が新たな運用コストになる一方、非AI訴求の差別化余地も広がる。",
  "AI規制（EU AI Act）とAI生成物の開示実務",
  "European Commission, Valve, Apple, Google", "",
  "規制,技術,海外", ""],
 [TODAY,
  "東映ゲームズ、台湾最大級の開発者会議「TGDF 2026」（7/15〜16）に登壇——日台のゲーム市場連携を議論",
  "https://www.4gamer.net/games/991/G999104/20260708042/",
  "東映は映像資産を核にゲーム事業を強化中で、アジア市場、とりわけ成長する台湾のインディー/開発者コミュニティとの連携を模索。TGDFは台湾最大級の開発者カンファレンスで、アジア各国の開発者が集う。",
  "東映ゲームズが7/15〜16に台北で開かれる「Taipei Game Developers Forum 2026」に登壇。7/16に新規事業開発マネージャーの長島寛晃氏が事業戦略セッションを行い、日本と台湾のゲーム市場の関係や協業の展望を語る。",
  "・東映ゲームズがTGDF 2026（7/15〜16・台北）に登壇\n・7/16に事業戦略セッション（日本語・中文字幕）\n・テーマは日台ゲーム市場の連携と東映の事業戦略\n・映像IPを持つ東映のゲーム分野アジア展開の一環",
  "日本の映像・IPホルダーが台湾/東南アジアの開発力と組む座組が増え、受託・共同開発やIPライセンスの日台クロスボーダー案件が拡大する可能性。",
  "日台クロスボーダーのゲーム協業とIP活用",
  "東映, 東映ゲームズ", "",
  "国内,海外", ""],
]
added_news = 0
for row in news:
    if row[2].strip() in ex:
        print("skip dup news:", row[1][:20]); continue
    wn.append(row); style(wn, wn.max_row, 12); added_news += 1
print("日次ニュース追加:", added_news)

wq = wb["横断的問い"]
exq = {str(wq.cell(r, 2).value or "")[:24] for r in range(2, wq.max_row + 1)}
ques = [
 [TODAY,
  "AI生成コンテンツのラベリング義務化（EU AI Act 8月施行）は、ゲームだけでなく広告・報道・EC商品画像など『AI生成物の開示』を他業界にどこまで波及させるか？",
  "EU AI Actの透明性義務が8/2適用開始（EU公式）",
  "European Commission, Valve",
  "広告（AIクリエイティブ表示）/ EC（AI商品画像の開示）/ 報道（AI生成記事の明示）",
  "オープン", "表示義務は横断的規制になりやすく、ゲームは先行実装される表示UXの実験場になりうる。"],
 [TODAY,
  "ゲームメディアのアフィリエイト/コマース（購入導線）依存モデルの解体（GameSpot）は、家電・金融・旅行など比較サイトを持つ他業界メディアでも同じ収益崩壊を招くか？",
  "GameSpotで再度のレイオフ、Fandomがコマースチームを整理（Aftermath）",
  "Fandom, GameSpot",
  "メディア全般（家電・金融・旅行の比較/レビュー）/ EC（送客マージン）/ アフィリエイト",
  "オープン", "プラットフォームが購入導線を内製化する構図はゲーム外でも進行中で、第三者メディアの中抜き余地が縮む共通課題。"],
]
added_q = 0
for row in ques:
    if row[1][:24] in exq:
        print("skip dup q"); continue
    wq.append(row); style(wq, wq.max_row, 7); added_q += 1
print("横断的問い追加:", added_q)

we = wb["イベントカレンダー"]
exe = {str(we.cell(r, 3).value or "").strip() for r in range(2, we.max_row + 1)}
events = [
 ["2026-07-10", "2026-07-12", "bilibili World 2026", "展示会", "上海（国家会展中心）", "オフライン",
  "中国最大級のサブカル/ゲーム展。世界130社が出展、Black Myth陣営やTotal War新作、Kevin Feige登壇。",
  "https://www.invenglobal.com/articles/23379/bilibili-world-2026-a-subculture-hub-featuring-kevin-feige", "開催中", TODAY],
 ["2026-07-15", "2026-07-16", "Taipei Game Developers Forum 2026 (TGDF)", "展示会", "台北（Taipei New Horizon）", "オフライン",
  "台湾最大級のゲーム開発者カンファレンス。アジア各国の開発者が集う。東映ゲームズ等日本勢も登壇。",
  "https://www.4gamer.net/games/991/G999104/20260708042/", "開催予定", TODAY],
]
added_e = 0
for row in events:
    if row[2].strip() in exe:
        print("skip dup event:", row[2]); continue
    we.append(row); style(we, we.max_row, 10); added_e += 1
print("イベント追加:", added_e)

wf = wb["今後起きそうなこと"]
wf.append([TODAY,
  "AI生成アセットの開示義務化（EU AI Act）を逆手に、『人の手/非AI保証』が差別化ラベルとして商品価値化。Steam等で制作にAI不使用のストア表記やフィルタ需要が生まれ、インディーのマーケ訴求軸として定着する。",
  "中", "半年〜年内", "Valve, インディー各社",
  "EU AI Actの透明性義務8/2適用開始（本日追記の日次ニュース・EU公式）",
  "Steamのストアページで『No Gen-AI』表記や絞り込み/バッジが増えるか、インディーが非AIを訴求軸に据える事例が増えるかを観測。"])
style(wf, wf.max_row, 7)
print("今後起きそうなこと追加:1")

wd = wb["ゲームデザイン・トレンド"]
wd.cell(7, 1, TODAY)
wd.cell(7, 7, "生成AIコンテンツの氾濫とユーザーの真正性選好。プラットフォームのAI開示ルール強化に加え、EU AI Actの透明性義務(8/2適用)でAI生成表示が制度化。")
wd.cell(7, 8, "真正性が効率に勝つニッチが拡大。表示義務化で非AIがラベル/フィルタとして商品価値化し、AI活用と非AI訴求の二極化が加速。")
wd.cell(7, 9, "EU AI Act透明性義務が8/2適用開始(本日追記の日次ニュース)")
for c in range(1, 11):
    cell = wd.cell(7, c); cell.font = copy(YG); cell.alignment = copy(AL)
print("ゲームデザイン更新:row7")

wb.save("master/ゲーム業界・時事情報収集.xlsx")
print("SAVED. 日次ニュース max_row:", wb["日次ニュース"].max_row)
