# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, Alignment
from datetime import date, datetime

PATH='master/ゲーム業界・時事情報収集.xlsx'
TODAY='2026-08-26'
wb=openpyxl.load_workbook(PATH)

FONT=Font(name='Yu Gothic')
ALIGN=Alignment(wrap_text=True, vertical='top')
def style_row(ws, r):
    for c in ws[r]:
        c.font=FONT; c.alignment=ALIGN

# ---------- 日次ニュース ----------
ws=wb['日次ニュース']
existing_urls=set()
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[3]: existing_urls.add(str(row[3]).strip())

NL="\n"
def P(*items): return NL.join(items)

rows=[]
# ===== GAME =====
rows.append(["ゲーム","gamescom Opening Night Live 2026、2時間で新作・復刻を一挙公開",
"https://www.gamespot.com/articles/gamescom-opening-night-live-2026-all-the-biggest-announcements-games-and-trailers/",
"欧州最大の games見本市gamescom（8/26〜30、ケルン）の前夜祭として、Geoff Keighley司会のOpening Night Liveが8/25に開催。新作AAAの開発長期化で発表点数が読みにくい中、復刻・リマスターと運営型タイトルの比重が高まっている。",
"約2時間の配信で、『ウィッチャー3』リマスター、Heroes of Might and Magic III リメイク、Gears of War: E-Day新映像、HoYoverseの新作などを一挙公開。新規AAAより既存IPの再活用・拡張の露出が目立つ構成となった。",
P("・復刻/リマスターと運営型の露出が新規AAAを上回る","・欧州見本市がショーケースの主戦場として定着","・大手からインディーまで“既存資産の再収益化”が共通線"),
"開発費高騰で新規AAAが減る中、各社は既存IPの再パッケージへ資源配分をシフト。今後もリマスター・拡張・アニメ化など“IP延命”の発表が主流化し、完全新規の希少性が上がる見込み。",
"AAA開発費インフレとIP再活用戦略",
"Geoff Keighley, Xbox, CD Projekt, HoYoverse","gamescom, Gears of War","新作,リメイク・移植,海外,トレンド","ONL2026総括"])

rows.append(["ゲーム","CD Projekt、『ウィッチャー3』リマスターを9/29に無料配布・Switch2対応",
"https://press.cdprojektred.com/en/news/1839/the-witcher-3-wild-hunt-remastered-announced-songs-of-the-past-gets-first-look",
"『ウィッチャー3』は2015年発売のロングセラーで、続編・新規IP開発が長期化するCD Projektにとって旗艦資産。旧作の再エンゲージが新作『The Witcher 4』前の需要維持に直結する局面にある。",
"9/29にPC/PS5/Xbox向けリマスターを既存所有者へ無料アップグレードで提供。Switch2ネイティブ版とBattle.net版も同日投入し、Blizzardと新提携。新拡張『Songs of the Past』は2027年予定で先出し公開。",
P("・既存所有者へ“無料”配布で再エンゲージを最大化","・Switch2/Battle.netへ販路を拡張","・有料拡張は2027年に接続し課金導線を確保"),
"無料リマスターで旧作の同接・話題を再点火し、Switch2新規層と2027年の有料拡張・新作へ送客する二段構え。他社の旗艦旧作でも“無料UP→拡張課金”の型が広がる可能性が高い。",
"レガシーIPの無料リマスター送客モデル",
"CD Projekt RED, Blizzard Entertainment, Nintendo","The Witcher 3, Songs of the Past","リメイク・移植,提携,ハード,海外","一次:CDPR"])

rows.append(["ゲーム","『Heroes of Might and Magic III』リメイク発表、名作SLG復刻の潮流",
"https://www.techradar.com/news/live/gamescom-opening-night-live-2026",
"1999年の名作ターン制ストラテジー『HoMM III』は今なお高い評価を保つカルト的定番。近年はレトロ/クラシックの再評価が進み、旧作リメイクが安定した需要層を持つ。",
"ONL 2026でシネマティックトレーラーと共にリメイクを発表。往年のファンと新規層の双方を狙い、ストラテジー領域でも“名作復刻”の動きが鮮明になった。",
P("・ニッチだが熱量の高いレガシー層を的確に取り込む","・ストラテジーでも復刻/リメイクが商機に","・新規AAA疲れの受け皿として旧作再現が機能"),
"リメイクは低リスクで既存ファンダムを収益化できるため、ストラテジーや90〜00年代PC名作の復刻が続く見込み。移植/リメイク専業スタジオの受注も増えると予想。",
"クラシックPC名作リメイク市場",
"Ubisoft","Heroes of Might and Magic","リメイク・移植,海外,トレンド","ミクロ×示唆"])

rows.append(["ゲーム","HoYoverseが癒し系ライフシム『Petit Planet』発表、脱ガチャで多角化",
"https://www.pcgamer.com/news/live/gamescom-opening-night-live-2026-live-coverage/",
"HoYoverseは『原神』『スターレイル』などガチャ型アクションRPGで急成長したが、ジャンル集中はリスク。運営型の課金疲れが指摘される中、より緩いコア外層への裾野拡大が課題だった。",
"ONL 2026でコージー系ライフシム『Petit Planet』を公開。銀河を舞台にした癒し系という、同社の主力アクション/ガチャとは異なる方向性を提示し、プレイヤー層の多角化を狙う。",
P("・主力ジャンル依存からの分散を明確化","・“コージーゲーム”需要の拡大を取り込む","・ライブ運営の課金疲れへの受け皿を用意"),
"癒し系はライト層・長期リテンションと相性が良く、HoYoが運営ノウハウを持ち込めば新たな収益柱になり得る。中国系大手のジャンル多角化と欧米コージー市場争奪が進む見込み。",
"コージーゲームとスタジオのジャンル多角化",
"HoYoverse","Petit Planet, 原神","新作,ライブサービス,海外,トレンド","ミクロ×示唆"])

rows.append(["ゲーム","『Path of Exile』が正式版1.0に到達、ハクスラARPG競争が新局面へ",
"https://www.gamesradar.com/news/live/gamescom-opening-night-live-live-coverage-2026-everything-announced/",
"基本無料ハクスラ『Path of Exile』は長期の早期アクセス/シーズン運営で拡大。『Diablo』系や新作ARPGが乱立し、ライブサービス型ARPGの競争が激化している。",
"ONL 2026で『Path of Exile』のバージョン1.0（正式版）到達を発表。同時にヴァンパイアRPG『The Blood of Dawnwalker』の9/3ローンチも示され、ARPG/ダークファンタジー領域の話題が集中した。",
P("・F2Pハクスラが“正式版”で節目を迎える","・シーズン運営型ARPGの競争が一段と激化","・ダークファンタジー新作の投入が続く"),
"1.0以降はシーズン更新の頻度と課金設計が競争軸に。『Diablo』系との客の奪い合いが続き、ライブ運営の巧拙が明暗を分ける。新作ARPGの参入もさらに増える見込み。",
"ライブサービス型ARPGの運営競争",
"Grinding Gear Games","Path of Exile, The Blood of Dawnwalker","ライブサービス,新作,海外","ARPG競争"])

# ===== MOVIE =====
rows.append(["映画","2026年夏の北米興行が40億ドル突破、ポストコロナで2度目の大台",
"https://variety.com/2026/film/box-office/summer-box-office-4-billion-benchmark-daily-variety-1236831774/",
"北米の劇場興行はコロナ後の回復途上。夏商戦の40億ドル超えは、2023年“バーベンハイマー”期に次ぐ水準で、劇場ビジネスの底堅さを測る指標となる。",
"2026年夏（サマー）の北米興行が40億ドルの大台を突破し、ポストコロナで2度目の到達となった。8月も従来の“夏枯れ”パターンに反して好調を維持し、通年での回復期待を後押ししている。",
P("・ポストコロナ2度目の夏40億ドル超え","・8月の失速がなく需要が持続","・通年100億ドル回復への追い風"),
"大型続編とサプライズヒットが噛み合えば通年100億ドルも視野。配給各社は公開本数を増やし、劇場回帰を鮮明にする。窓口戦略（劇場先行）の見直しも進む可能性。",
"劇場興行のポストコロナ回復基調",
"Sony, Warner Bros., Paramount","","興行,海外,トレンド,消費者インサイト","マクロ"])

rows.append(["映画","『スパイダーマン: ブランニューデイ』2週目も7000万ドルで首位維持",
"https://variety.com/2026/film/box-office/spiderman-box-office-brand-new-day-dominates-end-of-oak-street-debut-1236834084/",
"『スパイダーマン: ブランニューデイ』は歴代2位級のオープニングを記録した大作。2週目のホールドは口コミの強さとフランチャイズ需要の持続性を示す。",
"公開2週目も約7000万ドルを稼ぎ首位を維持。新規参入の『End of Oak Street』は約2100万ドルでデビューした。夏興行の牽引役として大作の粘りが際立つ結果となった。",
P("・2週目7000万ドルで高いホールド率","・新作『End of Oak Street』は2100万ドル始動","・フランチャイズ需要の持続性を再確認"),
"強い2週目継続はソニー/MCUのIP求心力を裏づけ、続く大型公開の弾みに。オリジナル新作は大作の谷間で健闘の余地があり、配給の公開日戦略が重要度を増す。",
"フランチャイズ映画のホールド力",
"Sony Pictures","Spider-Man: Brand New Day","興行,海外,チャート・ランキング","2週目"])

rows.append(["映画","北米興行が年間100億ドル回復へ期待、レガシー各社が公開本数を拡大",
"https://variety.com/2026/film/box-office/movie-theaters-rebound-box-office-obsession-project-hail-mary-1236810214/",
"コロナ後の劇場興行は作品数不足が回復の足かせだった。Z世代の劇場回帰やサプライズヒットが相次ぎ、供給側も本数を戻しつつある。",
"2026年は115〜120本のワイド公開が見込まれ、Warner/Sony/Paramountが本数を拡大、Amazon MGMも劇場配給に回帰。『Dune 3』『Avengers: Doomsday』を控え通年100億ドル回復への期待が高まる。",
P("・年間ワイド公開が115〜120本に回復","・Amazon MGMが劇場配給へ再参入","・大型IP続編が通年興行を下支え"),
"公開本数の回復は稼働率と多様性を高め、劇場エコシステムを安定化。配信専業だったAmazon等の劇場回帰は、窓口戦略が“配信一辺倒”から再調整に向かう兆し。",
"劇場供給の回復と窓口戦略の再調整",
"Amazon MGM, Warner Bros., Sony","Dune 3, Avengers: Doomsday","興行,海外,トレンド","マクロ"])

rows.append(["映画","国内週末興行、『映画ちいかわ』が首位を独走—ファミリー需要で堅調",
"https://eiga.com/ranking/jp/",
"『映画ちいかわ 人魚の島のひみつ』は人気キャラIPの劇場版。夏休み期間のファミリー層を取り込み、公開後も上位を維持し続けている。",
"8/21〜23の国内映画ランキングで『映画ちいかわ』が動員82.5万人・興収12.4億円で首位を継続。キャラクターIPの劇場作品が邦画興行の中心を担う構図が続いている。",
P("・夏休みファミリー需要を安定的に獲得","・公開から時間が経っても上位を維持","・キャラIP劇場版が邦画の主力に"),
"キャラクターIPは幅広い年齢層を集客でき、ロングランと物販の相乗も見込める。今後も既存人気IPの劇場化が邦画興行の柱として続く見通し。",
"キャラクターIP劇場版の集客持続力",
"","ちいかわ","興行,国内,チャート・ランキング","国内週末"])

# ===== ANIME =====
rows.append(["アニメ","『サイバーパンク エッジランナーズ2』10/20にNetflix配信、全10話の新章",
"https://www.animenewsnetwork.com/news/2026-08-20/cyberpunk-edgerunners-2-anime-new-teaser-reveals-october-20-netflix-debut/.240770",
"ゲーム『Cyberpunk 2077』発のアニメ第1作は世界的ヒットし、原作ゲーム本編の売上・同接を再燃させた“逆輸入”成功例。続編はゲームIPのアニメ活用モデルの試金石となる。",
"Anime NYC 2026で新ティザーを公開し、10/20にNetflixで全10話一挙配信と発表。新キャラ集団を描く独立した物語で、監督は五十嵐海が初起用。世界同時配信を前提とした設計。",
P("・ゲームIP発アニメの続編が世界配信で投入","・新キャラ/新作画で“アンソロジー型”に","・原作ゲームへの再送客効果が焦点"),
"配信ヒットが再びゲーム本編の需要を押し上げれば、ゲーム各社が“アニメ＝新作前の需要喚起装置”として内製・専属契約を増やす契機に。アニメ制作の受注先がゲーム企業へ広がる可能性。",
"ゲームIPのアニメ化と逆輸入マーケティング",
"CD Projekt, Studio Trigger, Netflix","Cyberpunk: Edgerunners, Cyberpunk 2077","配信,新作,海外,トレンド","ゲーム→アニメ"])

rows.append(["アニメ","TVアニメ『Fate/strange Fake』新シーズン制作決定、TYPE-MOON人気作",
"https://www.animenewsnetwork.com/news/2026-08-22/fate-strange-fake-tv-anime-confirms-new-season/.240827",
"『Fate』は多メディア展開が続く長寿IPで、スピンオフ『strange Fake』も根強い人気。第1期の完結後、続編の可否が注目されていた。",
"『Fate/strange Fake』TVアニメの新シーズン制作が正式に確定。原作の人気とシリーズ全体の求心力を背景に、フランチャイズの継続的な拡張が続く。",
P("・人気スピンオフの続編が正式決定","・長寿IPのアニメ供給が継続","・原作/ゲームとの相乗が期待"),
"続編投入は原作小説・関連ゲームの再注目を促す。TYPE-MOON系IPは今後も配信前提の続編・スピンオフを積み増し、グローバルでの露出を最大化する見込み。",
"長寿フランチャイズの続編供給戦略",
"TYPE-MOON","Fate/strange Fake, Fate","新作,国内,配信","続編"])

rows.append(["アニメ","Anime NYC 2026でCrossed Heartsが新レーベル設立・新規ライセンス発表",
"https://www.animenewsnetwork.com/press-release/2026-08-24/crossed-hearts-to-launch-studio-hearts-and-unveil-new-licenses-live-at-anime-nyc-2026/.240898",
"北米最大級のアニメイベントAnime NYCは、海外ローカライズ各社の新作ライセンス発表の場。北米市場の拡大で英語版パブリッシャー間の獲得競争が激化している。",
"英語版ローカライズを手がけるCrossed Heartsが、Anime NYC 2026で自社インプリント“Studio Hearts”の設立と新規ライセンスを発表。北米での翻訳出版・配信の獲得競争が一段と強まった。",
P("・北米ローカライズ大手が自社レーベルを新設","・新規ライセンス獲得競争が激化","・イベントが権利取引のショーケースに"),
"北米のアニメ/マンガ需要拡大で、ライセンス費の高騰と独占契約の争奪が進む。日本の権利元にとっては海外二次収益の交渉力が増し、窓口戦略の再設計が必要になる。",
"北米ローカライズ市場のライセンス競争",
"Crossed Hearts","","権利・ライセンス,海外,出版","Anime NYC"])

rows.append(["アニメ","『映画ちいかわ』累計興収92億円、100億の大台迫りキャラIP映画が絶好調",
"https://cinema.eiga.com/ranking/20260810/",
"アニメ映画は近年、キャラクターIPを軸にした作品が邦画興行を牽引。『映画ちいかわ』は7月末公開後、動員を積み上げてきた。",
"8/14〜16集計時点で累計動員645万人・興収92億円に達し、早くも100億円の大台が視野に。キャラクターIPアニメ映画の集客力の高さを改めて示した。",
P("・累計92億円で100億円が射程","・キャラIPアニメ映画の強さが持続","・ファミリー〜大人まで幅広く集客"),
"100億円到達なら非“劇場版TVアニメ”系キャラIPの実力を証明。玩具・物販との相乗も大きく、出版・玩具各社がキャラIPの劇場化に一層傾斜する見込み。",
"キャラクターIPアニメ映画の興行力",
"","ちいかわ","興行,国内,チャート・ランキング,消費者インサイト","ミクロ×示唆"])

# ===== MANGA =====
rows.append(["漫画","2026年上半期の出版市場は7664億円、紙は減も電子コミックはプラス維持",
"https://shuppankagaku.com/",
"出版市場は紙の縮小を電子（特に電子コミック）が補う構造が続く。上半期の集計は年間トレンドを占う重要指標となる。",
"出版科学研究所によると2026年上半期の出版市場（紙＋電子）は前年同期比1.3%減の7664億円。紙は3.1%減、電子は2.0%増で、電子出版のうちコミックはプラスを確保した。",
P("・市場全体は微減も電子コミックは成長維持","・紙の縮小トレンドは継続","・電子コミックが市場の下支え役"),
"電子コミックの成長鈍化が進めば、出版社は縦読み/webtoonや海外配信など新販路の開拓を急ぐ。紙依存の高い作品・レーベルの構造転換が一段と迫られる。",
"電子コミック成長鈍化と販路多様化",
"","","出版,業績,国内,トレンド","マクロ"])

rows.append(["漫画","業界紙・新文化が特集『マンガ復権へ』、図書館と小学館の共同施策に注目",
"https://www.shinbunka.co.jp/",
"紙コミックの縮小が続く中、読者接点の再構築が課題。公共図書館とのマンガ振興は新たな読者育成・地域連携の実験となる。",
"出版業界紙・新文化の8/20号が「マンガ復権へ」を特集し、白河市立図書館と小学館の共同の取り組みを取り上げた。図書館を起点にしたマンガ需要喚起の事例として注目される。",
P("・図書館×出版社でマンガ読者接点を再構築","・地域連携型の需要喚起が試行","・紙コミック縮小への対抗策の一例"),
"図書館連携が読者育成に効けば、他出版社・自治体へ横展開する余地。紙・電子双方の入口を増やす“読者接点の多層化”が、出版社の中期戦略に組み込まれる可能性。",
"公共図書館とマンガ振興の連携",
"小学館","","出版,国内,トレンド","ニッチ×示唆"])

rows.append(["漫画","週間マンガ単行本ランキングで『薬屋のひとりごと』が首位、既刊IPが強い",
"https://natalie.mu/comic/news/683681",
"単行本市場ではメディアミックス連動作が上位を占める傾向。アニメ放送・続編が控える作品ほど原作の売上が伸びやすい。",
"コミックナタリー週間ランキング（7/27〜8/2）で『薬屋のひとりごと』17巻が首位。『宇宙兄弟』46巻、『ONE PIECE』115巻など既刊シリーズが上位を占め、新規IPを抑える構図が続いた。",
P("・アニメ展開中の『薬屋』最新刊が首位","・長期連載の既刊IPが安定して上位","・メディアミックス連動が販売を押し上げ"),
"アニメ新シーズンや映像化の“前後”で原作が伸びる相乗が定着。出版社は新規発掘と並行し、既存ヒットIPの映像化タイミングと連動した販促設計を強める見込み。",
"メディアミックス連動とコミック販売",
"集英社, 小学館, 講談社","薬屋のひとりごと, ONE PIECE","チャート・ランキング,出版,国内,消費者インサイト","ミクロ×示唆"])

# ===== NOVEL =====
rows.append(["小説","出版取次・中央社の25年5月期決算は減収減益、総売上高191.8億円",
"https://www.shinbunka.co.jp/cat/newsflash/kessan",
"出版流通の中核である取次は、紙市場の縮小と物流コスト増で経営環境が厳しい。中堅取次の業績は書店・出版流通全体の体力を映す。",
"新文化によると、出版取次の中央社が2025年6月〜2026年5月期決算を承認。総売上高191億8420万円（前年比4.2%減）で営業減益となり、減収減益の決算となった。",
P("・取次売上は前年比4.2%減で減収減益","・紙流通の縮小が取次経営を圧迫","・書店・流通全体の体力低下を示唆"),
"取次の収益悪化が続けば、配本・返品の仕組みや直取引・電子シフトの加速が不可避に。出版社が流通を内製化・DTC化する動きが強まり、商流の主導権が移る可能性。",
"出版取次の収益悪化と流通再編",
"中央社","","業績,出版,国内","出版流通"])

rows.append(["小説","文芸情報誌『ダ・ヴィンチ』8月号、ノーラン監督『オデュッセイア』原案コミカライズ掲載",
"https://ddnavi.com/book/",
"『ダ・ヴィンチ』は書籍・文芸を横断的に扱う情報誌。近年は映画・アニメ等の話題を積極的に取り込み、文芸と他メディアの接点を広げている。",
"『ダ・ヴィンチ』8月号が、クリストファー・ノーラン最新作『オデュッセイア（The Odyssey）』の原案コミカライズを掲載。万城目学・月村了衛らのインタビューと併せ、映画×文芸×漫画の横断的な誌面構成となった。",
P("・映画原案を文芸誌がコミカライズで先行展開","・映画→漫画→書籍の相互送客を狙う","・文芸誌が話題IPの接点として機能"),
"公開前映画のコミカライズ先行掲載は、文芸誌の集客と映画の話題喚起を両立させる手法として広がる可能性。出版社が映画IPの二次展開に前のめりになる流れが強まる。",
"文芸誌と映画IPのクロスメディア展開",
"KADOKAWA","The Odyssey","出版,映像化,国内","横断×示唆"])

# ===== MUSIC =====
rows.append(["音楽","Mrs. GREEN APPLE、3年ぶり6thアルバムの全16曲・曲順を発表、書き下ろし9曲",
"https://realsound.jp/2026/08/post-2500087.html",
"Mrs. GREEN APPLEはタイアップを軸に高い露出を維持する国民的バンド。3年ぶりのオリジナルフルアルバムは、楽曲がIP販促と結びつく“タイアップ経済”の縮図でもある。",
"9/30発売の6thアルバム（タイトルは9/9発表）全16曲の曲名・曲順を公開。書き下ろし新曲9曲に加え、映画『スパイダーマン: ブランニューデイ』日本主題歌やアニメ・朝ドラ・CMタイアップ曲を集約した構成。",
P("・3年ぶりフルアルバムで書き下ろし9曲","・映画/アニメ/朝ドラ/CMのタイアップ曲を集約","・楽曲がIP販促装置として機能"),
"タイアップ曲中心の設計は露出最大化と相性が良く、配信での再生も稼ぎやすい。今後もアーティストの制作が主題歌・CM連動を前提に組まれ、“タイアップ主導”の楽曲経済が強まる見込み。",
"タイアップ主導の楽曲経済とアルバム設計",
"ユニバーサルミュージック","Mrs. GREEN APPLE","新作,権利・ライセンス,国内,消費者インサイト","ミクロ×示唆"])

rows.append(["音楽","ソニーとSpotifyが新グローバル契約、出版部門は米国で直接契約に移行",
"https://www.musicbusinessworldwide.com/sony-spotify-strike-expanded-global-deal-including-a-direct-agreement-for-sony-music-publishing-in-the-us/",
"メジャーとDSPの契約は、ストリーミング分配率やAI利用の条件を左右する要。ソニーは原盤・出版の両輪を持ち、交渉での発言力が大きい。",
"ソニーミュージックとSpotifyが新たなグローバル契約を締結。ソニー・ミュージックパブリッシングは米国でSpotifyと直接契約に移行し、出版権の対価やAI関連の条件で主導権を確保する狙い。",
P("・原盤に加え出版権でも直接契約へ","・米国で出版の対価条件を再設定","・AI時代の権利対価を睨んだ布石"),
"出版の直接契約は分配の透明化と対価向上につながる。他メジャーも追随すれば、DSPのコスト構造とAI楽曲の許諾条件に波及し、権利者優位の再設計が進む可能性。",
"メジャーとDSPの権利対価再交渉",
"Sony Music, Spotify","","提携,配信,海外,権利・ライセンス","海外業務提携"])

rows.append(["音楽","ユニバーサルがSpotify株一部売却で403億円超、MerlinはAIカバーで許諾",
"https://www.musicbusinessworldwide.com/from-sunos-vinyl-move-to-spotifys-fan-made-merlin-deal-its-mbws-weekly-round-up/",
"メジャー各社はDSP株の含み益をキャッシュ化しつつ、AI生成音楽の商用化にどう関与するかを模索。独立系連合Merlinの動向も業界全体の指標となる。",
"UMGは2026年上半期にSpotify株の一部売却で約4.03億ユーロ（約4.67億ドル）を得て、アーティストへ約1.31億ドルを分配。独立系連合MerlinはSpotifyのAIによるファンメイド・カバー/リミックス機能でライセンス契約を結んだ。",
P("・UMGがDSP株売却益をアーティストに還元","・MerlinがAIカバー機能で許諾契約","・AI二次利用の“許諾済み化”が前進"),
"AIカバー/リミックスの許諾モデルが定着すれば、UGC商用化の枠組みが音楽から他ジャンルへ波及する余地。権利者は株式益と許諾料の二重収益で交渉力を高める。",
"AI二次利用の許諾モデルと権利収益",
"Universal Music Group, Merlin, Spotify","","AI,権利・ライセンス,海外,業績","海外"])

rows.append(["音楽","Spotifyの有料会員が3億人突破、音声配信で初、Q2営業益も過去最高水準",
"https://www.musicbusinessworldwide.com/spotify-hits-300-million-premium-subscribers-in-q2-202/",
"サブスク音楽は成熟局面に入りつつあり、会員増と収益性の両立が焦点。Spotifyの規模は業界全体の分配原資に直結する。",
"Spotifyは2026年Q2にPremium有料会員が3億人に到達（前年比+9%）。音声配信サービスで初の大台で、MAUは777万人ではなく7.77億人に拡大。総売上47.8億ユーロ、粗利率33.4%と過去最高、営業益6.55億ユーロを計上した。",
P("・音声配信で初の有料3億人","・粗利率33.4%と収益性が最高水準","・分配原資の拡大で権利者にも波及"),
"会員増の鈍化を値上げと収益性改善で補う局面。粗利改善はコンテンツ費交渉やAI・ライブ連携（Reserved等）への投資余力を生み、権利者との分配交渉に影響する。",
"サブスク成熟期の会員増と収益性",
"Spotify, Live Nation","","業績,配信,海外","マクロ"])

added=0; skipped=0
for r in rows:
    date_v=TODAY
    url=r[2].strip()
    if url in existing_urls:
        skipped+=1; continue
    genre,head,u,haikei,summary,points,yosoku,theme,companies,ip,tags,biko=r
    ws.append([date_v,genre,head,u,haikei,summary,points,yosoku,theme,companies,ip,tags,biko])
    style_row(ws, ws.max_row)
    existing_urls.add(url); added+=1

print(f"[日次ニュース] added={added} skipped={skipped}")

# ---------- 横断的問い ----------
wsq=wb['横断的問い']
existing_q=set(str(row[1]).strip() for row in wsq.iter_rows(min_row=2, values_only=True) if row[1])
qrows=[
 [TODAY,"ゲームIP発アニメが原作本編の売上を再燃させる“逆輸入”モデル（Edgerunners）は、漫画原作アニメ→原作重版と同じ構造としてゲーム業界の標準戦略になるか？",
  "『サイバーパンク エッジランナーズ2』10/20 Netflix配信決定(ANN)","CD Projekt, Netflix, Studio Trigger","漫画→アニメ→原作重版の相乗／映像を新作前の需要喚起装置に","オープン","アニメの配信ヒットがゲーム本編DAU/売上を押し上げるかを、配信後の同接・販売データで検証したい","横断"],
 [TODAY,"SpotifyのAIファンメイド・カバー/リミックスのライセンス（Merlin）で始まる“UGC商用化の許諾モデル”は、ゲームのMOD/UGC収益化やアニメの二次創作商用ガイドラインにも波及するか？",
  "Merlin×SpotifyがAIカバー/リミックスでライセンス契約(MBW)","Merlin, Spotify","ゲームUGC・MOD収益化／二次創作の商用許諾／権利者への分配設計","オープン","他ジャンルで“許諾済みUGC”の分配スキームが提示されるかを、PF各社のガイドライン改定で観測","横断"],
]
qadded=0
for q in qrows:
    if str(q[1]).strip() in existing_q: continue
    wsq.append(q); style_row(wsq, wsq.max_row); existing_q.add(str(q[1]).strip()); qadded+=1
print(f"[横断的問い] added={qadded}")

# ---------- メディアミックス追跡 ----------
wsm=wb['メディアミックス追跡']
existing_m=set()
for row in wsm.iter_rows(min_row=2, values_only=True):
    if row[1] and row[4]:
        existing_m.add((str(row[1]).strip(), str(row[4]).strip()))
mrows=[
 [TODAY,"Cyberpunk 2077","ゲーム","CD Projekt RED","アニメ","Netflix完全新作アニメ『エッジランナーズ2』全10話（新キャラ・新作画）","2026-08-20","2026-10-20","公開間近","CD Projekt, Studio Trigger, Netflix","ゲームIP発アニメの続編。原作本編への逆輸入送客が焦点"],
 [TODAY,"Fate/strange Fake","小説","TYPE-MOON","アニメ","TVアニメ新シーズン制作決定","2026-08-22","未定","発表","TYPE-MOON","人気スピンオフの続編。原作/ゲームとの相乗を狙う"],
 [TODAY,"The Odyssey（映画）","映画","Christopher Nolan/配給元","漫画","『ダ・ヴィンチ』8月号で原案コミカライズ掲載","2026-08","掲載済","公開済","KADOKAWA","映画→漫画の逆方向展開。公開前の話題喚起を兼ねる"],
]
madded=0
for m in mrows:
    key=(str(m[1]).strip(), str(m[4]).strip())
    if key in existing_m: continue
    wsm.append(m); style_row(wsm, wsm.max_row); existing_m.add(key); madded+=1
print(f"[メディアミックス追跡] added={madded}")

# ---------- イベントカレンダー ステータス更新 ----------
wse=wb['イベントカレンダー']
today_d=date(2026,8,26)
def parse_d(v):
    if v is None or v=="" : return None
    if isinstance(v,(datetime,date)):
        return v.date() if isinstance(v,datetime) else v
    s=str(v).strip()
    for fmt in ("%Y-%m-%d","%Y/%m/%d"):
        try: return datetime.strptime(s,fmt).date()
        except: pass
    return None
updated_e=0
# columns: 0開催日 1終了日 2名 8ステータス 9最終更新
for row in wse.iter_rows(min_row=2):
    name=row[2].value
    if not name: continue
    st_cell=row[8]; upd_cell=row[9]
    sd=parse_d(row[0].value); ed=parse_d(row[1].value)
    cur=str(st_cell.value) if st_cell.value else ""
    if cur=="日程未確定" or sd is None:
        continue
    end=ed if ed else sd
    if end < today_d: newst="終了"
    elif sd <= today_d <= end: newst="開催中"
    else: newst="開催予定"
    if newst!=cur:
        st_cell.value=newst
        upd_cell.value=TODAY
        st_cell.font=FONT; st_cell.alignment=ALIGN
        upd_cell.font=FONT; upd_cell.alignment=ALIGN
        updated_e+=1
print(f"[イベントカレンダー] status updated={updated_e}")

# ---------- 今後起きそうなこと ----------
wsp=wb['今後起きそうなこと']
existing_p=[str(row[1]).strip() for row in wsp.iter_rows(min_row=2, values_only=True) if row[1]]
pred="ゲームIP発アニメの配信ヒットが新作前の需要喚起に有効と実証されると、ゲーム大手がアニメを“内製の販促装置”として抱え込み、専属スタジオ契約や制作出資を増やす。結果、アニメ制作会社の受注先が製作委員会からゲーム企業へシフトし、元請の主導権が移る。"
def near_dup(t):
    for e in existing_p:
        if e[:18]==t[:18]: return True
    return False
padded=0
if not near_dup(pred):
    wsp.append([TODAY,pred,"中","6〜12ヶ月","CD Projekt, Netflix, 大手ゲーム各社",
      "『サイバーパンク エッジランナーズ2』10/20 Netflix配信決定（Anime News Network）",
      "大手ゲーム企業のアニメ内製部門/専属契約/制作出資の発表と、主要アニメ制作会社の“ゲーム企業向け元請”比率の推移を観測","横断"])
    style_row(wsp, wsp.max_row); padded+=1
print(f"[今後起きそうなこと] added={padded}")

# ---------- ゲームデザイン・トレンド ----------
wsd=wb['ゲームデザイン・トレンド']
existing_d=[str(row[1]).strip() for row in wsd.iter_rows(min_row=2, values_only=True) if row[1]]
dname="レガシー旧作の無料リマスター配布による“再エンゲージ→新作送客”"
dadded=0
if not any(dname[:12]==e[:12] for e in existing_d):
    wsd.append([TODAY,dname,"マネタイズ/運営","台頭",
      "10年級の旗艦旧作を既存所有者へ無料アップグレードで再配布し、同接・話題を再点火。新販路（新ハード/ストア）開放と有料拡張・続編への送客をセットで設計する運営型の型。",
      "The Witcher 3: Wild Hunt Remastered, Heroes of Might and Magic III Remake",
      "新規AAA開発費の高騰でレガシー資産の再収益化が合理化。無料配布で母数を回復し、拡張・続編・新ハード版で回収する導線が成立",
      "無料UPで母数を回復→有料拡張/続編/新ハード版で回収する“旧作エバーグリーン化”が標準化。移植/リマスター専業の受注も拡大",
      "gamescom Opening Night Live 2026（GameSpot／CD Projekt発表, 8/25）",4])
    style_row(wsd, wsd.max_row); dadded+=1
print(f"[ゲームデザイン・トレンド] added={dadded}")

wb.save(PATH)
print("SAVED")
