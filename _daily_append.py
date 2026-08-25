# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, Alignment
from datetime import date, datetime

TODAY = "2026-08-25"
PATH = "master/ゲーム業界・時事情報収集.xlsx"
wb = openpyxl.load_workbook(PATH)
FONT = Font(name="Yu Gothic", size=11)
ALIGN = Alignment(wrap_text=True, vertical="top")

def style_row(ws, r, ncols):
    for c in range(1, ncols+1):
        cell = ws.cell(row=r, column=c); cell.font=FONT; cell.alignment=ALIGN

def append_row(ws, values):
    r = ws.max_row + 1
    for i, v in enumerate(values, start=1):
        ws.cell(row=r, column=i, value=v)
    style_row(ws, r, len(values)); return r

# ---------- 日次ニュース ----------
ws = wb["日次ニュース"]
existing_urls=set()
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[3]: existing_urls.add(str(row[3]).strip())

news = [
 [TODAY,"ゲーム","gamescom Opening Night Live 2026開幕、Witcher3新DLCやFF7 Revelation等を世界初公開",
  "https://screenrant.com/gamescom-2026-opening-night-time-where-watch-confirmed-reveal/",
  "8/25(米時間)にGeoff Keighley司会のgamescom ONL 2026が開催。ケルンの大型展示会gamescom(8/26-30)の開幕を飾るショーケースで、E3消滅後は実質的な世界最大級の新作発表の場になっている。",
  "The Witcher 3の新DLC『Songs of the Past』、FF7 Revelation続報、Gears of War: E-Day、Metro 2039(生演奏付き)、Path of Exile 2、Silent Hill: Townfallなどが登場。旧作の追加展開と続編が目立つ構成。",
  "E3不在で単独ショーケースが新作発表のハブ化\n10年前のWitcher3への新DLCなどレガシーIP再収益化が顕著\n続編・リメイク偏重で完全新規AAAは限定的\nTGS(9月)前の海外先行イベント",
  "レガシーIPへの有料大型DLC投下と続編集中は当面続き、新規IPリスクを避ける保守化が鮮明に。9月TGSでは和ゲーの続編・移植が主軸となり、AAA新作より『既存資産の再展開』で年末商戦を組み立てる流れが強まる。",
  "レガシーIP再収益化とショーケース経済圏","CD Projekt RED, Square Enix, Xbox Game Studios","The Witcher, Final Fantasy VII, Gears of War, Metro, Silent Hill",
  "新作,ライブ・イベント,海外,トレンド","gamescom ONL 2026(8/25)"],
 [TODAY,"ゲーム","任天堂26年4-6月期は純利益+53.5%、Switch2は2年目で販売-34%に減速",
  "https://game.watch.impress.co.jp/docs/news/2036140.html",
  "任天堂が2026年3月期第1四半期決算を発表。前年のSwitch2発売ブーストの反動局面に入った。",
  "連結純利益は前年同期比53.5%増の1474億円。Switch2は四半期販売が34%減の382万台、世界累計は2368万台。高採算ソフトと映画『スーパーマリオギャラクシー・ムービー』好調で営業益は2.5倍。",
  "ハード2年目で早くも販売失速\n利益はソフト/映画の高マージンが牽引\nハード台数よりアタッチ率・IP収益が利益源に\n映画事業が関連収入を押し上げ",
  "台数依存から『1台あたりソフト/IP収益最大化』への構造転換が進む。ハード値下げは当面回避し、映画・テーマパーク等の非ゲーム収益でIPライフサイクルを延ばす戦略が続く。年末商戦の目玉ソフト次第でハード再加速を狙う。",
  "据置機2年目の失速とIP多角化の採算","任天堂","Nintendo Switch 2, スーパーマリオ",
  "業績,ハード,国内,消費者インサイト","26年3月期Q1"],
 [TODAY,"ゲーム","バンダイナムコHD Q1が過去最高益、ガンダムが5四半期連続600億円規模を維持",
  "https://gamebiz.jp/news/430754",
  "バンダイナムコHDが第1四半期決算を発表。トイホビー(ガンプラ/TCG/ガシャポン/一番くじ)が業績を牽引した。",
  "大幅増収増益で過去最高業績。ガンダムは5四半期連続で600億円規模を維持し、中間期見通しを減収減益→増収増益へ上方修正。IP軸の多角展開が奏功した。",
  "ゲーム以外のフィジカルIP商材が利益の柱\nガンダム単独IPが安定収益源に\n中間期ガイダンスを上方修正\nTCG/一番くじ等の高粗利物販が寄与",
  "単一メガIPを玩具・TCG・ゲーム・映像で回す『IPアクシス』モデルの優位が続く。海外トイホビー展開と新規IP育成が次の焦点で、ガンダム依存の分散が中期課題になる。",
  "フィジカルIPマネタイズとメガIP依存リスク","バンダイナムコHD","ガンダム","業績,国内","26年3月期Q1"],
 [TODAY,"ゲーム","東京ゲームショウ2026、史上初の5日間・30周年で9/17-21開催、759社が出展",
  "https://gamebiz.jp/news/420741",
  "CESA主催のTGS2026が開催概要を発表。30周年の節目で規模を拡大する。",
  "会期は9/17-21(ビジネスデイ2日+一般公開3日)、幕張メッセでハイブリッド開催。テーマは『史上最長、遊びづくしの5DAYS』。7/8時点で759社(国内484/海外275)、51の国と地域が出展、来場30万人想定。",
  "史上初の5日間開催で商談機会を拡大\n海外出展275社と国際化が進む\nBtoBとBtoCを明確に分離\ngamescom(8月)に続く秋の新作発表の場",
  "会期長期化はインディー/中小の露出機会を増やし、商談特化の色が強まる。海外パブリッシャーの日本市場再評価とアジア展開のハブ化が進行。年末商戦前の最終プロモ舞台として重要度が上がる。",
  "展示会の商談機能とインディー露出設計","CESA","東京ゲームショウ","展示会,ライブ・イベント,国内","9/17-21幕張メッセ"],
 [TODAY,"ゲーム","Paramountの米ワーナー買収でWB Games主要スタジオの去就が不透明に",
  "https://en.wikipedia.org/wiki/Proposed_acquisition_of_Warner_Bros._Discovery",
  "NetflixとParamount SkydanceがWBD争奪。26年2月にWBD取締役会がParamountの1109億ドル(1株31ドル)提案を優越と判断、Netflixは対抗せず撤退。WBD傘下にRocksteady/NetherRealm/Avalanche/TT Games等のゲームスタジオがある。",
  "買収の帰趨が固まる中、有力ゲームスタジオ群(バットマン/モータルコンバット/ホグワーツ・レガシー/レゴ)の売却・分社観測が浮上。メディア再編がゲーム部門の再配置に直結する構図。",
  "映像大手の統合がゲーム資産再編を誘発\nRocksteady等の身売り/分社観測\nDCやハリポタIPのゲーム化権の帰属が焦点\n買い手候補に他パブリッシャーやファンド",
  "WB Gamesの一部スタジオ/IPライセンスが切り出され、事業パブリッシャーや国家系ファンドが買い手に回る可能性。DC/ハリポタのゲーム化は権利再交渉で座組が変わり、日本勢のIPライセンス獲得機会も生じうる。",
  "メディアコングロマリット再編とゲーム資産の切り出し","Paramount Skydance, Warner Bros. Discovery, Netflix","Batman: Arkham, Mortal Kombat, Hogwarts Legacy, LEGO",
  "M&A,海外,権利・ライセンス","WBD買収の余波"],
 [TODAY,"ゲーム","2026年上期のゲーム業界レイオフが8千-1.2万人規模、105件超で人材流出続く",
  "https://gaminglayoffs.com/",
  "MicrosoftのXbox大規模再編(7月、id/Bethesda等に波及)を含め、2026年上期は業界横断で人員削減が継続。GamesBeat集計で7月までに105件のレイオフ。",
  "2026年上期のグローバルなゲーム業界人員削減は推計8000-12000人。大手の統合・再編と受託縮小が背景で、熟練人材の流出と中小スタジオへの拡散が進む。",
  "大手の再編がレイオフの主因\n熟練人材が独立/中規模スタジオへ流出\n受託開発の需要収縮\nAI導入による工程圧縮圧力も併存",
  "流出した熟練者による中規模スタジオ新設・インディー化が加速し、パブリッシャーは自社雇用より外部委託/パブリッシング契約に軸足を移す。人材の分散が数年後の新規IP供給源になる一方、短期は品質・納期リスクが増す。",
  "業界再編下の人材流動とスタジオ新設サイクル","Microsoft, Xbox Game Studios","","人事・組織,海外,トレンド","2026上期集計"],
 [TODAY,"映画","『スパイダーマン: ブランニューデイ』が3.55億ドルで歴代2位デビュー、6日で10億ドル突破",
  "https://variety.com/2026/film/box-office/spider-man-brand-new-day-billion-dollars-six-days-1236826658/",
  "ソニー/マーベルの新作スパイダーマンが北米で記録的スタートを切った。",
  "北米オープニング3.55億ドルは『アベンジャーズ/エンドゲーム』(3.57億)に迫る歴代2位。世界でも歴代2位級の立ち上がりで、公開6日で全世界10億ドルを突破、2026年最高興収に。",
  "スーパーヒーローIPの集客力が健在\nオープニング偏重(初動型)興行の極大化\n2026年は10億ドル作品が続出\n劇場体験×大型IPの相性",
  "MCU/大型IPの初動最大化戦略が続き、公開窓口の短縮と配信移行の綱引きが激化。ヒーロー疲れ論の一方で『イベント級IP』への集中投資が強まり、中規模作品の劇場枠を圧迫する二極化が進む。",
  "イベント映画の初動設計と興行二極化","Sony Pictures, Marvel/Disney","Spider-Man","興行,海外,消費者インサイト","Variety(2026/8)"],
 [TODAY,"映画","ノーラン『The Odyssey』が11億ドル、監督自己最高かつIMAX歴代最高興収に",
  "https://variety.com/2026/film/box-office/summer-box-office-4-billion-benchmark-daily-variety-1236831774/",
  "ユニバーサルのクリストファー・ノーラン監督作が大ヒットした。",
  "『The Odyssey』が世界11億ドルを記録し、ノーラン監督の自己最高興収かつIMAX公開作として歴代最高に。非フランチャイズのオリジナル大作でも大型ヒットが可能なことを示した。",
  "非フランチャイズの作家性大作がメガヒット\nIMAX/大型フォーマットが単価と動員を牽引\nノーランのブランド力\n劇場先行の価値を再証明",
  "プレミアム大型フォーマット(IMAX等)への上映集中と単価上昇が続き、『劇場でしか得られない体験』を軸にした大作戦略が強化される。オリジナル大作の成功は続編偏重への反証として企画多様化を後押しする。",
  "プレミアムフォーマット興行と作家性大作の採算","Universal Pictures, IMAX","The Odyssey","興行,海外,トレンド","Variety(2026/8)"],
 [TODAY,"映画","2026年夏の北米興行が40億ドル超えへ、ポストコロナ最大の夏に",
  "https://en.wikipedia.org/wiki/List_of_2026_box_office_number-one_films_in_the_United_States",
  "大型IP作とオリジナル大作の連続ヒットで夏興行が回復した。",
  "2026年夏の北米興行は40億ドルの節目超えが視野に。コロナ後で40億ドル超えは2023年(バーベンハイマー)以来で、10億ドル作品数もパンデミック以降最多となった。",
  "夏興行がコロナ前水準へ回復\n10億ドル級ブロックバスターの多発\n劇場ビジネスの底堅さを示す\n配信一辺倒論への揺り戻し",
  "劇場の回復基調は配信各社の劇場公開戦略(窓口の再設計)を促す。大作の劇場先行→配信の順序が固まり、劇場を『マーケティングと一次収益の場』として再評価する動きが強まる。",
  "ポストコロナ興行回復と公開窓口戦略","","","興行,市場,海外,トレンド","Variety(2026/8)"],
 [TODAY,"映画","Paramount SkydanceがワーナーをEV1109億ドルで買収、Netflixは対抗断念",
  "https://about.netflix.com/en/news/netflix-to-acquire-warner-bros",
  "WBD争奪戦でParamount SkydanceがNetflixを上回る提案を提示した。",
  "26年2月26日、WBD取締役会がParamountの1株31ドル・総額約1109億ドルの提案を優越と判断、Netflixは対抗せず撤退。ハリウッド大手の統合が現実味を帯び、DC/ハリポタ等のIP運用主体が再編される。",
  "映像大手の大型統合が進行\nDC/ハリウッドIPの運用主体が変わる\n配信×スタジオの垂直統合競争\nゲーム/出版等の周辺IP事業にも波及",
  "統合後はライブラリとIPの集中が進み、配信の独占供給とライセンス方針の見直しが一気に進む。日本の配給/配信各社は独占コンテンツ調達コスト増に直面し、自社出資・共同製作へのシフトが加速する。",
  "メディア統合とIP集中の競争政策","Paramount Skydance, Warner Bros. Discovery, Netflix","DC, Harry Potter","M&A,海外,配信,権利・ライセンス","WBD買収"],
 [TODAY,"映画","『映画ちいかわ 人魚の島のひみつ』が週末12.4億で首位継続、キッズ/ファミリーIPの底力",
  "https://eiga.com/ranking/jp/",
  "東宝配給の劇場アニメが公開後も動員を伸ばし国内興行首位を維持した。",
  "8/21-23の週末3日で動員82.5万人・興収12.4億円と前週を上回り首位。SNS発の低年齢〜大人層のキャラIPが、続編大作に伍する集客力を示した。",
  "SNS発キャラIPの劇場動員力\n前週超えのロングラン型ヒット\nファミリー層のリピート消費\nグッズ/配信への波及基盤",
  "キャラクターIPの映画化は『短期爆発』より『ロングラン+多角展開』で収益を積む型が定着。出版・グッズ・配信を束ねた製作委員会が、低予算高回転のファミリー映画枠を安定収益源として拡大させる。",
  "キャラクターIP映画のロングラン設計","東宝","ちいかわ","興行,国内,消費者インサイト,映像化","映画.com(8/21-23)"],
 [TODAY,"映画","金城宗幸原作『ブルーロック』実写映画が公開、人気サッカー漫画の実写化が本格化",
  "https://ddnavi.com/article/1325113/a/",
  "講談社の人気サッカー漫画が8/7に実写映画公開、国内ランキング入りした。",
  "潔世一役に高橋文哉、凪誠士郎役にK(&TEAM)を起用した実写版が公開。アニメ・ゲームに続く実写化で、少年漫画IPのメディアミックス最終段としての実写映像化が加速している。",
  "少年漫画IPの実写映画化\nアニメ→実写のマルチ展開\nアイドル起用による集客\n原作売上への還流",
  "ヒット漫画の『アニメ化→ゲーム化→実写化』フルライン展開が標準化。実写は海外配信を意識したキャスティングと同時展開が増え、原作出版社の映像出資比率が上がる。",
  "少年漫画IPの実写化と出版社の映像出資","講談社, 東宝","ブルーロック","映像化,国内,興行","ダ・ヴィンチ2026年9月号特集"],
 [TODAY,"アニメ","アニメ制作市場が初の4000億円突破(2025年+9.9%)、TDBは26年ピークアウトを警告",
  "https://www.tdb.co.jp/report/industry/20260819-animation25y/",
  "帝国データバンクが8/19にアニメ制作市場調査を公表。配信需要と劇場大作で25年は最高を更新した。",
  "2025年市場は前年比9.9%増の4065億円で初の4000億円超。ただ26年は大型劇場版の反動、配信投資の一服、アニメーター不足で5年ぶり減少の可能性。好調の裏で人手が制約になる。",
  "配信(Netflix等)向け需要が25年を押し上げ\n26年はピークアウト懸念\nアニメーター不足が構造的制約\n元請参入と委員会出資で二次収益取り込み",
  "制作能力の上限が需要の天井になり、単価上昇と内製化/前方統合が進む。配信各社は制作会社の囲い込み(出資/専属)を強め、スタジオは受託から権利保有側へ移る交渉力の反転が始まる。",
  "制作キャパ制約とスタジオの前方統合","帝国データバンク, Netflix","","業績,配信,国内,トレンド","TDB(2026/8/19)"],
 [TODAY,"アニメ","スタジオが製作委員会出資・元請参入で二次収益を取り込む垂直統合が進む",
  "https://www.projectdesign.jp/articles/news/05df7143-e43b-4de6-a8cf-f23a5459a86d",
  "TDB調査で、グロス請〜専門スタジオが元請・プロデュース機能を強化する動きが顕在化した。",
  "スタジオが製作委員会へ出資して二次収益(配信/商品化/海外)を還流させ、自社オリジナルやタイアップの製作受託に踏み込む例が増加。受託一辺倒からの脱却が進む。",
  "受託から権利保有・出資側へ\nプロデュース機能の内製化\n二次収益(配信/MD/海外)の取り込み\n中小スタジオの元請参入",
  "権利を持つスタジオと単純受託スタジオの収益格差が拡大。資本力のある元請/配信PFがスタジオ買収・出資でIPを囲い込み、業界は『制作＝権利者』への再編が進む。",
  "アニメ制作の権利保有シフトと収益格差","","","配信,国内,トレンド,権利・ライセンス","TDB調査(2026/8)"],
 [TODAY,"アニメ","Netflixが日本アニメのスレートを拡充、京アニ新作など世界独占配信を強化",
  "https://about.netflix.com/ja/news/anime-japan-2026",
  "NetflixがAnimeJapan/アヌシー等で26年配信ラインナップを順次発表。日本アニメを世界配信の主力に据える。",
  "京都アニメーション新作や複数の話題作を世界独占配信として発表。配信PFが制作費と独占権を握り、日本アニメの一次流通が配信に移る流れが強まる。",
  "世界独占配信で日本アニメを囲い込み\n制作費負担と引き換えに独占権\n京アニ等トップスタジオとの直接契約\n地上波→配信の一次窓口移行",
  "配信PFの独占集中で、テレビ局・製作委員会の一次窓口としての地位が低下。スタジオは配信直取引で収益安定を得る一方、IP権利の従属リスクが高まり、権利保持を巡る交渉が焦点化する。",
  "配信独占と一次窓口の移行","Netflix, 京都アニメーション","","配信,海外,トレンド","Netflix(AnimeJapan/アヌシー2026)"],
 [TODAY,"アニメ","安田佳澄『フールナイト』アニメ化決定、Netflixで世界独占配信へ",
  "https://www.oricon.co.jp/news/2432864/full/",
  "集英社のSFヒューマンドラマ漫画がアニメ化。Netflix世界独占配信枠として展開される。",
  "『フールナイト』のアニメ化が決定し、Netflixで2026年に世界独占配信。作家性の強い漫画原作を配信独占アニメで世界に届ける、漫画→アニメの典型的メディアミックス。",
  "漫画原作の配信独占アニメ化\n作家性作品の世界同時展開\n集英社IPの映像化\n原作既刊への販売還流",
  "配信独占アニメ化は原作の海外電子売上を押し上げ、出版社は映像化前提でIPを設計する。ヒット指標が『巻数』より『配信での話題化』に移り、原作選定基準が変わる。",
  "配信前提のIP設計と原作選定","集英社, Netflix","フールナイト","映像化,配信,国内","オリコン(Netflix発表)"],
 [TODAY,"漫画","2026年上半期の出版指標が公表、コミック・電子が牽引し紙の縮小続く",
  "https://shuppankagaku.com/",
  "全国出版協会・出版科学研究所が『季刊 出版指標』2026年夏号で上半期市場を公表した。",
  "紙+電子の出版市場は縮小基調が続く一方、コミックと電子が牽引役。2025年は市場全体1兆5462億円(前年比1.6%減)、コミック占有44.8%・電子は増勢で、紙誌からデジタル/版権への構造転換が鮮明。",
  "紙市場の縮小が継続\n電子/コミックが市場を下支え\nコミックが出版市場の最大セグメント\n版権/海外収入への依存増",
  "出版社の収益源は『紙販売』から『電子+IPライセンス+海外』へ本格移行。ヒット漫画の映像化・ゲーム化を前提にしたIP運用が出版社の中核KPIになり、編集体制も版権部門主導へ再編が進む。",
  "出版のIP化と電子/海外シフト","出版科学研究所","","出版,業績,国内,トレンド","季刊出版指標2026夏号"],
 [TODAY,"漫画","原泰久『キングダム』実写最新作『魂の決戦』が公開、大型漫画IPの実写シリーズ化続く",
  "https://natalie.mu/comic/news",
  "集英社の人気歴史漫画の実写映画シリーズ最新作が公開、初日舞台挨拶を実施した。",
  "『キングダム 魂の決戦』が公開。累計発行部数の大きい漫画IPを実写映画で継続的にシリーズ化する、邦画の大型フランチャイズ戦略の代表例。",
  "大型漫画IPの実写映画シリーズ化\n累計部数を背景にした集客\n出版×映画の共同IP運用\n原作売上への還流",
  "累計部数の大きい漫画は実写映画で複数作フランチャイズ化する型が定着。出版社は映像出資と一体でIPを設計し、実写の海外配信展開も見据えたキャスティング/同時公開が増える。",
  "大型漫画IPの実写フランチャイズ設計","集英社, 東宝","キングダム","映像化,国内,興行","コミックナタリー(2026/8)"],
 [TODAY,"小説","KADOKAWA、ラノベ『なろう・異世界偏重』で出版が営業赤字、構造改革を急ぐ",
  "https://www.itmedia.co.jp/news/articles/2605/15/news107.html",
  "KADOKAWAが出版事業で前年32億円の営業黒字→10億円の赤字に転落。特定ジャンル依存と刊行過多が背景。",
  "『なろう・異世界系』への偏重で市場が飽和。編集者増員で刊行点数を増やしたが、質・新規性が伴わずヒットが出ず、1タイトルあたり部数が減少。ポートフォリオ再編と点数最適化へ動く。",
  "特定ジャンル偏重で市場飽和\n刊行点数増→1タイトル部数減\nヒット創出力の低下\n作品ポートフォリオ再編へ",
  "『点数を絞りIP化前提で厚く育てる』方向へ転換。メディアミックス適性のある作品に投資を集中し、Web発の量産モデルは選別が進む。供給過多→単価下落はゲーム/音楽と共通の構造課題として顕在化。",
  "コンテンツ供給過多と単価下落の構造","KADOKAWA","","出版,業績,国内,トレンド","ITmedia(2026)"],
 [TODAY,"小説","2026年本屋大賞は朝井リョウ『イン・ザ・メガチャーチ』、書店発の話題化が販売を牽引",
  "https://www.hontai.or.jp/history/hontai2026.html",
  "書店員投票の本屋大賞で朝井リョウ作が大賞。翻訳部門は『空、はてしない青』が受賞した。",
  "『イン・ザ・メガチャーチ』(日経BP)が2026年本屋大賞を受賞。文学賞の中でも販売直結力が高い本屋大賞は、受賞を起点にした増刷・映像化提案の呼び水になる。",
  "書店発の賞が販売を牽引\n受賞→増刷/映像化の連鎖\n文芸のIP化入口\n話題化の起点としての賞",
  "本屋大賞級の『販売直結型』文学賞は、受賞を起点に映像化・オーディオ化の権利提案が動く。出版社は受賞候補段階からメディアミックスを仕込み、文芸IPの多角化を狙う。",
  "販売直結型文学賞のIP化機能","日経BP","イン・ザ・メガチャーチ","受賞,出版,国内","本屋大賞2026"],
 [TODAY,"音楽","Billboard JAPAN Hot 100はNumber_iが首位、Mrs. GREEN APPLEが2位",
  "https://www.billboard-japan.com/charts/detail/?a=hot100",
  "直近のBillboard JAPAN総合ソングチャートの上位動向。",
  "総合チャートでNumber_iが首位、Mrs. GREEN APPLEが2位、米津玄師らが上位。複合指標(CD/DL/ストリーミング/動画/ラジオ等)で男性グループとバンドが競る構図。",
  "複合指標での首位争い\n男性グループの動員力\nストリーミング比重の上昇\nCD初動型とサブスク型の併存",
  "複合指標化でCD初動一本足打法が弱まり、ストリーミング/ショート動画の継続再生が順位を左右。アーティストは配信ロングテールとライブ動員を軸にした収益設計へ移行する。",
  "複合チャートと収益構造の変化","","Number_i, Mrs. GREEN APPLE","チャート・ランキング,国内","Billboard JAPAN(2026/8)"],
 [TODAY,"音楽","ライブ・エンタメ市場が2025年8564億円で過去最高、ぴあ総研は35年1兆円予測",
  "https://corporate.pia.jp/news/detail_live_enta_market20260617.html",
  "ぴあ総研が国内ライブ・エンタメ市場の確定値と将来予測を公表した。",
  "2025年の市場規模(音楽コンサート+ステージのチケット販売額)は8564億円で過去最高。推し活/トキ消費、大型会場の新設・高稼働、単価上昇が牽引。2035年に1兆円規模へ拡大と予測。",
  "体験消費(ライブ)が最高更新\nチケット単価の上昇\n大型会場の新設/高稼働\n音源収益からライブ収益への重心移動",
  "音源(配信)は集客の入口、収益の柱はライブ/グッズ/ファンダムへという構図が固定化。ダイナミックプライシングや会員制の高単価施策が広がり、会場・興行インフラへの投資競争が続く。",
  "体験経済とライブ収益の高単価化","ぴあ","","ライブ・イベント,業績,国内,トレンド","ぴあ総研(2026)"],
 [TODAY,"音楽","Warner Music、AI音楽ライセンス契約が2026年後半のサブスク収益を押し上げと予想",
  "https://www.musicbusinessworldwide.com/warner-music-expects-ai-music-deals-to-boost-subscription-streaming-revenue-from-late-2026/",
  "メジャーが生成AI音楽をライセンス収益源として組み込む方針を鮮明化した。",
  "Warner Musicは、AI音楽関連のライセンス契約が2026年後半からサブスク・ストリーミング収益を押し上げると見込む。訴訟対象だった生成AIを『対価を取るライセンス先』に転換する潮流。",
  "AIを訴訟対象からライセンス先へ\n新たなサブスク収益源化\n権利者への対価還元モデル\nメジャー主導の枠組み作り",
  "メジャーはAI生成/加工を有料機能として取り込み、収益の重心が『原盤販売』から『AI利用ライセンス+権利管理』へ移る。独立系/権利管理団体を巻き込んだ包括ライセンスが標準化し、AI音楽の市場化が進む。",
  "生成AI音楽のライセンス市場化","Warner Music Group","","AI,配信,権利・ライセンス,海外","MBW(2026)"],
 [TODAY,"音楽","Spotify×Universalの『AIカバー/リミックス』が有料アドオン化、Merlin・Kobaltにも拡大",
  "https://www.musicbusinessworldwide.com/spotify-and-universal-music-group-strike-landmark-deal-to-let-fans-create-ai-covers-and-remixes-as-a-paid-premium-add-on/",
  "SpotifyがUniversalに続きMerlin/Kobaltともファンによるライセンス済みAIカバー/リミックスで合意した。",
  "SpotifyのAIカバー/リミックス機能はPremium有料アドオンとして提供され、参加はアーティスト/作家のオプトイン。権利者が生成物の価値を直接分配される設計で、独立系にも枠組みが拡大。",
  "AI生成をオプトイン+有料アドオンで収益化\n権利者への直接分配\n独立系(Merlin/Kobalt)へ拡大\nWMG/Sony Musicは未締結",
  "『ファンが作るライセンス済みAI二次創作』が新たな課金レイヤーになり、UGCとメジャー原盤の境界が溶ける。参加率と分配率が焦点で、メジャー全社が揃えばAI二次創作が主要な収益カテゴリに育つ。",
  "ライセンス済みAI二次創作の課金モデル","Spotify, Universal Music Group, Merlin, Kobalt","","AI,配信,権利・ライセンス,海外","MBW(2026)"],
]

added=0; by_genre={}
for row in news:
    if str(row[3]).strip() in existing_urls: continue
    append_row(ws,row); existing_urls.add(str(row[3]).strip()); added+=1
    by_genre[row[1]]=by_genre.get(row[1],0)+1
print("日次ニュース 追加:",added,by_genre)

# ---------- 横断的問い ----------
ws = wb["横断的問い"]
existing_q=set()
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[1]: existing_q.add(str(r[1])[:24])
cross=[
 [TODAY,"配信PF(Netflix)による日本アニメの制作費負担＋世界独占という囲い込みは、ゲームのプラットフォーマー直営スタジオ化と同じ『資本提供と引き換えの権利従属』構造か。制作側の交渉力はどこで反転するか",
  "TDBアニメ制作市場4000億・26年ピークアウト(8/19)＋NetflixアニメスレートAnimeJapan/アヌシー",
  "Netflix, 京都アニメーション","動画配信×ゲーム開発の垂直統合、通信キャリアのコンテンツ内製","オープン","制作キャパ(アニメーター)の希少性が交渉力反転の起点になりうる","横断"],
 [TODAY,"Paramountの米ワーナー買収に象徴されるメディア大手の垂直統合は、通信・広告・小売の『自社コンテンツ囲い込み』にどう波及するか。日本の通信/EC(ソニー・KDDI・楽天等)のIP投資を加速させるか",
  "Paramount SkydanceのWBD買収($110.9B)確定・Netflix撤退",
  "Paramount, Warner Bros. Discovery, Netflix","通信キャリアの動画/IP内製、EC×コンテンツのサブスク囲い込み","オープン","コンテンツ調達コスト増→川上(製作出資)への内製回帰が進む","横断"],
]
addq=0
for row in cross:
    if str(row[1])[:24] in existing_q: continue
    append_row(ws,row); addq+=1
print("横断的問い 追加:",addq)

# ---------- メディアミックス追跡 ----------
ws = wb["メディアミックス追跡"]
existing_mm=set()
for r in ws.iter_rows(min_row=2, values_only=True):
    existing_mm.add((str(r[1]).strip() if r[1] else "", str(r[4]).strip() if r[4] else ""))
mm=[
 [TODAY,"フールナイト","漫画","安田佳澄/集英社","アニメ","TVアニメ化・Netflix世界独占配信","2026-08","2026年内","発表","集英社, Netflix","作家性SF漫画の配信独占アニメ化"],
 [TODAY,"キングダム","漫画","原泰久/集英社","映画","実写映画『魂の決戦』公開","2026","2026-08(公開済)","公開済","集英社, 東宝","実写シリーズ最新作"],
 [TODAY,"ブルーロック","漫画","金城宗幸・ノ村優介/講談社","映画","実写映画化","2026","2026-08-07(公開済)","公開済","講談社, 東宝","アニメ→実写のフルライン展開"],
]
addmm=0
for row in mm:
    k=(row[1],row[4])
    if k in existing_mm: continue
    append_row(ws,row); existing_mm.add(k); addmm+=1
print("メディアミックス 追加:",addmm)

# ---------- イベントカレンダー ----------
ws = wb["イベントカレンダー"]
existing_ev=set()
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[2]: existing_ev.add(str(r[2]).strip())
new_ev=[
 ["2026-08-25","2026-08-30","gamescom 2026 / Opening Night Live","展示会","ケルン(独)","リアル+配信","欧州最大級ゲーム展示会。ONLで新作を世界初公開","https://www.gamescom.global/en/event/gamescom-opening-night-live-2026","開催中",TODAY],
 ["2026-09-17","2026-09-21","東京ゲームショウ2026","展示会","幕張メッセ","リアル+配信","史上初5日間・30周年。759社/51の国と地域が出展","https://gamebiz.jp/news/420741","開催予定",TODAY],
]
adde=0
for row in new_ev:
    if row[2] in existing_ev: continue
    append_row(ws,row); adde+=1
def pdate(v):
    if v is None or v=="": return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    s=str(v).strip()[:10]
    try: return datetime.strptime(s,"%Y-%m-%d").date()
    except: return None
today_d=date(2026,8,25); upd=0
for row in range(2, ws.max_row+1):
    start=pdate(ws.cell(row=row,column=1).value); end=pdate(ws.cell(row=row,column=2).value)
    st=ws.cell(row=row,column=9); newst=None
    if end and end<today_d: newst="終了"
    elif start and end and start<=today_d<=end: newst="開催中"
    elif start and start>today_d: newst="開催予定"
    if newst and st.value!=newst:
        st.value=newst; st.font=FONT; st.alignment=ALIGN
        lc=ws.cell(row=row,column=10); lc.value=TODAY; lc.font=FONT; lc.alignment=ALIGN; upd+=1
print("イベント 追加:",adde,"ステータス更新:",upd)

# ---------- 今後起きそうなこと ----------
ws = wb["今後起きそうなこと"]
existing_fc=set()
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[1]: existing_fc.add(str(r[1])[:24])
fc=[
 [TODAY,"Paramountの米ワーナー買収完了に伴い、WB Games傘下スタジオ(Rocksteady/NetherRealm等)またはDC・ハリポタのゲーム化権が切り出され、事業パブリッシャー(バンナム/セガ/スクエニ)や国家系ファンドが買い手に回る",
  "中","6〜12ヶ月","Paramount, WB Games, Batman/Mortal Kombat, Hogwarts Legacy",
  "Paramount SkydanceのWBD買収確定(Wikipedia/Netflix IR)＋WB Games去就不透明(日次8/25)",
  "WB Gamesスタジオの売却・分社やIPライセンス移管の適時開示/海外報道を観測。半年内に動きが出るか","横断"],
]
addfc=0
for row in fc:
    if str(row[1])[:24] in existing_fc: continue
    append_row(ws,row); addfc+=1
print("今後起きそうなこと 追加:",addfc)

# ---------- ゲームデザイン・トレンド (既存更新) ----------
ws = wb["ゲームデザイン・トレンド"]
gd=0
for row in range(2, ws.max_row+1):
    name=ws.cell(row=row,column=2).value
    if name and "エバーグリーン" in str(name):
        ws.cell(row=row,column=1,value=TODAY)
        c9=ws.cell(row=row,column=9); prev=c9.value or ""
        c9.value=(str(prev)+" / gamescom ONL 2026: Witcher3新DLC『Songs of the Past』(10年前の名作へ有料大型拡張)").strip(" /")
        ws.cell(row=row,column=8,value="旧作への有料大型DLC・準ライブ運営で買い切り大作を長期収益化する型が定着。gamescomでもレガシー作へのDLC投入が続き、新作より既存資産の延命が主流化")
        for cc in (1,8,9):
            cell=ws.cell(row=row,column=cc); cell.font=FONT; cell.alignment=ALIGN
        gd+=1
print("ゲームデザイン 更新:",gd)

wb.save(PATH)
print("SAVED")
