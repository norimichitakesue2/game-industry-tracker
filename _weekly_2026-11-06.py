# -*- coding: utf-8 -*-
import openpyxl
PATH = "master/ゲーム業界・時事情報収集.xlsx"
wb = openpyxl.load_workbook(PATH)

added = {}

def existing_keys(ws, idxs, start=2):
    keys = set()
    for r in ws.iter_rows(min_row=start, values_only=True):
        key = tuple("" if r[i] is None else str(r[i]).strip() for i in idxs)
        keys.add(key)
    return keys

def append_rows(sheet, rows, key_idxs):
    ws = wb[sheet]
    keys = existing_keys(ws, key_idxs)
    n = 0
    for row in rows:
        key = tuple("" if row[i] is None else str(row[i]).strip() for i in key_idxs)
        if key in keys:
            continue
        ws.append(row)
        keys.add(key)
        n += 1
    added[sheet] = n

# ===== 新作リリースログ (13 cols) =====
# 発売日,タイトル,開発,パブリッシャー,プラットフォーム,ジャンル,Metacritic,OpenCritic,Steam同接ピーク,初週推定本数,話題度,感想/メモ,URL
game_rel = [
 ["2026-11-05","ゼルダの伝説 時のオカリナ（The Legend of Zelda: Ocarina of Time）","グレッゾ／任天堂","任天堂","Switch2","アクションアドベンチャー","","","","",5,
  "N64/3DSの名作『時のオカリナ』のSwitch2向けリメイク。ゼルダ40周年Directで発表、11/5発売。amiibo対応。","https://dengekionline.com/article/202609/87215"],
 ["2026-11-03","Zero Parades: For Dead Spies ディレクターズカット（ゼロ・パレード）","ZA/UM","ZA/UM","PS5/Switch2","諜報RPG","","","","",4,
  "『ディスコ エリジウム』開発元ZA/UMの新作諜報RPGのコンソール版（Director's Cut）。PC版は2026/5/21発売済。","https://rpgamer.com/2026/09/zero-parades-for-dead-spies-directors-cut-ps5-release-date-announced/"],
 ["2026-10-30","ハローキティ パーティーランド（Hello Kitty Party Land）","","","Switch2/Switch","パーティー","","","","",2,
  "ファミ通週販(10/26-11/1)でSwitch2版7位・Switch版8位にランクイン。サンリオIPのパーティーゲーム。","https://gematsu.com/2026/11/famitsu-sales-10-26-26-11-1-26"],
 ["2026-11-12(予定)","メタファー：リファンタジオ（Metaphor: ReFantazio）Switch2版","アトラス","セガ／アトラス","Switch2","RPG","","","","",4,
  "【来週】高評価ファンタジーRPGのSwitch2移植。11/12発売。携帯機でもプレイ可能に。","https://www.rpgsite.net/news/19530-metaphor-refantazio-switch-2-release-date-direct"],
]
append_rows("新作リリースログ", game_rel, (0,1))

# ===== インディー注目 (9 cols) =====
# 更新日,タイトル,開発,発売日/状況,プラットフォーム,ジャンル,注目点,Steam実績,メモ
indie = [
 ["2026-11-06","Moon Sweepers（ムーンスイーパーズ）","2人の個人開発者","2026-10 製品版配信中","PC(Steam)","最大16人協力 清掃シム",
  "『PowerWash Simulator』系の月面清掃ゲーム。2人の開発者制作でSteamのAAA大作を上回る人気に。",
  "Steam週間トップセラー8位(前週比+12、10/27-11/3)","バイラルヒット。配信・実況で拡散。要ウォッチ。"],
]
append_rows("インディー注目", indie, (1,))

# ===== 週間ランキング (10 cols) =====
# 集計週,区分,順位,タイトル,プラットフォーム,週販本数,累計本数,Steam同接ピーク,出典,メモ
FAMI_SRC = "https://gematsu.com/2026/11/famitsu-sales-10-26-26-11-1-26"
fami_week = "2026-10-26〜11-01"
fami = [
 [fami_week,"国内ファミ通",1,"Nintendo Switch Sports Resort（Nintendo Switch スポーツ リゾート）","Switch2","23,466","88,098","",FAMI_SRC,"ファミ通週間推定(gematsu集計)"],
 [fami_week,"国内ファミ通",2,"ファイナルファンタジー レゾナンス","Switch2","14,462","66,730","",FAMI_SRC,"ファミ通週間推定(gematsu集計)"],
 [fami_week,"国内ファミ通",3,"Minecraft（マインクラフト）","Switch2","11,754","","",FAMI_SRC,"新作(累計=週販)。Switch2版"],
 [fami_week,"国内ファミ通",4,"ファイナルファンタジー レゾナンス","PS5","11,181","61,633","",FAMI_SRC,"ファミ通週間推定(gematsu集計)"],
 [fami_week,"国内ファミ通",5,"リズム天国 ミラクルスターズ","Switch","8,753","1,111,851","",FAMI_SRC,"ロングセラー継続"],
 [fami_week,"国内ファミ通",6,"ファイナルファンタジー レゾナンス","Switch","8,425","38,896","",FAMI_SRC,"3機種でトップ10入り"],
 [fami_week,"国内ファミ通",7,"ハローキティ パーティーランド","Switch2","7,365","","",FAMI_SRC,"新作。サンリオIP"],
 [fami_week,"国内ファミ通",8,"ハローキティ パーティーランド","Switch","6,252","","",FAMI_SRC,"新作。サンリオIP"],
 [fami_week,"国内ファミ通",9,"ファイアーエムブレム 万紫千紅","Switch2","4,346","186,494","",FAMI_SRC,"ファミ通週間推定(gematsu集計)"],
 [fami_week,"国内ファミ通",10,"Call of Duty: Modern Warfare 4","PS5","4,024","16,840","",FAMI_SRC,"ファミ通週間推定(gematsu集計)"],
]
STEAM_SRC = "https://www.gamegrin.com/news/weekly-top-selling-games-on-steam-27th-of-october-to-3rd-of-november-2026/"
steam_week = "2026-10-27〜11-03"
steam = [
 [steam_week,"Steamトップセラー",1,"Phantom Blade Zero（ファントムブレイドゼロ）","PC","","","",STEAM_SRC,"GameGrin週間トップセラー(順位のみ)。前週3位→1位"],
 [steam_week,"Steamトップセラー",2,"Call of Duty: Modern Warfare 4","PC","","","",STEAM_SRC,"GameGrin週間トップセラー(順位のみ)"],
 [steam_week,"Steamトップセラー",3,"AION 2","PC","","","",STEAM_SRC,"GameGrin週間トップセラー(順位のみ)"],
 [steam_week,"Steamトップセラー",4,"Counter-Strike 2","PC","","","",STEAM_SRC,"GameGrin週間トップセラー(順位のみ)"],
 [steam_week,"Steamトップセラー",5,"Dota 2","PC","","","",STEAM_SRC,"GameGrin週間トップセラー(順位のみ)"],
 [steam_week,"Steamトップセラー",6,"PUBG: BATTLEGROUNDS","PC","","","",STEAM_SRC,"GameGrin週間トップセラー(順位のみ)"],
 [steam_week,"Steamトップセラー",7,"Probably Stolen - Cyberpunk Shopkeeper Sim","PC","","","",STEAM_SRC,"NEWランクイン。インディー店舗経営シム"],
 [steam_week,"Steamトップセラー",8,"Moon Sweepers（ムーンスイーパーズ）","PC","","","",STEAM_SRC,"前週比+12。協力清掃ゲームのバイラルヒット"],
 [steam_week,"Steamトップセラー",9,"Aniimo（アニーモ）","PC","","","",STEAM_SRC,"NEWランクイン。F2Pクリーチャーコレクター"],
 [steam_week,"Steamトップセラー",10,"EA SPORTS FC 27","PC","","","",STEAM_SRC,"GameGrin週間トップセラー(順位のみ)"],
]
append_rows("週間ランキング", fami + steam, (0,1,2))

# ===== 【映画】公開・興行 (13 cols) =====
# 公開日,タイトル,監督,配給,製作国,ジャンル,興収(国内),興収(世界),RT,MC,話題度,メモ,URL
EIGA = "https://eiga.com/news/20261102/21/"
movie = [
 ["2026-10-31","RYUJI 竜二","水田伸生","カルチュア・パブリッシャーズ","日本","ドラマ(リメイク)","週末興行 初登場4位","","","",3,
  "1983年公開作のリメイク。柳楽優弥主演、川栄李奈・萩原利久ら共演。10/30-11/1週末4位スタート。",EIGA],
 ["2026-10-31","クレイフェイス（Clayface）","ジェームズ・ワトキンス","東和ピクチャーズ／東宝","米国","心理スリラー/ホラー","週末興行 初登場9位","","","",3,
  "DCヴィラン「クレイフェイス」の怪物化を描く。ナオミ・アッキーら出演。週末9位初登場。",EIGA],
 ["2026-10-31","ヤマトよ永遠に REBEL3199 第七章 虹色の輪廻","ヤマトナオミチ（総監督:福井晴敏）","松竹ODS事業室","日本","アニメ映画(SF)","週末興行 初登場10位","","","",3,
  "『宇宙戦艦ヤマト』リメイクシリーズ最新章。週末10位初登場。トップは『ママせか』が3週連続首位。",EIGA],
]
append_rows("【映画】公開・興行", movie, (0,1))

# ===== 【アニメ】放送・配信ログ (12 cols) =====
# 開始日,タイトル,原作,制作会社,クール,配信PF,ジャンル,話数,評価,話題度,メモ,URL
anime = [
 ["2026-11-07(予定)","カードファイト!! ヴァンガード Divinez 運命星戦編","ブシロード／TCG","","2026年秋","テレビ放送(全3週予定)","カードバトル","","",2,
  "【新番組】11/7(土)08:00〜放送開始予定、全3週の短期集中シリーズ。","https://gamebiz.jp/news/433270"],
]
append_rows("【アニメ】放送・配信ログ", anime, (0,1))

# ===== 【漫画】新刊・話題作 (11 cols) =====
# 発売日,タイトル,作者,出版社,掲載誌,巻数,初版部数,ランキング,話題度,メモ,URL
BUN = "https://books.bunshun.jp/articles/-/11464"
manga = [
 ["2026-11","満月珈琲店の星詠み 1","望月麻衣(原作)／まさきりょう(作画)","","",1,"","文春週間ベストセラー コミック1位(集計10/26-11/1)",3,
  "人気小説のコミカライズ最新刊が文春週間コミック首位。",BUN],
 ["2026-11","小隊 REVENGE 2","砂川文次／柏葉比呂樹","","",2,"","文春週間ベストセラー コミック2位(集計10/26-11/1)",2,
  "芥川賞作家・砂川文次原作の戦記漫画シリーズ続編。",BUN],
]
append_rows("【漫画】新刊・話題作", manga, (0,1))

# ===== 【音楽】リリース・チャート (11 cols) =====
# 集計日/発売日,区分,タイトル/曲名,アーティスト,レーベル,形態,チャート順位,ストリーミング数,話題度,メモ,URL
BB = "https://www.billboard-japan.com/d_news/detail/167121"
music = [
 ["2026-11-04(付)","チャート","ザッツ！超常","M!LK","ソニー・ミュージック","シングル/配信","Billboard JAPAN Hot 100 1位(2026-11-04付/集計10-26〜11-01)","",4,
  "M!LKが「ザッツ！超常」で初登場総合首位。自身2作目のHot100首位。",BB],
 ["2026-11-04(付)","チャート","しわあわせ","Vaundy","","配信","Billboard JAPAN Hot 100 3位(2026-11-04付)","",4,
  "Vaundy「しわあわせ」が自身初のHot100トップ3入り(3位)。",BB],
 ["2026-11-04(付)","チャート","シンデレラ","Mrs. GREEN APPLE","ユニバーサル","配信","Billboard JAPAN Hot 100 4位(2026-11-04付)","",3,
  "ミセスは「シンデレラ」4位ほか複数曲がトップ10入り。",BB],
]
append_rows("【音楽】リリース・チャート", music, (0,1,2,3))

wb.save(PATH)
print("=== 追加件数 ===")
for k,v in added.items():
    print(f"  {k}: +{v}")
