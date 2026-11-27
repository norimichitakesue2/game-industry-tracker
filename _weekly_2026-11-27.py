# -*- coding: utf-8 -*-
import openpyxl, sys
P = 'master/ゲーム業界・時事情報収集.xlsx'
wb = openpyxl.load_workbook(P)
added = {}

def hdr(ws):
    return [c.value for c in ws[1]]

def existing_keys(ws, key_idx):
    keys = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        keys.add(tuple((str(r[i]).strip() if i < len(r) and r[i] is not None else '') for i in key_idx))
    return keys

def append_rows(sheet, rows, key_idx):
    ws = wb[sheet]
    h = hdr(ws)
    ncol = len(h)
    keys = existing_keys(ws, key_idx)
    n = 0
    for row in rows:
        row = list(row) + [''] * (ncol - len(row))
        row = row[:ncol]
        k = tuple((str(row[i]).strip() if row[i] is not None else '') for i in key_idx)
        if k in keys:
            continue
        ws.append(row)
        keys.add(k)
        n += 1
    added[sheet] = added.get(sheet, 0) + n
    return n

# ===== 新作リリースログ =====
# key: 発売日(0)+タイトル(1)
release_rows = [
    ['2026-11-20','Football Manager 27 Console（フットボールマネージャー27 コンソール）','Sports Interactive','SEGA','PS5/Xbox Series X|S','サッカー経営シム','','','','','3','今週リリース。人気サッカー経営シムのコンソール版が11/20発売（PC/ダウンロード版は11/11）。GTA6発売週の中堅話題作。','https://www.famitsu.com/schedule/ps5/202611'],
    ['2026-11-26','楽園追放 -The Stellar Angel-','','','Switch/PS5','アクション','','','','','2','今週リリース。劇場アニメ『楽園追放 心のレゾナンス』公開に合わせたゲーム化。メディアミックス展開。','https://www.famitsu.com/schedule/switch/202611'],
    ['2026-11-26','Way of the Hunter 2（ウェイ・オブ・ザ・ハンター2）','Nine Rocks Games','THQ Nordic','PS5/Xbox/PC','ハンティングシム','','','','','2','今週リリース。実写系ハンティングシムの続編。ニッチ人気シリーズ。','https://www.famitsu.com/schedule/ps5/202611'],
]
append_rows('新作リリースログ', release_rows, [0,1])

# ===== 週間ランキング =====
# key: 集計週(0)+区分(1)+順位(2)
rank_rows = []
# 国内ファミ通 2026-11-16〜11-22 (famitsu 95209)
fami_url='https://www.famitsu.com/article/202611/95209'
fami=[
 (1,'グランド・セフト・オートVI（Grand Theft Auto VI）','PS5','200759','200759','13年ぶりシリーズ最新作が初登場で圧倒的首位。国内初週推定約20万本（PS5のみ）。'),
 (2,'ゼルダの伝説 時のオカリナ（The Legend of Zelda: Ocarina of Time）','Switch2','35295','658433','前週1位から2位。'),
 (3,'Nintendo Switch Sports Resort','Switch2','10486','126291',''),
 (4,'リズム天国 ミラクルスターズ','Switch','7830','1136253',''),
 (5,'トモダチコレクション わくわく生活','Switch','4570','1654904',''),
 (6,'スプラトゥーン レイダース','Switch2','3077','700755',''),
 (7,'桃太郎電鉄2 ～あなたの町も きっとある～ 東日本編＋西日本編','Switch','2969','390053','Switch2無料体験版(11/17配信)効果で再浮上。'),
 (8,'Minecraft','Switch2','2869','21081',''),
 (9,'メタファー：リファンタジオ（Metaphor: ReFantazio）','Switch2','2426','13593','前週2位から9位に下降。'),
 (10,'桃太郎電鉄2 Nintendo Switch 2 Edition 東日本編＋西日本編','Switch2','2385','331659',''),
]
for rk,title,pf,wk,cum,memo in fami:
    rank_rows.append(['2026-11-16〜11-22','国内ファミ通',rk,title,pf,wk,cum,'',fami_url,('集計11/16-22（ファミ通推定値）。'+memo).strip()])

# Steamトップセラー 2026-11-10〜11-17 (GameGrin, 掲載順)
gg1='https://www.gamegrin.com/news/weekly-top-selling-games-on-steam-10th17th-of-november-2026/'
steam1=['Apex Legends','Euro Truck Simulator 2','Counter-Strike 2','Call of Duty: Modern Warfare 4','Bodycam','AION 2','Dota 2','Phantom Blade Zero','Football Manager 27','Euro Truck Simulator 2 - Coaches']
for i,t in enumerate(steam1,1):
    memo='順位はGameGrin掲載順（Valve公式トップセラー準拠・絶対順位非掲載）。'
    if i==1: memo+='Apexが首位。'
    if t.startswith('Euro Truck'): memo+='ETS2関連が新規ランクイン。'
    rank_rows.append(['2026-11-10〜11-17','Steamトップセラー',i,t,'PC(Steam)','','','',gg1,memo])

# Steamトップセラー 2026-11-17〜11-24 (GameGrin, 掲載順)
gg2='https://www.gamegrin.com/news/weekly-top-selling-games-on-steam-17th24th-of-november-2026/'
steam2=['Counter-Strike 2','Steam Deck','EA SPORTS FC 27','Dota 2','ROB RUN','AION 2','Apex Legends','Call of Duty: Modern Warfare 4','Bodycam','Grand Theft Auto V Enhanced']
for i,t in enumerate(steam2,1):
    memo='順位はGameGrin掲載順（Valve公式トップセラー準拠・絶対順位非掲載）。'
    if i==1: memo+='Counter-Strike 2が首位。'
    if t.startswith('Grand Theft Auto V'): memo+='GTA6発売を受け旧作GTA Vが新規ランクイン。'
    if t=='Steam Deck': memo+='ハード(Steam Deck)が商品として上位に。'
    rank_rows.append(['2026-11-17〜11-24','Steamトップセラー',i,t,'PC(Steam)','','','',gg2,memo])

append_rows('週間ランキング', rank_rows, [0,1,2])

# ===== 映画 公開・興行 =====
# key: 公開日(0)+タイトル(1)
eiga_url='https://eiga.com/ranking/jp/'
movie_rows=[
 ['2026-11-20','魔法使いの夜（まほうつかいのよる）','','東宝／アニプレックス','日本','アニメ映画','週末興行 初登場2位（動員25.6万人・興収4.51億円/3日間）','','','','4','TYPE-MOON原作の劇場アニメ。11月下旬公開で週末興行初登場2位、3日間で興収4.5億円。ゴジラ-0.0に次ぐ好スタート。',eiga_url],
 ['2026-11-20','装甲騎兵ボトムズ 灰色の魔女（ヘクセ）','','バンダイナムコフィルムワークス','日本','アニメ映画(SF/ロボット)','週末興行 初登場3位（動員9.6万人・興収1.65億円）','','','','3','『装甲騎兵ボトムズ』シリーズの劇場新作。初登場3位。',eiga_url],
 ['2026-11-20','ないものねだりの君に光の花束を','','ショウゲート','日本','青春/ドラマ','週末興行 初登場7位','','','','2','原菜乃華・作間龍斗(ACEes)主演の青春映画。初登場7位。',eiga_url],
]
append_rows('【映画】公開・興行', movie_rows, [0,1])

# ===== アニメ 放送・配信ログ =====
# key: 開始日(0)+タイトル(1)
anime_rows=[
 ['2026-11-26','フールナイト（Fool Night）','安田佳澄（小学館）','','2026年秋','Netflix(世界独占)','SFサスペンス/ヒューマンドラマ','全12話一挙','','3','安田佳澄のSFサスペンス漫画が原作。11/26よりNetflixで全12話を世界独占一挙配信。','https://animeanime.jp/'],
]
append_rows('【アニメ】放送・配信ログ', anime_rows, [0,1])

# ===== 漫画 新刊・話題作 =====
# key: 発売日(0)+タイトル(1)
bunshun_url='https://books.bunshun.jp/articles/-/11530'
manga_rows=[
 ['2026-11','死体は語る 1','上野正彦(原作)／小針かわず(作画)','','','1','','文春週間ベストセラー コミック4位(集計11/16-22)','2','法医学エッセイ『死体は語る』のコミカライズ。2巻に続き1巻もトップ5入り（11/16-22集計で4位）。',bunshun_url],
]
append_rows('【漫画】新刊・話題作', manga_rows, [0,1])

# ===== 小説 新刊・受賞 =====
# key: 発売日発表日(0)+タイトル(1)
novel_rows=[
 ['2026-11-26','台湾漫遊録（臺灣漫遊錄 / Taiwan Travelogue）','楊双子（ヤン・ソウシ）／Lin King訳','','','受賞','英Waterstones 2026年ブック・オブ・ザ・イヤー（小説部門）','','3','英最大書店チェーンWaterstonesの2026年年間最優秀書・小説部門に台湾の楊双子『台湾漫遊録』（英訳版）が選出。翻訳文学の評価高まる。','https://www.thebookseller.com/'],
]
append_rows('【小説】新刊・受賞', novel_rows, [0,1])

# ===== 音楽 リリース・チャート =====
# key: 集計日発売日(0)+区分(1)+タイトル曲名(2)
bbj='https://www.billboard-japan.com/charts/detail?a=hot100'
music_rows=[
 ['2026-11-25(付)','チャート','IGNITE','JO1','','配信/シングル','Billboard JAPAN Hot 100 1位','','4','JO1「IGNITE」が前週33位から急上昇し総合首位を獲得。',bbj],
 ['2026-11','チャート',"Choosin' Texas",'Ella Langley','','配信','Billboard Hot 100(米) 1位','','3','Ella Langleyのカントリー曲が米Billboard Hot 100で通算25週超の歴代級ロングラン首位を継続。','https://www.billboard.com/charts/hot-100/'],
 ['2026-11-25(付)','新譜/チャート','僕は味方だ','僕が見たかった青空','','シングル','Billboard JAPAN Hot 100 初登場5位','','3','「僕が見たかった青空」の新曲が初登場5位。今週の新規曲で最高位。',bbj],
]
append_rows('【音楽】リリース・チャート', music_rows, [0,2,2] if False else [0,1,2])

# ===== 作品/アーティストウォッチ (任意・各0〜2) =====
# アニメ作品ウォッチ key: 作品名(0)
aw=[['幼稚園WARS','千葉ヨウ（集英社/少年ジャンプ+）','サンライズ','2027年4月開始予定','2027春','新規ウォッチ','2026/11発表。2027年4月放送開始、サンライズ制作、OP等情報公開。','ジャンプ+発のアクション作。','2026-11-27']]
append_rows('【アニメ】作品ウォッチ', aw, [0])
# 音楽アーティストウォッチ key: アーティスト名(0)
mw=[['JO1','LAPONE','LAPONE ENTERTAINMENT','J-POP/ボーイズグループ','日本','新規ウォッチ','2026/11 「IGNITE」でBillboard JAPAN Hot 100総合首位。','日本発グローバルボーイズグループ。','2026-11-27']]
append_rows('【音楽】アーティストウォッチ', mw, [0])

wb.save(P)
print('ADDED per sheet:')
for k,v in added.items():
    print(f'  {k}: {v}')
print('TOTAL:', sum(added.values()))
