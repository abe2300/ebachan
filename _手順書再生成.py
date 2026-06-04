# -*- coding: utf-8 -*-
"""定例報告アプリ 運用手順書 (.docx) 生成スクリプト"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(THIS_DIR, "定例報告_運用手順書.docx")

# テーマカラー（紺色ベース）
THEME_COLOR = "1E3A8A"
ACCENT_COLOR = "F59E0B"

doc = Document()

# ----- 既定スタイル -----
style = doc.styles["Normal"]
style.font.name = "游ゴシック"
style.font.size = Pt(10.5)
rPr = style.element.get_or_add_rPr()
rFonts = rPr.find(qn("w:rFonts"))
if rFonts is None:
    rFonts = OxmlElement("w:rFonts")
    rPr.append(rFonts)
rFonts.set(qn("w:eastAsia"), "游ゴシック")
rFonts.set(qn("w:ascii"), "游ゴシック")
rFonts.set(qn("w:hAnsi"), "游ゴシック")

# ページ余白
for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)


def set_jp_font(run, name="游ゴシック", size=None, bold=None, color=None, mono=False):
    if mono:
        name = "Consolas"
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name if not mono else "ＭＳ ゴシック")
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def add_h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    set_jp_font(r, size=18, bold=True, color=RGBColor(0x1E, 0x3A, 0x8A))
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), THEME_COLOR)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def add_h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_jp_font(r, size=13.5, bold=True, color=RGBColor(0x33, 0x33, 0x33))
    return p


def add_h3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("■ " + text)
    set_jp_font(r, size=11.5, bold=True, color=RGBColor(0x1E, 0x3A, 0x8A))
    return p


def add_p(text, bold=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_jp_font(r, size=10.5, bold=bold, color=color)
    return p


def add_bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    set_jp_font(r, size=10.5)
    return p


def add_number(text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    set_jp_font(r, size=10.5)
    return p


def add_note(text, color="FFF8E1", border="F59E0B"):
    """注意ボックス"""
    t = doc.add_table(rows=1, cols=1)
    t.autofit = True
    cell = t.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "8")
        b.set(qn("w:color"), border)
        tcBorders.append(b)
    tcPr.append(tcBorders)
    p = cell.paragraphs[0]
    for line in text.split("\n"):
        if p.runs:
            p = cell.add_paragraph()
        r = p.add_run(line)
        set_jp_font(r, size=10)
    doc.add_paragraph()


def add_table(headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Light Grid Accent 1"
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        p = cell.paragraphs[0]
        r = p.add_run(h)
        set_jp_font(r, size=10.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), THEME_COLOR)
        tcPr.append(shd)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.rows[ri + 1].cells[ci]
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            set_jp_font(r, size=10)
    doc.add_paragraph()


# =============================================
# 表紙
# =============================================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(80)
r = p.add_run("定例報告アプリ")
set_jp_font(r, size=28, bold=True, color=RGBColor(0x1E, 0x3A, 0x8A))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("運用手順書")
set_jp_font(r, size=20, bold=True, color=RGBColor(0x33, 0x33, 0x33))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(40)
r = p.add_run("毎日の経済ニュース・税制改正情報を確認するPWAアプリ")
set_jp_font(r, size=12, color=RGBColor(0x66, 0x66, 0x66))

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(200)
r = p.add_run("作成日：2026年6月2日")
set_jp_font(r, size=11, color=RGBColor(0x99, 0x99, 0x99))

doc.add_page_break()

# =============================================
# 1. アプリの概要
# =============================================
add_h1("1. アプリの概要")
add_p("「定例報告アプリ」は、毎日の経済ニュース・税制改正情報を一覧できるPWA（プログレッシブウェブアプリ）です。PC・スマホ・タブレットから同じ画面を確認でき、複数人での情報共有も可能です。")

add_h3("主な機能")
add_bullet("主要経済ニュース 5件（各ニュースに株価への影響コメント付き）")
add_bullet("税制改正情報（会社員関連項目は詳細解説モーダル付き）")
add_bullet("税制改正の年度切替機能（過去年度の参照可、自動年度判定）")
add_bullet("日付の自動更新（今日の日付・報告データ更新日）")
add_bullet("印刷／PDF出力ボタン")
add_bullet("全文コピー機能")
add_bullet("報告ルール・情報ソースの参照")

add_h3("特徴")
add_bullet("ローカル起動：自分のPCで完結（オフライン対応）")
add_bullet("Web公開：Netlify経由で携帯・他人とも共有可")
add_bullet("PWA対応：スマホのホーム画面に追加してアプリのように使える")

# =============================================
# 2. フォルダ構成
# =============================================
add_h1("2. フォルダ構成")
add_p("「クロード用 > 定例報告」フォルダに、以下のファイルが含まれています。")

add_table(
    ["ファイル名", "役割"],
    [
        ["index.html", "メイン画面（HTML本体）"],
        ["manifest.json", "PWAマニフェスト（アプリ情報）"],
        ["sw.js", "Service Worker（オフライン対応）"],
        ["icon-192.png / icon-512.png", "アプリアイコン"],
        ["server.py", "ローカルHTTPサーバー（ポート8767）"],
        ["定例報告を起動.bat", "ローカル起動（ワンクリック）"],
        ["定例報告サーバー停止.bat", "サーバー停止"],
        ["公開用を更新.bat", "Netlify公開用ファイルの更新"],
        ["公開用/", "Netlifyにアップロードするフォルダ"],
        ["_手順書再生成.py", "本手順書を再生成するスクリプト"],
    ],
)

# =============================================
# 3. ローカルで起動する方法
# =============================================
add_h1("3. ローカルで起動する方法（自分のPCで使う）")

add_h2("起動手順")
add_number("「定例報告」フォルダを開く")
add_number("「定例報告を起動.bat」をダブルクリック")
add_number("Microsoft Edge がアプリモード（URLバーなし）で自動起動")
add_number("画面に主要経済ニュース・税制改正情報が表示される")

add_note("💡 ヒント：BATファイルを右クリック→「ショートカットの作成」でデスクトップに置いておくと、毎朝ワンクリックで起動できます。")

add_h2("停止手順")
add_p("基本的にはブラウザを閉じるだけでOKです。サーバーを完全停止したい場合：")
add_number("「定例報告サーバー停止.bat」をダブルクリック")
add_number("バックグラウンドで動いているPythonサーバーが停止")

add_h2("起動の仕組み")
add_bullet("BATファイルが Python の HTTPサーバー（pythonw）をバックグラウンド起動")
add_bullet("サーバーは 127.0.0.1:8767 でローカル待機")
add_bullet("Microsoft Edge をアプリモードで開き、http://localhost:8767/index.html を表示")

# =============================================
# 4. 携帯・他人と共有する方法
# =============================================
add_h1("4. 携帯・他人と共有する方法（Netlifyで公開）")
add_p("Netlifyという無料サービスを使って、誰でもURLからアクセスできる公開サイトを作ります。PayPay運用管理と同じ仕組みです。")

add_h2("初回セットアップ（最初の1回だけ）")
add_number("ブラウザで https://app.netlify.com/ にアクセス")
add_number("PayPay運用管理用のアカウントでログイン（または新規作成）")
add_number("「Add new site」→「Deploy manually」をクリック")
add_number("「定例報告 > 公開用」フォルダをページにドラッグ＆ドロップ")
add_number("自動でURLが発行される（例：https://teirei-houkoku-xxx.netlify.app）")
add_number("発行されたURLをメモ・ブックマークしておく")

add_note("⚠️ 注意：URLは複雑な文字列で発行されます。Netlifyの「Site settings」→「Change site name」で覚えやすい名前に変更できます。")

add_h2("内容を更新する手順（毎回）")
add_number("「公開用を更新.bat」をダブルクリック")
add_number("最新の index.html などが「公開用」フォルダにコピーされる")
add_number("ブラウザで Netlify ダッシュボードを開く")
add_number("対象サイトを選択→「Deploys」タブ→「Production deploys」エリアへ")
add_number("「公開用」フォルダをドラッグ＆ドロップ")
add_number("数秒〜数十秒でサイトが更新される")

add_h2("共有方法")
add_bullet("発行されたURLをLINE・メール・Teamsなどで送信")
add_bullet("受け取った人はブラウザでURLを開くだけ")
add_bullet("スマホでアクセスして「ホーム画面に追加」すればアプリ風に使える")

add_h2("制限・セキュリティ")
add_bullet("Netlify無料枠：月間100GB帯域、十分すぎる容量")
add_bullet("URLを知っている人全員がアクセス可能")
add_bullet("非公開にしたい場合：Netlify有料プランでパスワード保護可能")

# =============================================
# 5. スマホで使う方法
# =============================================
add_h1("5. スマホで使う方法（PWA化）")
add_p("公開されたURLにスマホでアクセスすると、アプリのようにホーム画面に追加できます。")

add_h2("iPhone (Safari)")
add_number("Safari で Netlify の公開URLを開く")
add_number("画面下の「共有」ボタン（□に↑のマーク）をタップ")
add_number("メニューから「ホーム画面に追加」を選択")
add_number("名前を確認して「追加」をタップ")
add_number("ホーム画面に「定例報告」のアイコンが追加される")

add_h2("Android (Chrome)")
add_number("Chromeで Netlify の公開URLを開く")
add_number("右上の「︙」メニューをタップ")
add_number("「ホーム画面に追加」を選択")
add_number("名前を確認して「追加」をタップ")

add_note("💡 PWA化のメリット：オフラインでも閲覧可、起動が早い、ブラウザのURLバーが表示されない（アプリ風）")

# =============================================
# 6. 日々の運用フロー
# =============================================
add_h1("6. 日々の運用フロー")

add_h2("6-1. 毎朝の流れ")
add_number("7:00 Claudeに「定例報告」と依頼")
add_number("Claudeが最新の経済ニュース5件＋税制改正情報を生成")
add_number("内容を確認し、index.html に反映してもらう")
add_number("「公開用を更新.bat」を実行→Netlifyにドラッグ＆ドロップ")
add_number("全員に最新情報が共有される")

add_h2("6-2. ニュース項目を変更する場合")
add_bullet("Claudeに「○○のニュースに差し替えて」と依頼")
add_bullet("Claudeが index.html を直接編集")
add_bullet("更新後、「公開用を更新.bat」→Netlify公開")

add_h2("6-3. 日付表示の見方")
add_p("画面ヘッダーに2種類の日付が表示されます。役割を理解しておくと、いつのデータかが分かりやすくなります。")
add_table(
    ["表示項目", "意味", "更新方法"],
    [
        ["大きな日付（例：2026年6月4日）", "アプリを開いた日（PCの今日の日付）", "毎日自動更新"],
        ["📌 報告データ更新日", "ニュース・税制情報の最終更新日", "Claudeに依頼して更新時のみ変更"],
        ["フッターの報告日", "アプリを開いた日（PCの今日の日付）", "毎日自動更新"],
    ],
)
add_note("💡 大きな日付と「報告データ更新日」が違う場合 → ニュース内容が古い可能性があります。Claudeに最新版への更新を依頼してください。")

add_h2("6-4. 税制改正の年度切替について")
add_p("税制改正情報セクションには、過去年度のデータも保管されており、ドロップダウンから選択できます。")

add_h3("年度バッジの自動判定")
add_table(
    ["バッジ", "色", "意味"],
    [
        ["最新", "緑", "現在の年度（4月～翌3月）— デフォルト表示"],
        ["前年度", "グレー", "1年前の年度"],
        ["過去", "薄グレー", "2年以上前の年度"],
        ["予定", "青", "未来の改正（早期発表時）"],
        ["更新待ち", "オレンジ", "今年度のデータが未登録（要更新）"],
    ],
)
add_note("4月1日になると、自動で「最新」バッジが次の年度に切り替わります。手動操作は不要です。")

add_h3("過去年度の参照方法")
add_number("税制改正情報セクションのドロップダウンをクリック")
add_number("見たい年度を選択（令和8年度 / 令和7年度 / 令和6年度 など）")
add_number("テーブルが瞬時に切り替わる")
add_number("詳細ボタンがある項目はクリックでモーダル表示")

add_h2("6-5. 新年度データを追加する方法")
add_p("毎年12月～1月頃に税制改正大綱が発表されるタイミングで、新年度のデータを追加してください。")

add_h3("Claudeへの依頼方法")
add_bullet("例：「令和9年度（2027年度）の税制改正データを追加して」")
add_bullet("Claudeが index.html の taxReformData に自動追加")
add_bullet("ドロップダウンに自動で表示される")
add_bullet("バッジは自動判定（西暦年から計算）なので手動設定不要")

add_h3("追加後の確認")
add_number("アプリを開き直す（または F5 で再読込）")
add_number("新年度がドロップダウンの最上段に表示されているか確認")
add_number("「最新」バッジが付いているか確認")
add_number("「公開用を更新.bat」→ Netlifyに公開")

add_h2("6-6. データ未追加の警告が出たら")
add_p("4月以降、「⚠️ 令和X年度のデータ未追加」というオレンジの警告が表示されることがあります。")
add_bullet("原因：新年度に入ったが、まだ新年度のデータが追加されていない")
add_bullet("対処：Claudeに「最新年度の税制改正データを追加して」と依頼")
add_bullet("追加後は警告が自動で消えます")

# =============================================
# 7. トラブルシューティング
# =============================================
add_h1("7. トラブルシューティング")

add_table(
    ["症状", "原因と対処"],
    [
        ["BATを実行しても起動しない", "Pythonがインストールされていない。python.org からインストール。"],
        ["ポート8767が使用中エラー", "「定例報告サーバー停止.bat」を実行してから再起動。"],
        ["Edge が起動しない", "別のブラウザ（Chrome）でも可。手動で http://localhost:8767/index.html を開く。"],
        ["Netlifyで更新が反映されない", "ブラウザのキャッシュをクリア（Ctrl+Shift+R）。sw.jsのバージョンを上げる手も有効。"],
        ["スマホPWAで古い画面が表示される", "Service Worker のキャッシュ。sw.jsのCACHE名を v2→v3 などに変更してデプロイ。"],
        ["PWAアイコンが古い", "スマホのキャッシュを削除→再度ホーム画面に追加。"],
        ["日付が今日に更新されない", "JSが動いていない可能性。F5で再読込。年が変わっても自動更新されます。"],
        ["税制改正が古い年度のまま", "4月1日に自動切替されない場合、ブラウザのリロード（Ctrl+Shift+R）を試す。"],
        ["「データ未追加」警告が出る", "新年度のデータが未登録。Claudeに「最新年度の税制改正データを追加して」と依頼。"],
        ["手順書を作り直したい", "「_手順書再生成.py」をダブルクリックで再生成。"],
    ],
)

# =============================================
# 8. 手順書の再生成
# =============================================
add_h1("8. 手順書の再生成")
add_p("本手順書は Python スクリプトで自動生成しています。内容を変更したい場合：")
add_number("「_手順書再生成.py」をテキストエディタで開く")
add_number("修正したい箇所を編集")
add_number("ファイルをダブルクリックで実行")
add_number("「定例報告_運用手順書.docx」が上書き生成される")

add_note("必要なPythonライブラリ：python-docx\nインストール：pip install python-docx")

# =============================================
# 保存
# =============================================
doc.save(OUT_PATH)
print(f"手順書を生成しました：{OUT_PATH}")
