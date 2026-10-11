# 川畑水産（八重山）Web受注・出荷管理システム 引き継ぎ書【人間・管理者用】

本書は、川畑水産における**Web注文フォーム・生体リクエストフォーム・現場出荷ピッキング管理システム**の全体構成、運用手順、設定変更、トラブルシューティングをまとめた後任者・管理者向けの引き継ぎ資料です。

---

## 1. システム全体概要

本システムは、**LINE公式アカウント**と**Googleスプレッドシート**を中核とし、顧客が簡単にスマートフォンから注文でき、川畑水産の現場スタッフがスムーズに出荷ピッキング・発送管理を行えるよう設計されたサーバーレスシステムです。

### 構成図
```
【一般のお客様】
       │
       ▼
 📱 公式LINE（@721pyqfn）
       │
       ├─ [Web注文する] ──► 🛒 商品注文フォーム (index.html)
       │                         │ (自動顧客認証 / 在庫連動 / 受取不可日)
       │                         ▼
       └─ [生体リクエスト] ─► 🎣 要望フォーム (request.html)
                                 │
                                 ▼
                   ┌─────────────────────────────┐
                   │  Google Apps Script (GAS)   │
                   │ (注文受付 / 納品書PDF自動生成) │
                   └──────────────┬──────────────┘
                                  │
                                  ▼
                   📊 Googleスプレッドシート (本番)
                     (顧客マスタ / 在庫 / 注文明細)
                                  ▲
                                  │
【川畑水産 現場スタッフ】          │ (リアルタイム同期 / 出荷済チェック / 在庫増減)
       │                          │
       ├─► 📦 出荷ピッキング (picking.html - 社内専用)
       │     ※ 未発送のみ集約、電話追加登録
       │
       └─► 🐠 在庫・生体管理 (stock.html - 社内専用)
             ※ スマホで在庫+-、カメラ写真即時Drive保存、新商品登録
```

---

## 2. 公開サイト・管理URL一覧

| ページ名 | 公開URL | 対象者 | 役割・用途 |
| :--- | :--- | :--- | :--- |
| **商品注文フォーム** | [https://kawabatasuisanyaeyama.github.io/](https://kawabatasuisanyaeyama.github.io/) | 一般顧客 | 生体・擬岩等のWeb注文（LINEログイン自動認証） |
| **生体リクエスト** | [https://kawabatasuisanyaeyama.github.io/request.html](https://kawabatasuisanyaeyama.github.io/request.html) | 一般顧客 | 未入荷生体や特定種の採集リクエスト受付 |
| **出荷ピッキング管理** | [https://kawabatasuisanyaeyama.github.io/picking.html](https://kawabatasuisanyaeyama.github.io/picking.html) | **社内専用** | 現場での出荷ピッキング、未発送管理、電話追加登録 |
| **在庫・生体管理** | [https://kawabatasuisanyaeyama.github.io/stock.html](https://kawabatasuisanyaeyama.github.io/stock.html) | **社内専用** | 在庫数ワンタップ増減、カメラ写真/ファイル添付(Drive自動連携)、新規商品登録 |
| **社内説明用スライド** | [https://kawabatasuisanyaeyama.github.io/slides.html](https://kawabatasuisanyaeyama.github.io/slides.html) | **社内専用** | 社長・役員へのシステム概要説明・プレゼン用スライド（PC全画面対応・PDF印刷対応） |
| **スプレッドシート(本番)** | [スプレッドシートを開く](https://docs.google.com/spreadsheets/d/1ZiNel8TGJ-e7JMolkRagVOSnr1UczvQzpTa-F-s4p3o/edit) | 管理者 | 在庫マスタ、顧客マスタ、全注文履歴・明細の原本 |
| **Google Drive(画像)** | `川畑水産_観賞魚画像` (Google Driveフォルダ) | 管理者 | 新規生体登録時にアップロードされた写真の自動格納庫 |
| **Apps Script(GAS)** | [GASエディタを開く](https://script.google.com/u/0/home/projects/1tIl_FYj2PWaSfsvajNXkBgZdgqr4-NMtb9yZsbWPQmSvdfCVQR3n6lv0/edit) | 開発/保守 | バックエンド処理（API、PDF生成、通知、画像保存） |
| **GitHubリポジトリ** | `kawabatasuisanyaeyama/kawabatasuisanyaeyama.github.io` | 開発/保守 | フロントエンド（HTML/CSS/JS）のソースコード管理 |

---

## 3. アカウント・各種設定パラメーター一覧

引き継ぎに必要な各種サービス・ID・キーの一覧です。

### ① Google関連
* **本番オーナーアカウント**: `kawabatasuisam@gmail.com`
* **作業・開発アカウント**: `katoakira1114@gmail.com`（編集者権限付与済み）
* **スプレッドシートID**: `1ZiNel8TGJ-e7JMolkRagVOSnr1UczvQzpTa-F-s4p3o`
* **スプレッドシート名**: `「石垣島観賞魚販売管理」本番`
* **Apps Script スクリプトID**: `1tIl_FYj2PWaSfsvajNXkBgZdgqr4-NMtb9yZsbWPQmSvdfCVQR3n6lv0`
* **GAS Web App デプロイURL**:
  `https://script.google.com/macros/s/AKfycbxIYsPdreUjaOFutfshhQIwOLb_Ke1JprcDZj_LAT9oZnc-l-3jZLsPAsFd_QdQJAQ4eA/exec`

### ② LINE関連
* **公式LINE ID**: `@721pyqfn`
* **公式LINEチャット起動URL**: `https://line.me/R/oaMessage/@721pyqfn/?...`
* **LINE Developers チャネルID**: `2011871129`（チャネル名: `isigaki 注文フォーム`）
* **LIFF ID**: `2011871129-P4IflaGc`（※文字の大文字「I」に注意）
* **LIFF エンドポイントURL**: `https://kawabatasuisanyaeyama.github.io/`

### ③ GitHub関連
* **GitHubアカウント**: `kawabatasuisanyaeyama`
* **公開リポジトリ**: `kawabatasuisanyaeyama.github.io`
* **ソースコード管理元**: `katoakira1114-png/-.git`（ローカル `shop-project`）

---

## 4. 各機能の解説と日常の運用フロー

### A. 注文受付から出荷までの流れ
1. **お客様が注文送信**
   - お客様が公式LINEのメニュー等から注文フォームを開きます。
   - LINEアカウント情報をもとにスプレッドシートの「顧客マスタ」と自動照合され、店舗名・電話番号・個別掛率が自動適用されます。
   - カートで商品を選び、受取不可日を選択して注文を送信すると、お客様のLINEアプリに控えメッセージが転送されます。
2. **スプレッドシートへの自動記帳＆PDF作成**
   - GASがスプレッドシートの「注文明細」シートへ注文データを追記します。
   - Google Drive内に「納品書PDF」が自動作成され、リンクがスプレッドシートに保存されます。
   - 該当カテゴリ（魚、サンゴ等）の在庫数が自動で減算されます。
3. **現場での出荷ピッキング作業**
   - スタッフは現場スマホで [出荷ピッキング管理画面](https://kawabatasuisanyaeyama.github.io/picking.html) を開きます。
   - **未発送の注文のみが自動表示**されます。
   - **「生体集約・合計」タブ**: 本日発送する生体の合計匹数（例: カクレクマノミ 計5匹）が集約されるため、水槽からまとめて掬う作業が迅速に行えます。
   - 各品目をタップすると打ち消し線が引かれ、ピッキング済みをチェックできます。
4. **電話での追加注文・生体追加**
   - お客様から「やっぱりあの魚も1匹追加で！」と電話があった場合：
     - 注文カードの **「➕ 生体を追加（電話追加など）」** をタップ。
     - 生体名を入力するとスプレッドシートの在庫から候補が表示されます（タップで確定）。
     - スプレッドシートにない生体の場合は一番下の「手入力」をタップして直接入力。
     - 匹数を入力（**在庫0やマイナスでも入力可能**）。
     - 「注文に追加する」を押すと、即座にスプレッドシートとピッキング一覧に反映されます。
   - 新規の電話注文の場合は、上部バーの **「📞 新規注文」** から直接作成可能です。
5. **発送完了とLINE通知**
   - 梱包が完了したら、カード内の **「📤 発送LINE」** をタップすると、お客様宛ての発送完了メッセージ（納品書PDFリンク付き）がLINEアプリに立ち上がり、そのまま1タップで送信できます。
   - **「📦 発送完了にする」** をタップすると、画面上から非表示になり、スプレッドシートの「発送済チェック」が自動でONになります。

### B. 在庫管理・写真撮影・新規商品登録の運用フロー（stock.html）
スプレッドシートを直接開かずに、スマホから現場作業中にサクサク在庫を調整・登録できます。

1. **在庫一覧・直接手入力＆クイック増減**
   - [在庫・生体管理コンソール](https://kawabatasuisanyaeyama.github.io/stock.html) にアクセスします。
   - 上部のカテゴリタブや検索バーで目的の生体を絞り込みます。
   - **在庫数の直接手入力**: 550匹など大量採集・まとまった入荷の生体は、カード内の在庫入力枠に数値を直接入力して **「確定」** ボタン（またはEnterキー）を押すだけで、即座にスプレッドシートへ反映されます。
   - **ワンタップ増減**: 日常の微調整には **「-5」「-1」「+1」「+5」** ボタンで素早く増減できます。
   - **「売切」** ボタンを押すと在庫が0になり「売り切れ」に設定されます。入荷時は **「補充+10」** で即再販できます。
   - **「詳細編集」** からは、サイズ規格・単価・ASK切り替え・容量volの修正も可能です。

2. **新規生体・商品の登録**
   - 上部タブの **「➕ 新規商品登録」** をタップします。
   - カテゴリ、限定公開区分（通常商品 / A限定 / B限定 / C限定）、生体名、サイズ、初期在庫、単価（またはASK）、容量volを入力します。
   - 商品IDは自動で規則に従い発番されます（例: `魚_1791730000000`、限定時は `A_魚_1791730000000`）。

3. **生体写真の複数枚添付・自動Google Drive保存連携**
   - **複数枚の写真を同時登録可能**: 「生体写真・画像添付」エリアから、スマホのカメラ撮影や写真ライブラリ・ファイル選択で**複数枚の画像を一度に選択可能**です（PCでのファイル選択・クリップボード貼り付けにも対応）。
   - 選択した写真はサムネイル一覧でプレビューされ、誤った写真は個別に削除（✕）できます。
   - 撮影・選択した全画像は自動で圧縮最適化され、**Googleドライブ内の「川畑水産_観賞魚画像」フォルダへ自動保存**されます。
   - スプレッドシートには各画像の表示URLがカンマ区切りで自動記録されます。
   - **社内での確認**: 在庫一覧カードの `📷 X枚` バッジや写真をタップすると、拡大ギャラリーモーダルが開き、登録されている全写真を矢印ボタンやスワイプで閲覧できます。
   - **お客様側の表示**: [商品注文フォーム (index.html)](https://kawabatasuisanyaeyama.github.io/) の商品カードにも `📷 X枚` バッジが表示され、お客様が写真をタップすると、スワイプや矢印ボタン・サムネイルで様々な角度からの生体写真をじっくり確認できます。

---

## 5. スプレッドシートの構造と編集方法

スプレッドシート `「石垣島観賞魚販売管理」本番` には以下のシート（タブ）があります。

| シート名 | 用途・管理内容 | 編集時の注意点 |
| :--- | :--- | :--- |
| **顧客マスタ** | 登録顧客の情報（LINE UID、店舗名、電話番号、掛率、カテゴリ許可等） | 1行目のヘッダー名は変更しないでください。新規登録申請があると自動追記されます。 |
| **魚 / サンゴ / 貝...** | 各カテゴリごとの商品マスタ（商品名、規格、単価、容量vol、在庫数、画像URL） | 在庫や新商品は [在庫管理画面](https://kawabatasuisanyaeyama.github.io/stock.html) からの編集を推奨します。直接シートを編集する場合は列の順序を変えないでください。 |
| **注文明細** | 受注データの全履歴 | ピッキング画面と双方向連動しています。K列が「発送済」チェックボックスです。 |
| **納品書テンプレート**| PDF納品書を作成するための雛形レイアウト | セルの位置を変更するとPDFの自動出力レイアウトが崩れるため変更時は注意。 |

---

## 6. よくある設定変更・メンテナンス手順

### ① 商品を追加したい / 在庫数を更新したい
- **推奨**: 専用の [在庫・生体管理画面 (stock.html)](https://kawabatasuisanyaeyama.github.io/stock.html) を使います。スマホカメラでの写真添付やワンタップでの増減、自動ID発番に対応しており、ミスなく安全に更新できます。
- **スプレッドシートから行う場合**:
  1. スプレッドシートを開き、該当カテゴリのシート（例:「魚」）を選択します。
  2. 最下行に新しい商品の行を追加（商品名、サイズ規格、単価、容量vol、在庫数）するか、既存商品の「在庫」セルを書き換えます。
  3. **ブラウザ側で約1分後に自動反映**されます。

### ② 注文フォームの案内文や規約を変更したい
1. ローカルの `index.html` または `request.html` の該当テキストを編集します。
2. 後述の「GitHub Pagesへの反映手順」を実施します。

### ③ バックエンド（GAS: 無題.js）のプログラムを修正・反映したい
本環境には **clasp（Google公式CLIツール）** が導入されているため、コマンド一発で本番反映できます。

1. ターミナル（PowerShell）でプロジェクトフォルダを開きます：
   ```powershell
   cd c:\Users\81909\Desktop\shop-project
   ```
2. ファイルを保存後、以下のコマンドでGASへアップロード：
   ```powershell
   npx @google/clasp push --force
   ```
3. 本番のデプロイバージョンを更新：
   ```powershell
   npx @google/clasp deploy -i AKfycbxIYsPdreUjaOFutfshhQIwOLb_Ke1JprcDZj_LAT9oZnc-l-3jZLsPAsFd_QdQJAQ4eA -d "修正内容の説明"
   ```

### ④ GitHub Pages（公開Webサイト）への反映手順
GitコマンドまたはPowerShellスクリプトで反映できます。

```powershell
# 変更したファイルをコミット
git add index.html picking.html request.html "無題.js"
git commit -m "修正内容の説明"
git push origin main
```
※本番リポジトリ（`kawabatasuisanyaeyama.github.io`）へ直接反映する場合は、GitHub REST APIまたは該当リポジトリへのプッシュを行ってください。

---

## 7. トラブルシューティング（よくある不具合と解決策）

### Q1. スマホで開くと「LIFF app '...' was not found.」や「404」と出る
- **原因**: LINE DevelopersのLIFF IDとHTML内の定数が一致していない、またはLIFFのステータスが非公開になっている。
- **対処**:
  1. `index.html` および `request.html` 内の `LIFF_ID` が `2011871129-P4IflaGc` になっているか確認。
  2. [LINE Developersコンソール](https://developers.line.biz/) にログインし、チャネル `2011871129` のLIFFアプリ設定で、エンドポイントURLが `https://kawabatasuisanyaeyama.github.io/` になっているか確認。

### Q2. 注文やピッキング画面で「通信エラー」や「アクセス権が必要です」と出る
- **原因**: GAS Web Appの実行権限が「自分のみ」になっている、または未承認状態。
- **対処**:
  1. [Apps Scriptエディタ](https://script.google.com/u/0/home/projects/1tIl_FYj2PWaSfsvajNXkBgZdgqr4-NMtb9yZsbWPQmSvdfCVQR3n6lv0/edit) を開く。
  2. 右上の「デプロイ」>「デプロイを管理」をクリック。
  3. 鉛筆マーク（編集）を押し、**「次のユーザーとして実行: 自分（kawabatasuisam@gmail.com）」**、**「アクセスできるユーザー: 全員」** になっていることを確認して再デプロイ。

### Q3. ピッキング画面を一般のお客様に見られたくない
- **状態**: ピッキング画面（`picking.html`）へのリンクは、一般公開サイト（`index.html` / `request.html`）から完全に撤去されています。
- **運用ルール**: スタッフ専用のURL（`/picking.html`）は、一般のお客様へは教えず、川畑水産の現場作業員のスマートフォンにのみブックマークまたはホーム画面追加してご利用ください。
