# 就活に関するアンケート（就活ナビ）

アンケートページ＋集計画面。回答は Google スプレッドシートに貯まります。

```
index.html        アンケート（回答者向け）
dashboard.html    集計画面（社内向け・URLは共有しない）
questions.js      設問定義（両ページ共通。文言や選択肢はここを直す）
config.js         送信先URLなどの設定
apps-script/Code.gs   スプレッドシート側のAPI
```

## セットアップ（10分）

### 1. スプレッドシートと Apps Script を用意
1. Google スプレッドシートを新規作成（名前は何でもOK）
2. メニュー **拡張機能 → Apps Script** を開く
3. `apps-script/Code.gs` の中身を貼り付け、`READ_KEY` を長いランダム文字列に変える
4. 保存 → 右上 **デプロイ → 新しいデプロイ**
   - 種類：**ウェブアプリ**
   - 次のユーザーとして実行：**自分**
   - アクセスできるユーザー：**全員**
5. 表示された **ウェブアプリのURL**（`https://script.google.com/macros/s/…/exec`）をコピー
   - 初回は権限の承認画面が出るので許可する
   - エディタで `testInsert_` を実行すると `responses` シートにテスト行が入り、動作確認できる（あとで行を消してOK）

### 2. 送信先を設定
`config.js` の `endpoint` にウェブアプリのURLを入れる。

```js
window.SURVEY_CONFIG = {
  endpoint: "https://script.google.com/macros/s/XXXX/exec",
  ...
};
```

### 3. ホスティング
静的ファイルなのでどこでも動く。GitHub Pages の例：

```bash
cd ~/shukatsu-survey && git init && git add . && git commit -m "survey" && gh repo create shukatsu-survey --private --source=. --push
```

GitHub の Settings → Pages で Branch: main / root を選ぶと
`https://<user>.github.io/shukatsu-survey/` で回答ページが公開される。

> `dashboard.html` も同じ場所に置かれるが、シートの合言葉（READ_KEY）を知らないと中身は見えない。
> より厳密に隠したい場合は dashboard.html を別の場所（ローカルなど）に置く。

### 4. 集計画面を開く
`dashboard.html` を開き、右上「設定」から **ウェブアプリURL** と **READ_KEY** を入力 → 保存して読み込む。
設定はそのブラウザにだけ保存される。

- 「サンプルデータで見る」…架空の60件で画面の雰囲気を確認できる
- 「JSON/CSVファイルから読む」…シートを「CSVでダウンロード」したファイルも読める（オフライン集計用）
- 上部チップで 学年／進捗／気持ち による絞り込み。全グラフに反映される
- 「CSV」で絞り込み後の回答を保存

## 設問を変えたいとき
`questions.js` だけ直せば、回答ページと集計画面の両方に反映される。
- `type`: `single`（1つ選択）/ `multi`（複数）/ `matrix`（表）/ `services`（サービス評価）/ `text` / `email`
- `other: true` で「その他」＋自由記述欄が付く
- `showIf: { q8: "ある" }` で条件表示
- `from: "q5"` で「前の設問で選んだものから選ぶ」

※ 選択肢を増やす分には既存データと共存できる。文言を変えると過去回答は旧文言のまま集計される。

## 動作メモ
- `endpoint` が空のときは送信せず、内容をブラウザのコンソールに出す（テストモード）
- 回答の途中経過は端末の localStorage に自動保存。送信完了で消える
- Apps Script への POST は `Content-Type` を付けずに送る（プリフライト回避のため）。変えないこと
