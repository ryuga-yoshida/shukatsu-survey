/**
 * 就活アンケート 収集API（Google Apps Script）
 *
 * 設置手順は README.md を参照。
 *  - POST: アンケートページからの回答を「responses」シートに1行追加
 *  - GET : ?key=READ_KEY で全回答をJSONで返す（集計画面が使用）
 */

var SHEET_NAME = "responses";
var READ_KEY = "CHANGE-ME-to-a-long-random-string"; // 集計画面から読むときの合言葉。必ず変更する

// 人が見やすいように展開する列（この順で並ぶ。無い項目は空欄）
var FLAT_COLUMNS = [
  "q0", "q0_other", "faculty", "faculty_other", "q1", "q2", "q3",
  "q4", "q4_other", "q5", "q5_other", "q6", "q6r", "q6r_other", "q6n",
  "q7",
  "q8", "q9", "q9r", "q9r_other", "q10", "q10_other",
  "q11", "q12", "q12_other",
  "q13", "q13_other", "q13n", "email"
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (!body || !body.answers) return json_({ ok: false, error: "bad_request" });
    var sheet = getSheet_();
    var a = body.answers;
    var row = [
      new Date(),
      body.id || "",
      body.submittedAt || "",
      JSON.stringify(body)
    ];
    FLAT_COLUMNS.forEach(function (k) { row.push(flat_(a[k])); });
    row.push(body.meta && body.meta.ua || "");
    sheet.appendRow(row);
    return json_({ ok: true, id: body.id });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var key = e && e.parameter && e.parameter.key;
  if (key !== READ_KEY) return json_({ ok: false, error: "unauthorized" });
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return json_({ ok: true, count: 0, responses: [] });
  var values = sheet.getRange(2, 1, last - 1, 4).getValues(); // timestamp, id, submittedAt, raw_json
  var out = [];
  values.forEach(function (r) {
    if (!r[3]) return;
    try { var o = JSON.parse(r[3]); o.receivedAt = r[0] instanceof Date ? r[0].toISOString() : String(r[0]); out.push(o); } catch (_) {}
  });
  return json_({ ok: true, count: out.length, responses: out });
}

// ---------- helpers ----------
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    var header = ["受信日時", "id", "送信日時(端末)", "raw_json"].concat(FLAT_COLUMNS).concat(["user_agent"]);
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground("#EAEEFB");
    sh.setFrozenRows(1);
    sh.hideColumns(4); // raw_json は集計用なので隠す
  }
  return sh;
}

function flat_(v) {
  if (v === undefined || v === null) return "";
  if (Array.isArray(v)) return v.join(" | ");
  if (typeof v === "object") {
    // q7 (matrix) → "自己分析: 数時間程度 | ..."、q11 (services) → "リクナビ(3): 理由 | ..."
    if (v.items) {
      if (v.none) return "（使っていない）";
      return v.items.map(function (it) {
        var name = it.name === "その他" ? (it.nameOther || "その他") : it.name;
        return name + "(" + (it.score || "-") + ")" + (it.reason ? ": " + it.reason : "");
      }).join(" | ");
    }
    return Object.keys(v).map(function (k) { return k + ": " + v[k]; }).join(" | ");
  }
  return String(v);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// エディタから一度実行して動作確認するためのテスト
function testInsert_() {
  var sample = { id: "r_test", submittedAt: new Date().toISOString(), answers: { q0: "大学3年生", q1: "少し始めた（自己分析や情報収集程度）", q4: ["自己分析"], q7: { "自己分析": "数時間程度" }, q11: { none: false, items: [{ name: "リクナビ", score: 3, reason: "テスト" }] } }, meta: { ua: "test" } };
  var res = doPost({ postData: { contents: JSON.stringify(sample) } });
  Logger.log(res.getContent());
}


// =====================================================================
// 集計シート（responses の列を COUNTIF で数えるだけ。回答が増えると自動で更新される）
//  スプレッドシートのメニュー「アンケート → 集計シートを作成/更新」から実行
// =====================================================================
var SUMMARY_SHEET = "集計";

function onOpen() {
  SpreadsheetApp.getUi().createMenu("アンケート")
    .addItem("集計シートを作成/更新", "buildSummary")
    .addToUi();
}

function buildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = getSheet_();
  var header = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0];
  var colOf = {}; header.forEach(function (h, i) { colOf[h] = columnLetter_(i + 1); });
  var sh = ss.getSheetByName(SUMMARY_SHEET);
  if (sh) sh.clear(); else sh = ss.insertSheet(SUMMARY_SHEET);

  var rows = [];   // [label, count, pct]
  var styles = []; // {row, kind}
  var total = "COUNTA(" + SHEET_NAME + "!$B$2:$B)";
  rows.push(["回答数", "=" + total, ""]); styles.push({ row: rows.length, kind: "total" });
  rows.push(["", "", ""]);

  SUMMARY_SPEC.forEach(function (q) {
    var col = colOf[q.id];
    if (!col) return;
    var rng = SHEET_NAME + "!$" + col + "$2:$" + col;
    var answered = "COUNTIF(" + rng + ",\"?*\")";
    if (q.type === "matrix") {
      rows.push([q.label, "=" + answered, "回答数"]); styles.push({ row: rows.length, kind: "head" });
      rows.push([""].concat(q.cols)); styles.push({ row: rows.length, kind: "sub" });
      q.rows.forEach(function (r) {
        var line = [r];
        q.cols.forEach(function (c) { line.push("=COUNTIF(" + rng + ",\"*" + esc_(r + ": " + c) + "*\")"); });
        rows.push(line);
      });
    } else {
      var isMulti = q.type === "multi", wild = isMulti || q.wild;
      rows.push([q.label + (isMulti ? "（複数選択・%は回答者に占める割合）" : ""), "=" + answered, "回答数"]); styles.push({ row: rows.length, kind: "head" });
      var headRow = rows.length;
      q.options.forEach(function (o) {
        var f = wild ? "=COUNTIF(" + rng + ",\"*" + esc_(o) + "*\")" : "=COUNTIF(" + rng + ",\"" + esc_(o) + "\")";
        rows.push([o, f, "=IF($B$" + headRow + "=0,\"\",B" + (rows.length + 1) + "/$B$" + headRow + ")"]);
      });
    }
    rows.push(["", "", ""]);
  });

  var width = Math.max.apply(null, rows.map(function (r) { return r.length; }));
  rows = rows.map(function (r) { while (r.length < width) r.push(""); return r; });
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.getRange(1, 3, rows.length, 1).setNumberFormat("0.0%");
  sh.setColumnWidth(1, 420); for (var c = 2; c <= width; c++) sh.setColumnWidth(c, 110);
  styles.forEach(function (st) {
    var r = sh.getRange(st.row, 1, 1, width);
    if (st.kind === "total") r.setFontWeight("bold").setFontSize(12);
    if (st.kind === "head") r.setFontWeight("bold").setBackground("#EAEEFB");
    if (st.kind === "sub") r.setFontColor("#5B6478").setFontSize(9);
  });
  sh.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("「" + SUMMARY_SHEET + "」シートを更新しました。数式なので回答が増えると自動で反映されます。");
}

function esc_(s) { return String(s).replace(/([*?~])/g, "~$1").replace(/"/g, '""'); }
function columnLetter_(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

// questions.js から自動生成（設問を変えたら README の手順で再生成する）
var SUMMARY_SPEC = [ 
  { "id": "q0", "label": "Q0 学年を教えてください。", "type": "single", "options": [ "大学2年生", "大学3年生", "大学4年生", "修士1年", "修士2年", "その他" ] }, 
  { "id": "faculty", "label": "学部の系統を教えてください（任意）", "type": "single", "options": [ "文学・人文・外国語系", "法・政治系", "経済・経営・商学系", "社会・国際・教育系", "理学系", "工学・情報系", "農学・生命科学系", "医学・薬学・看護系", "芸術・デザイン・体育系", "その他" ] }, 
  { "id": "q1", "label": "Q1 就活はどの程度進んでいますか？", "type": "single", "options": [ "全く手をつけていない", "少し始めた（自己分析や情報収集程度）", "本格的に動いている（説明会参加・ES提出等）", "就活は終わっている" ] }, 
  { "id": "q2", "label": "Q2 就活を始めた・始めようと思った時期はいつですか？", "type": "single", "options": [ "大学1・2年生のころ", "大学3年生の春（4〜6月）", "大学3年生の夏（7〜9月）", "大学3年生の秋冬（10〜12月）", "大学3年生の終わり〜4年生（1月以降）", "まだ始めていない" ] }, 
  { "id": "q3", "label": "Q3 就活に対して今どんな気持ちですか？", "type": "single", "options": [ "焦っている", "もっとできたと後悔している", "なんとなく不安だが動けていない", "まだ大丈夫だと思っている", "特に気にしていない" ] }, 
  { "id": "q4", "label": "Q4 「やらなきゃ」と思いながらも、結局できなかった就活関連のことはありますか？", "type": "multi", "options": [ "自己分析", "業界・企業研究", "ES（エントリーシート）の作成", "SPI・筆記試験の対策", "面接の練習", "GD（グループディスカッション）の対策", "インターンへの応募・参加", "OB・OG訪問", "説明会・イベントへの参加", "キャリアセンターへの相談", "特にない（やろうと思ったことは大体できた）", "その他" ] }, 
  { "id": "q5", "label": "Q5 就活を進める上で、何がブレーキになっていると思いますか？", "type": "multi", "options": [ "何から始めればいいかわからない", "調べようとしたが情報が多すぎて疲れた", "自己分析が億劫・やりたくない", "周りもまだ動いていないから急がなくていいと思っている", "時間がない・就活の優先度が低い", "その他" ] }, 
  { "id": "q6", "label": "Q6 一番大きいブレーキ（Q5で選んだ中から1つ）", "type": "single", "wild": true, "options": ["何から始めればいいかわからない", "調べようとしたが情報が多すぎて疲れた", "自己分析が億劫・やりたくない", "周りもまだ動いていないから急がなくていいと思っている", "時間がない・就活の優先度が低い", "その他"] }, 
  { "id": "q6r", "label": "Q6 それが一番大きい理由に近いものはどれですか？", "type": "single", "options": [ "やることの全体像が見えないから", "まとまった時間が取れないから", "失敗が怖い・自信がないから", "面倒でやる気が出ないから", "周りと比べて焦りを感じないから", "相談できる人がいないから", "何を目指したいか自体が決まっていないから", "その他" ] }, 
  { "id": "q7", "label": "Q7 以下の就活対策それぞれに、どのくらいの時間をかけましたか？（かける予定ですか？）", "type": "matrix", "rows": [ "自己分析", "業界・企業研究", "ES作成", "SPI・筆記試験対策", "面接対策", "GD対策", "インターン参加" ], "cols": [ "ほとんどやっていない", "数時間程度", "10〜30時間程度", "30時間以上" ] }, 
  { "id": "q8", "label": "Q8 大学のキャリアセンターイベントや相談窓口を利用したことがありますか？", "type": "single", "options": [ "ある", "ない" ] }, 
  { "id": "q9", "label": "Q9 利用してみてどうでしたか？", "type": "single", "options": [ "とても役に立った", "まあまあ役に立った", "あまり役に立たなかった", "役に立たなかった" ] }, 
  { "id": "q9r", "label": "Q9 そう感じた理由に近いものを選んでください（複数可）", "type": "multi", "options": [ "具体的なアドバイスがもらえた", "自分の状況に合った話ではなく一般論だった", "予約が取りにくい・時間が合わない", "気軽に行ける雰囲気ではなかった", "イベントの内容が自分の段階と合っていなかった", "何を相談すればいいかわからなかった", "その他" ] }, 
  { "id": "q10", "label": "Q10 なぜ利用しませんでしたか？（複数可）", "type": "multi", "options": [ "存在や場所を知らなかった", "行くのが面倒・ハードルが高い", "必要性を感じなかった", "予約や開催日時が合わなかった", "何を相談すればいいかわからなかった", "友人や先輩、ネットの情報で十分だと思った", "その他" ] }, 
  { "id": "q12", "label": "Q12 以下の機能の中で、あったら使いたいと思うものをすべて選んでください。", "type": "multi", "options": [ "AIによる自己分析・性格診断サポート", "AIとの対話でガクチカ・志望動機を作成", "ES添削・採点", "SPI演習", "AI模擬面接（音声で回答・採点）", "GD練習（他ユーザーやAIと）", "業界研究・企業比較", "選考管理（通知メールとの自動連携）", "仮想インターン体験", "就活講義動画・音声コンテンツ", "OB・OG訪問のマッチング", "就活仲間とのコミュニティ・グループ機能", "企業からのスカウト機能", "選考体験談・口コミの閲覧", "面接日程の自動調整", "メンター（社会人）への相談機能", "その他" ] }, 
  { "id": "q13", "label": "Q13 就活を始める・続けるにあたって、どんなサポートがあると助かりますか？（複数可）", "type": "multi", "options": [ "何から始めるか、順番を示してくれる", "締切ややることをリマインドしてくれる", "5分・10分でできる短い練習や課題", "気軽に相談できる先輩・社会人", "同じ状況の仲間と一緒に進められる環境", "ESや面接のフィードバックがすぐ返ってくる", "自分の状況に合わせた具体的なアドバイス", "選考の体験談やリアルな情報", "交通費・スーツ代など金銭面のサポート", "自分ひとりで完結できるツール（人と関わらなくていい）", "その他" ] } ];
