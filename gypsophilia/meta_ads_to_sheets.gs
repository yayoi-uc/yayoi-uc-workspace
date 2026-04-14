// ============================================================
// gypsophilia Meta広告 日次データ取得 → Google Sheets書き込み
// ============================================================
// 【セットアップ手順】
// 1. スプレッドシートを開く → 拡張機能 → Apps Script
// 2. このコードを貼り付けて保存
// 3. setupDailyTrigger() を一度手動実行してトリガーを設定
// 4. アクセストークンは60日で期限切れ → 期限前にトークンを更新すること
// ============================================================

// ====== 設定 ======
const CONFIG = {
  AD_ACCOUNT_ID:  '154735669086181',
  ACCESS_TOKEN:   'EAAFklc4Xw3kBRGkVVLb8QcvewZC2aeuWxDtG4j511m8AnM7oRaMXOuNAQC4R4o0ODjAQ5vs3dDrRTX92hHz3SBOFbunvN1FSEGif49wgdDXGuD2OzC8TauVtZCVEEX7fintBnrGYppB90NiFWoLjaSFB5TuUGehRT68jkKPZAeXnzpdZCFrEjAqMa30ktIqZAj7h6mrK947TBvy31CG8dug8IWvyOWe0wmqP3fryjAPwqx2Yh',
  SPREADSHEET_ID: '1kd6HD4NOECKNl4knsklsrcUDWAOs-ve3xhULEtKrEt8',
  SHEET_GID:      918877948,
  DAILY_BUDGET:   33333,
  API_VERSION:    'v20.0',
};

// ====== メイン関数（毎朝8時に自動実行） ======
function appendDailyData() {
  const date = getYesterdayJST();
  Logger.log(`対象日: ${date}`);

  try {
    const insight = fetchMetaInsight(date);
    const row = buildRow(date, insight);
    writeToSheet(row);
    Logger.log(`✅ ${date} のデータを書き込みました`);
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
    throw e;
  }
}

// ====== 昨日の日付（JST）を返す ======
function getYesterdayJST() {
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  jst.setDate(jst.getDate() - 1);
  return Utilities.formatDate(jst, 'Asia/Tokyo', 'yyyy-MM-dd');
}

// ====== Meta Ads API からインサイト取得 ======
function fetchMetaInsight(date) {
  const fields = [
    'spend',
    'impressions',
    'reach',
    'clicks',
    'inline_link_clicks',
    'ctr',
    'cpm',
    'cpc',
    'actions',
    'action_values',
  ].join(',');

  const params = [
    `fields=${fields}`,
    `time_range={"since":"${date}","until":"${date}"}`,
    `access_token=${CONFIG.ACCESS_TOKEN}`,
  ].join('&');

  const url = `https://graph.facebook.com/${CONFIG.API_VERSION}/act_${CONFIG.AD_ACCOUNT_ID}/insights?${params}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());

  if (json.error) {
    throw new Error(`Meta API エラー: ${json.error.message}`);
  }

  return (json.data && json.data.length > 0) ? json.data[0] : null;
}

// ====== インサイトから1行分のデータを組み立て ======
function buildRow(date, insight) {
  if (!insight) {
    // データなし
    return [date, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '-', 0, 0, 0, '🔴即停止'];
  }

  const spend       = Math.round(parseFloat(insight.spend)           || 0);
  const impressions = parseInt(insight.impressions)                  || 0;
  const reach       = parseInt(insight.reach)                        || 0;
  const clicks      = parseInt(insight.clicks)                       || 0;
  const linkClicks  = parseInt(insight.inline_link_clicks)           || 0;
  const ctr         = round2(parseFloat(insight.ctr)                 || 0);
  const cpm         = Math.round(parseFloat(insight.cpm)             || 0);
  const cpc         = Math.round(parseFloat(insight.cpc)             || 0);

  const cv          = getActionVal(insight.actions,       'purchase');
  const cvRevenue   = Math.round(getActionVal(insight.action_values, 'purchase'));
  const addToCart   = getActionVal(insight.actions,       'add_to_cart');
  const checkout    = getActionVal(insight.actions,       'initiate_checkout');

  const cpa         = cv > 0 ? Math.round(spend / cv) : 0;
  const roas        = spend > 0 ? round2(cvRevenue / spend) : 0;
  const budgetRate  = Math.round((spend / CONFIG.DAILY_BUDGET) * 100);
  const alert       = calcAlert(cpa, cv);

  return [
    date,
    spend,
    budgetRate,
    impressions,
    reach,
    clicks,
    linkClicks,
    ctr,
    cpm,
    cpc,
    cv,
    cvRevenue,
    cv > 0 ? cpa : '-',
    roas,
    addToCart,
    checkout,
    alert,
  ];
}

// ====== スプレッドシートに行を追記 ======
function writeToSheet(row) {
  const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = getSheetByGid(ss, CONFIG.SHEET_GID);
  if (!sheet) throw new Error(`GID ${CONFIG.SHEET_GID} のシートが見つかりません`);
  sheet.appendRow(row);
}

// ====== GIDでシートを取得 ======
function getSheetByGid(spreadsheet, gid) {
  return spreadsheet.getSheets().find(s => s.getSheetId() === gid) || null;
}

// ====== アクションの値を取得 ======
function getActionVal(actions, type) {
  if (!actions) return 0;
  const hit = actions.find(a => a.action_type === type);
  return hit ? Math.round(parseFloat(hit.value)) : 0;
}

// ====== アラートレベル判定 ======
function calcAlert(cpa, cv) {
  if (cv === 0)    return '🔴即停止';
  if (cpa > 6500)  return '🔴即停止';
  if (cpa > 4000)  return '🟠3日観察';
  if (cpa > 3000)  return '🟡週次改善';
  if (cpa > 2000)  return '🟢現状維持';
  return '🔵予算拡大';
}

// ====== 小数点2桁に丸める ======
function round2(n) {
  return Math.round(n * 100) / 100;
}

// ====== 毎日8:00 JST トリガーをセットアップ ======
// ※ 初回のみ手動で実行してください
function setupDailyTrigger() {
  // 既存トリガーを全削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('appendDailyData')
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .inTimezone('Asia/Tokyo')
    .everyDays(1)
    .create();

  Logger.log('✅ 毎日 8:00 JST のトリガーを設定しました');
}

// ====== 動作テスト用：指定日付でデータ取得・書き込み ======
function testDate() {
  const date = '2026-04-11'; // テストしたい日付に変更
  const insight = fetchMetaInsight(date);
  const row = buildRow(date, insight);
  Logger.log(JSON.stringify(row));
  // writeToSheet(row); // 実際に書き込む場合はコメントアウトを外す
}
