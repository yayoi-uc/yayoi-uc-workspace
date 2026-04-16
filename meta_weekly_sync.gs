// ============================================================
// Meta広告 全シート自動同期スクリプト v2.0
// 対象シート:
//   日次データ / 週次データ / 月次データ / クリエイティブ
//   / キャンペーン別 週次パフォーマンス比較
// ============================================================

// ─── 設定（トークン期限切れ時は META_TOKEN だけ更新）───
const META_TOKEN       = 'EAAFklc4Xw3kBRDzqgmCf0Tk3fgRRPrZB7nKvKJKl4RHnUFnoQDgMUeXi1UBVIoQif3hhPgmmDQUUJmW8jkMuG1yt2G59JKTsv77dEmC0dZAH0s1CsSGxuHSTMRCT0PbVcVwBgxoFThoxDQesYJJF1y8PfuBo4QlJJarEIwRCag8NtMii2aWKDPbRKldAp6PZCNQGtTGotlI2QXtuH4ZCb3JWR4mcnlfUbH9XxJKgLPMSlAmijEHIv584UG4YgPiZC8Qk5OF5FVfPNZA8ZBZA4UJJxTlY5QqXRJuB6QZDZD';
const AD_ACCOUNT       = 'act_154735669086181';
const DAILY_BUDGET_JPY = 33300; // 日次予算（予算変更時は手動更新）

// ─── シート名 ───
const SHEET_DAILY    = '日次データ';
const SHEET_WEEKLY   = '週次データ';
const SHEET_MONTHLY  = '月次データ';
const SHEET_CREATIVE = 'クリエイティブ';
const SHEET_COMP     = 'キャンペーン別 週次パフォーマンス比較';

// ─── キャンペーン対応表（表示名 → Meta上の名前）───
const CAMPAIGN_MAP = {
  'Advantage+ 0426 動画': 'Advantage+ 2025/04/26/動画専用',
  'Advantage+ 1228 動画': 'Advantage+20251228動画専用',
  'Advantage+ Display':   'Advantage+ 2024/05/15/Display専用',
  'collection_購入':      'collection_購入',
  '新規動画（テスト）':   'Advantage+20260326動画アパレル',
};
const CAMPAIGN_ORDER = Object.keys(CAMPAIGN_MAP);

// ============================================================
// メイン：毎週月曜9時に全シート更新
// ============================================================
function weeklyUpdate() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();

  // 先週（月曜〜日曜）を計算
  const dow    = today.getDay(); // 0=日, 1=月...6=土
  const lastMon = new Date(today);
  lastMon.setDate(today.getDate() - dow - 6);
  lastMon.setHours(0, 0, 0, 0);
  const lastSun = new Date(lastMon);
  lastSun.setDate(lastMon.getDate() + 6);
  const weekFrom = fmtDate(lastMon);
  const weekTo   = fmtDate(lastSun);

  Logger.log(`== 週次更新開始 ${weekFrom}〜${weekTo} ==`);

  // 1. 日次データ（先週7日分を追記）
  try { updateDailySheet(ss, weekFrom, weekTo); } catch(e) { Logger.log('日次エラー: ' + e); }

  // 2. 週次データ（先週1行を追記）
  try { updateWeeklySheet(ss, weekFrom, weekTo); } catch(e) { Logger.log('週次エラー: ' + e); }

  // 3. 月次データ（月が変わった週のみ追記）
  if (lastMon.getDate() <= 7) {
    try {
      const prevMonth = new Date(lastMon.getFullYear(), lastMon.getMonth() - 1, 1);
      const monthStr  = Utilities.formatDate(prevMonth, 'Asia/Tokyo', 'yyyy-MM');
      updateMonthlySheet(ss, monthStr);
    } catch(e) { Logger.log('月次エラー: ' + e); }
  }

  // 4. クリエイティブ（直近28日を全件リフレッシュ）
  try {
    const ago28 = new Date(lastSun);
    ago28.setDate(lastSun.getDate() - 27);
    updateCreativeSheet(ss, fmtDate(ago28), weekTo);
  } catch(e) { Logger.log('クリエイティブエラー: ' + e); }

  // 5. キャンペーン別 週次パフォーマンス比較
  try {
    const campData = fetchCampaignInsights(weekFrom, weekTo);
    if (campData) updateCompSheet(ss, campData, lastMon);
  } catch(e) { Logger.log('比較シートエラー: ' + e); }

  Logger.log('== 全シート更新完了 ==');
}

// ============================================================
// 1. 日次データ：先週7日分を追記（重複スキップ）
// ============================================================
function updateDailySheet(ss, dateFrom, dateTo) {
  const sheet = ss.getSheetByName(SHEET_DAILY);
  if (!sheet) { Logger.log('シートなし: ' + SHEET_DAILY); return; }

  // 既存の日付を収集（4行目以降がデータ行）
  const lastRow = sheet.getLastRow();
  const existingDates = new Set();
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, 1).getValues().forEach(r => {
      try { if (r[0]) existingDates.add(fmtDate(new Date(r[0]))); } catch(e) {}
    });
  }

  const rows = fetchAccountInsights(dateFrom, dateTo, 1);
  if (!rows) return;

  let added = 0;
  for (const d of rows) {
    if (existingDates.has(d.date_start)) continue;
    const cpa  = d.cv > 0    ? Math.round(d.spend / d.cv)            : 0;
    const roas = d.spend > 0 ? +((d.revenue / d.spend)).toFixed(2)   : 0;
    const rate = Math.round(d.spend / DAILY_BUDGET_JPY * 100);

    sheet.appendRow([
      new Date(d.date_start + 'T09:00:00+09:00'),
      d.spend, rate,
      d.impressions, d.reach, d.clicks, d.link_clicks,
      +d.ctr.toFixed(2), Math.round(d.cpm), Math.round(d.cpc),
      d.cv, d.revenue, cpa, roas,
      d.add_to_cart, d.checkout,
      getAlert(cpa),
    ]);
    const nr = sheet.getLastRow();
    sheet.getRange(nr, 1).setNumberFormat('yyyy-MM-dd');
    sheet.getRange(nr, 2).setNumberFormat('#,##0');
    added++;
  }
  Logger.log('日次データ: ' + added + '件追加');
}

// ============================================================
// 2. 週次データ：先週1行を追記（重複スキップ）
// ============================================================
function updateWeeklySheet(ss, weekFrom, weekTo) {
  const sheet = ss.getSheetByName(SHEET_WEEKLY);
  if (!sheet) { Logger.log('シートなし: ' + SHEET_WEEKLY); return; }

  const periodLabel = weekFrom + '〜' + weekTo;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const periods = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
                        .map(r => String(r[0]).trim());
    if (periods.includes(periodLabel)) {
      Logger.log('週次データ: 既存 ' + periodLabel);
      return;
    }
  }

  const rows = fetchAccountInsights(weekFrom, weekTo, null);
  if (!rows || rows.length === 0) return;
  const d    = rows[0];
  const cpa  = d.cv > 0    ? Math.round(d.spend / d.cv)          : 0;
  const roas = d.spend > 0 ? +((d.revenue / d.spend)).toFixed(2) : 0;
  const rate = Math.round(d.spend / (DAILY_BUDGET_JPY * 7) * 100);

  sheet.appendRow([
    periodLabel, d.spend, rate,
    d.impressions, d.reach, d.clicks, d.link_clicks,
    +d.ctr.toFixed(2), Math.round(d.cpm), Math.round(d.cpc),
    d.cv, d.revenue, cpa, roas,
    d.add_to_cart, d.checkout,
    getAlert(cpa),
  ]);
  Logger.log('週次データ: 追加 ' + periodLabel);
}

// ============================================================
// 3. 月次データ：指定月を追記（重複スキップ）
// ============================================================
function updateMonthlySheet(ss, monthStr) {
  const sheet = ss.getSheetByName(SHEET_MONTHLY);
  if (!sheet) { Logger.log('シートなし: ' + SHEET_MONTHLY); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const months = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
                       .map(r => String(r[0]).trim());
    if (months.includes(monthStr)) {
      Logger.log('月次データ: 既存 ' + monthStr);
      return;
    }
  }

  const [y, m] = monthStr.split('-').map(Number);
  const dateFrom = monthStr + '-01';
  const dateTo   = fmtDate(new Date(y, m, 0)); // 月末日
  const days     = new Date(y, m, 0).getDate();

  const rows = fetchAccountInsights(dateFrom, dateTo, null);
  if (!rows || rows.length === 0) return;
  const d    = rows[0];
  const cpa  = d.cv > 0    ? Math.round(d.spend / d.cv)          : 0;
  const roas = d.spend > 0 ? +((d.revenue / d.spend)).toFixed(2) : 0;
  const rate = Math.round(d.spend / (DAILY_BUDGET_JPY * days) * 100);

  sheet.appendRow([
    monthStr, d.spend, rate,
    d.impressions, d.reach, d.clicks, d.link_clicks,
    +d.ctr.toFixed(2), Math.round(d.cpm), Math.round(d.cpc),
    d.cv, d.revenue, cpa, roas,
    d.add_to_cart, d.checkout,
    getAlert(cpa),
  ]);
  Logger.log('月次データ: 追加 ' + monthStr);
}

// ============================================================
// 4. クリエイティブ：直近28日の広告別実績を全件リフレッシュ
// ============================================================
function updateCreativeSheet(ss, dateFrom, dateTo) {
  const sheet = ss.getSheetByName(SHEET_CREATIVE);
  if (!sheet) { Logger.log('シートなし: ' + SHEET_CREATIVE); return; }

  const ads = fetchAdInsights(dateFrom, dateTo);
  if (!ads || ads.length === 0) return;

  ads.sort((a, b) => b.spend - a.spend);

  // 新しいデータ行を作成
  const newRows = ads.map(ad => {
    const cpa  = ad.cv > 0    ? Math.round(ad.spend / ad.cv)          : 0;
    const roas = ad.spend > 0 ? +((ad.revenue / ad.spend)).toFixed(2) : 0;
    return [
      ad.ad_id, ad.ad_name, ad.campaign_name, ad.adset_name, ad.status,
      ad.spend, ad.impressions, ad.clicks,
      +ad.ctr.toFixed(2), Math.round(ad.cpm), Math.round(ad.cpc),
      ad.cv, ad.revenue, cpa, roas,
      getAlert(cpa),
    ];
  });

  // 2行目以降を上書き（既存データをクリアして書き直し）
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 16).clearContent();
  }
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, 16).setValues(newRows);
  }
  Logger.log('クリエイティブ: ' + ads.length + '件更新');
}

// ============================================================
// Meta API: アカウントレベル インサイト取得
// timeIncrement=1 → 日別配列, null → 期間合計1行
// ============================================================
function fetchAccountInsights(dateFrom, dateTo, timeIncrement) {
  const fields = [
    'date_start', 'spend', 'impressions', 'reach',
    'clicks', 'inline_link_clicks', 'ctr', 'cpm', 'cpc',
    'actions', 'action_values'
  ].join(',');

  let url = 'https://graph.facebook.com/v21.0/' + AD_ACCOUNT + '/insights'
    + '?access_token=' + META_TOKEN
    + '&fields=' + encodeURIComponent(fields)
    + '&level=account'
    + '&time_range=' + encodeURIComponent(JSON.stringify({ since: dateFrom, until: dateTo }))
    + '&limit=100';
  if (timeIncrement) url += '&time_increment=' + timeIncrement;

  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.error) { Logger.log('APIエラー: ' + JSON.stringify(json.error)); return null; }
    return (json.data || []).map(r => ({
      date_start:  r.date_start,
      spend:       Math.round(parseFloat(r.spend) || 0),
      impressions: parseInt(r.impressions) || 0,
      reach:       parseInt(r.reach) || 0,
      clicks:      parseInt(r.clicks) || 0,
      link_clicks: parseInt(r.inline_link_clicks) || 0,
      ctr:         parseFloat(r.ctr) || 0,
      cpm:         parseFloat(r.cpm) || 0,
      cpc:         parseFloat(r.cpc) || 0,
      cv:          getAction(r.actions,       'purchase'),
      revenue:     getAction(r.action_values, 'purchase'),
      add_to_cart: getAction(r.actions,       'add_to_cart'),
      checkout:    getAction(r.actions,       'initiate_checkout'),
    }));
  } catch(e) {
    Logger.log('Fetchエラー: ' + e.message);
    return null;
  }
}

// ============================================================
// Meta API: 広告（クリエイティブ）レベル インサイト取得
// Step1: 広告リストでステータス取得
// Step2: インサイトで実績取得 → 結合
// ============================================================
function fetchAdInsights(dateFrom, dateTo) {
  // Step1: 広告リスト（ステータス取得）
  const adMap = {};
  try {
    const listUrl = 'https://graph.facebook.com/v21.0/' + AD_ACCOUNT + '/ads'
      + '?access_token=' + META_TOKEN
      + '&fields=id,name,effective_status,campaign{name},adset{name}'
      + '&limit=200';
    const res  = UrlFetchApp.fetch(listUrl, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (!json.error) {
      (json.data || []).forEach(ad => {
        adMap[ad.id] = {
          status:        ad.effective_status || 'UNKNOWN',
          campaign_name: (ad.campaign && ad.campaign.name) || '',
          adset_name:    (ad.adset    && ad.adset.name)    || '',
        };
      });
    }
  } catch(e) { Logger.log('広告リスト取得エラー: ' + e.message); }

  // Step2: インサイト取得
  const fields = [
    'ad_id', 'ad_name', 'campaign_name', 'adset_name',
    'spend', 'impressions', 'clicks', 'ctr', 'cpm', 'cpc',
    'actions', 'action_values'
  ].join(',');

  const url = 'https://graph.facebook.com/v21.0/' + AD_ACCOUNT + '/insights'
    + '?access_token=' + META_TOKEN
    + '&fields=' + encodeURIComponent(fields)
    + '&level=ad'
    + '&time_range=' + encodeURIComponent(JSON.stringify({ since: dateFrom, until: dateTo }))
    + '&limit=200';

  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.error) { Logger.log('APIエラー: ' + JSON.stringify(json.error)); return null; }
    return (json.data || []).map(r => {
      const meta = adMap[r.ad_id] || {};
      return {
        ad_id:         r.ad_id,
        ad_name:       r.ad_name,
        campaign_name: meta.campaign_name || r.campaign_name || '',
        adset_name:    meta.adset_name    || r.adset_name    || '',
        status:        meta.status        || 'UNKNOWN',
        spend:         Math.round(parseFloat(r.spend) || 0),
        impressions:   parseInt(r.impressions) || 0,
        clicks:        parseInt(r.clicks) || 0,
        ctr:           parseFloat(r.ctr) || 0,
        cpm:           parseFloat(r.cpm) || 0,
        cpc:           parseFloat(r.cpc) || 0,
        cv:            getAction(r.actions,       'purchase'),
        revenue:       getAction(r.action_values, 'purchase'),
      };
    });
  } catch(e) {
    Logger.log('Fetchエラー: ' + e.message);
    return null;
  }
}

// ============================================================
// Meta API: キャンペーンレベル インサイト取得
// ============================================================
function fetchCampaignInsights(dateFrom, dateTo) {
  const fields = [
    'campaign_name', 'campaign_id',
    'spend', 'impressions', 'clicks', 'ctr', 'cpc',
    'actions', 'action_values'
  ].join(',');

  const url = 'https://graph.facebook.com/v21.0/' + AD_ACCOUNT + '/insights'
    + '?access_token=' + META_TOKEN
    + '&fields=' + fields
    + '&level=campaign'
    + '&time_range=' + encodeURIComponent(JSON.stringify({ since: dateFrom, until: dateTo }))
    + '&limit=50';

  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.error) { Logger.log('Meta APIエラー: ' + JSON.stringify(json.error)); return null; }

    const results = {};
    for (const row of (json.data || [])) {
      const cv      = getAction(row.actions,       'purchase');
      const revenue = getAction(row.action_values, 'purchase');
      const spend   = parseFloat(row.spend) || 0;
      results[row.campaign_name] = {
        campaignId:   row.campaign_id,
        campaignName: row.campaign_name,
        spend:        Math.round(spend),
        impressions:  parseInt(row.impressions) || 0,
        clicks:       parseInt(row.clicks) || 0,
        ctr:          parseFloat(row.ctr) || 0,
        cpc:          parseFloat(row.cpc) || 0,
        cv, revenue,
        cpa:  cv > 0    ? Math.round(spend / cv)          : 0,
        roas: spend > 0 ? +((revenue / spend)).toFixed(2) : 0,
      };
    }
    return results;
  } catch(e) {
    Logger.log('Fetchエラー: ' + e.message);
    return null;
  }
}

// ============================================================
// 5. キャンペーン別 週次パフォーマンス比較シートを更新
// ============================================================
function updateCompSheet(ss, campData, weekStart) {
  const sheet = ss.getSheetByName(SHEET_COMP);
  if (!sheet) { Logger.log('シートなし: ' + SHEET_COMP); return; }

  const monthLabel   = (weekStart.getMonth() + 1) + '/1';
  const weekNum      = getWeekNumInMonth(weekStart);
  Logger.log('比較シート: ' + monthLabel + ' W' + weekNum);

  let monthRow = findRowByValue(sheet, 1, monthLabel);
  if (!monthRow) monthRow = addMonthBlock(sheet, monthLabel);

  const weekColStart = (weekNum - 1) * 5 + 2;
  const allValues    = sheet.getDataRange().getValues();
  let totalSpend = 0, totalCv = 0;

  for (const displayName of CAMPAIGN_ORDER) {
    const metaName = CAMPAIGN_MAP[displayName];
    const d = campData[metaName];
    if (!d) { Logger.log('データなし: ' + displayName); continue; }

    const campRow = findCampaignRow(allValues, monthRow, displayName);
    if (!campRow) { Logger.log('行なし: ' + displayName); continue; }

    const cpa   = d.cv > 0 ? Math.round(d.spend / d.cv) : 0;
    const judge = getJudgment(d.spend, d.cv, cpa, d.roas);

    sheet.getRange(campRow, weekColStart    ).setValue(d.spend).setNumberFormat('#,##0');
    sheet.getRange(campRow, weekColStart + 1).setValue(d.cv);
    sheet.getRange(campRow, weekColStart + 2).setValue(cpa).setNumberFormat('#,##0');
    sheet.getRange(campRow, weekColStart + 3).setValue(d.roas);
    sheet.getRange(campRow, weekColStart + 4).setValue(judge);

    totalSpend += d.spend;
    totalCv    += d.cv;
    Logger.log(displayName + ': ¥' + d.spend + ' / ' + d.cv + 'CV / CPA¥' + cpa + ' / ROAS' + d.roas);
  }

  const totalRow = findCampaignRow(allValues, monthRow, '合計');
  if (totalRow) {
    const totalCpa = totalCv > 0 ? Math.round(totalSpend / totalCv) : 0;
    sheet.getRange(totalRow, weekColStart    ).setValue(totalSpend).setNumberFormat('#,##0');
    sheet.getRange(totalRow, weekColStart + 1).setValue(totalCv);
    sheet.getRange(totalRow, weekColStart + 2).setValue(totalCpa).setNumberFormat('#,##0');
  }

  SpreadsheetApp.flush();
  Logger.log('比較シート更新: ' + monthLabel + ' W' + weekNum);
}

// ============================================================
// キャンペーン別比較シート: 月ブロック新規作成
// ============================================================
function addMonthBlock(sheet, monthLabel) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  let startRow  = lastRow + 2;

  const wVals = Array(26).fill('');
  wVals[0] = monthLabel;
  [1, 6, 11, 16, 21].forEach((c, i) => { wVals[c] = 'W' + (i + 1); });
  sheet.getRange(startRow, 1, 1, 26).setValues([wVals]).setFontWeight('bold').setBackground('#d9e8ff');
  startRow++;

  const hVals = Array(26).fill('');
  hVals[0] = '　';
  for (let w = 0; w < 5; w++) {
    const b = w * 5 + 1;
    hVals[b]     = '費用';
    hVals[b + 1] = 'CV';
    hVals[b + 2] = 'CPA';
    hVals[b + 3] = 'ROAS';
    hVals[b + 4] = '判定';
  }
  sheet.getRange(startRow, 1, 1, 26).setValues([hVals]).setBackground('#eaf2ff');
  startRow++;

  for (const name of [...CAMPAIGN_ORDER, '合計']) {
    const r = sheet.getRange(startRow, 1, 1, 26);
    const v = Array(26).fill('');
    v[0] = name;
    r.setValues([v]);
    if (name === '合計') r.setFontWeight('bold').setBackground('#f5f5f5');
    startRow++;
  }

  const monthHeaderRow = startRow - CAMPAIGN_ORDER.length - 3;
  Logger.log('月ブロック作成: ' + monthLabel + ' 行' + monthHeaderRow);
  return monthHeaderRow;
}

// ============================================================
// 判定ロジック（キャンペーン別比較シート用）
// ============================================================
function getJudgment(spend, cv, cpa, roas) {
  if (spend >= 30000 && cpa > 6500) return '🔴 即停止';
  if (spend >= 50000 && cpa > 5300) return '🟡 停止検討';
  if (roas >= 3.8 && cpa <= 3400)   return '🟢 拡大+15%';
  if (roas >= 2.8)                  return '🟢 拡大+10%';
  if (cpa > 0 && cpa <= 5000)       return '✅ 正常';
  if (cpa === 0)                    return '—';
  return '⚠️ 要確認';
}

// アラート判定（日次/週次/月次/クリエイティブ共通）
function getAlert(cpa) {
  if (!cpa || cpa <= 0) return '';
  if (cpa < 2000)  return '🔵予算拡大';
  if (cpa < 3000)  return '🟢現状維持';
  if (cpa < 4000)  return '🟡週次改善';
  if (cpa <= 6500) return '🟠3日観察';
  return '🔴即停止';
}

// actions / action_values から指定タイプの値を取得
function getAction(arr, type) {
  if (!arr) return 0;
  const hit = arr.find(a => a.action_type === type);
  return hit ? parseFloat(hit.value) : 0;
}

// ============================================================
// ユーティリティ
// ============================================================
function fmtDate(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function getWeekNumInMonth(date) {
  const year  = date.getFullYear();
  const month = date.getMonth();
  const firstDay = new Date(year, month, 1);
  const dow = firstDay.getDay();
  const offsetToMonday = dow === 0 ? 6 : dow - 1;
  const firstMonday = new Date(firstDay);
  firstMonday.setDate(firstDay.getDate() - offsetToMonday);
  return Math.floor((date - firstMonday) / 86400000 / 7) + 1;
}

function findRowByValue(sheet, col, value) {
  if (sheet.getLastRow() === 0) return null;
  const data = sheet.getRange(1, col, sheet.getLastRow(), 1).getValues();
  const [searchMonth, searchDay] = value.split('/').map(Number);
  for (let i = 0; i < data.length; i++) {
    const cell = data[i][0];
    if (!cell && cell !== 0) continue;
    if (cell instanceof Date) {
      if (cell.getMonth() + 1 === searchMonth && cell.getDate() === searchDay) return i + 1;
      continue;
    }
    const cellStr = String(cell).trim();
    if (cellStr === value) return i + 1;
    const parts = cellStr.split('/');
    if (parts.length >= 2) {
      const m = parseInt(parts[parts.length - 2]);
      const d = parseInt(parts[parts.length - 1]);
      if (m === searchMonth && d === searchDay) return i + 1;
    }
  }
  return null;
}

function findCampaignRow(allValues, monthRow, campaignName) {
  for (let i = monthRow; i < Math.min(monthRow + 10, allValues.length); i++) {
    if (String(allValues[i][0]).trim() === campaignName) return i + 1;
  }
  return null;
}

// ============================================================
// 日次更新：平日8時に日次データのみ更新
// ============================================================
function dailyUpdate() {
  const today = new Date();
  const dow   = today.getDay(); // 0=日, 6=土
  if (dow === 0 || dow === 6) {
    Logger.log('土日のためスキップ');
    return;
  }

  // 昨日の日付
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const dateStr = fmtDate(yesterday);

  Logger.log('日次更新: ' + dateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    updateDailySheet(ss, dateStr, dateStr);
  } catch(e) {
    Logger.log('日次更新エラー: ' + e);
  }
}

// ============================================================
// トリガー設定（初回に1回だけ実行）
// ============================================================
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 毎週月曜9時：全シート更新
  ScriptApp.newTrigger('weeklyUpdate')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // 毎日8時：日次データのみ更新（土日は関数内でスキップ）
  ScriptApp.newTrigger('dailyUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('✅ トリガー設定完了：月曜9時（全シート）＋毎日8時（日次）');
}

// ============================================================
// 今すぐテスト実行
// ============================================================
function runNow() {
  weeklyUpdate();
}
