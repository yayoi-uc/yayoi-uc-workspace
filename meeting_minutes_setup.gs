// ============================================================
// 週次ミーティング 議事録セットアップ & 自動入力スクリプト
// 対象スプレッドシート: 議事録
// ============================================================
// 【セットアップ手順】
// 1. 議事録スプレッドシートを開く → 拡張機能 → Apps Script
// 2. このコードを貼り付けて保存
// 3. setupAllSheets() を一度手動実行してシートを初期化
// 4. 毎週月曜8:00に createWeeklyMeeting() が自動実行される
// ============================================================

// ─── 他シートのスプレッドシートID ───
const SPREADSHEET_IDS = {
  META_ADS:    '1kd6HD4NOECKNl4knsklsrcUDWAOs-ve3xhULEtKrEt8', // Meta広告データ
  INFLUENCER:  '1prt8F8EAoHlGlkIv6uLkTGXZwIUD13-_QDtIdvEOMIE', // インフルエンサーマスター
  MARKETING:   '1fQg_01ViGrZEoT38LifM_zh-U4vah_8ORE4xTodz708', // マーケティング全体数値
};

// ─── Meta広告 目標値 ───
const TARGETS = {
  ROAS_IDEAL:        7.0,   // 理想ROAS（700%）
  ROAS_MIN:          5.0,   // 最低ラインRoas（500%）
  ROAS:              5.0,   // 判定基準（最低ラインを基準値として使用）
  CPA:               2000,
  GIFTING_PER_MONTH: 10,
  DM_PER_MONTH:      35,
  CREATIVE_PER_WEEK: 2,
};

// ─── シート名 ───
const SHEET_WEEKLY  = '火曜ミーティング';
const SHEET_FRIDAY  = '金曜進捗';
const SHEET_MONTHLY = '月次サマリー';
const SHEET_施策LOG = '施策ログ';

// ============================================================
// 初期セットアップ（一度だけ手動実行）
// ============================================================
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存シートを取得or新規作成
  _getOrCreateSheet(ss, SHEET_WEEKLY);
  _getOrCreateSheet(ss, SHEET_FRIDAY);
  _getOrCreateSheet(ss, SHEET_MONTHLY);
  _getOrCreateSheet(ss, SHEET_施策LOG);

  setupWeeklySheet(ss);
  setupFridaySheet(ss);
  setupMonthlySheet(ss);
  setup施策LogSheet(ss);

  setupTriggers();

  SpreadsheetApp.getUi().alert('✅ 議事録セットアップ完了！\n\n各シートのヘッダーを設置しました。\n月曜8:00に自動データ取得が始まります。');
}

// ============================================================
// 火曜ミーティングシート セットアップ
// ============================================================
function setupWeeklySheet(ss) {
  const sh = ss.getSheetByName(SHEET_WEEKLY);
  sh.clearContents();
  sh.clearFormats();

  // 列幅設定
  sh.setColumnWidth(1, 30);   // A: インデックス
  sh.setColumnWidth(2, 180);  // B: 項目
  sh.setColumnWidth(3, 120);  // C: 実績
  sh.setColumnWidth(4, 120);  // D: 前週比
  sh.setColumnWidth(5, 100);  // E: 目標
  sh.setColumnWidth(6, 280);  // F: メモ/コメント

  // ヘッダー行（固定）
  const headerRow = ['', '項目', '実績', '前週比', '目標', 'メモ'];
  sh.getRange(1, 1, 1, 6).setValues([headerRow]);
  sh.getRange(1, 1, 1, 6)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sh.setFrozenRows(1);

  // サンプルとして今週のエントリを作成
  _appendWeeklyEntry(sh, _getWeekRange(new Date()));

  SpreadsheetApp.flush();
  Logger.log('✅ 火曜ミーティングシート セットアップ完了');
}

// ============================================================
// 金曜進捗シート セットアップ
// ============================================================
function setupFridaySheet(ss) {
  const sh = ss.getSheetByName(SHEET_FRIDAY);
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 280);

  const headerRow = ['', 'アクション', '担当', 'ステータス', 'メモ'];
  sh.getRange(1, 1, 1, 5).setValues([headerRow]);
  sh.getRange(1, 1, 1, 5)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sh.setFrozenRows(1);

  _appendFridayEntry(sh, new Date());

  SpreadsheetApp.flush();
  Logger.log('✅ 金曜進捗シート セットアップ完了');
}

// ============================================================
// 月次サマリーシート セットアップ
// ============================================================
function setupMonthlySheet(ss) {
  const sh = ss.getSheetByName(SHEET_MONTHLY);
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 280);

  const headers = ['', '指標', '今月実績', '先月実績', '目標', 'メモ'];
  sh.getRange(1, 1, 1, 6).setValues([headers]);
  sh.getRange(1, 1, 1, 6)
    .setBackground('#16213e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sh.setFrozenRows(1);

  SpreadsheetApp.flush();
  Logger.log('✅ 月次サマリーシート セットアップ完了');
}

// ============================================================
// 毎週月曜8:00 自動実行: 火曜ミーティング用エントリ追加
// ============================================================
function createWeeklyMeeting() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(SHEET_WEEKLY);
  const week = _getWeekRange(new Date());

  _appendWeeklyEntry(sh, week);
  Logger.log(`✅ 週次ミーティングエントリ追加: ${week.label}`);
}

// ============================================================
// 毎週木曜8:00 自動実行: 金曜進捗ミーティング用エントリ追加
// ============================================================
function createFridayMeeting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_FRIDAY);

  _appendFridayEntry(sh, new Date());
  Logger.log('✅ 金曜進捗エントリ追加');
}

// ============================================================
// 火曜ミーティング エントリ本体
// ============================================================
function _appendWeeklyEntry(sh, week) {
  const lastRow = sh.getLastRow();
  let row = lastRow + (lastRow > 1 ? 2 : 1); // 前のエントリと1行あける

  // ── データ取得 ──
  const ads  = _fetchMetaAdsWeekly(week.from, week.to);
  const shop = _fetchShopifyWeekly(week.from, week.to);
  const inf  = _fetchGiftingMonthly();

  // ─────────────────────────────────────────
  // セクション0: ヘッダー（日付・対象期間）
  // ─────────────────────────────────────────
  const titleText = `🗓️ 火曜ミーティング  ${week.label}  ▶ 対象期間: ${week.from} 〜 ${week.to}`;
  sh.getRange(row, 1, 1, 6).merge()
    .setValue(titleText)
    .setBackground('#0f3460')
    .setFontColor('#e94560')
    .setFontWeight('bold')
    .setFontSize(12);
  row++;

  // ─────────────────────────────────────────
  // セクション1: 売上サマリー
  // ─────────────────────────────────────────
  row = _writeSectionHeader(sh, row, '① 売上サマリー（先週）', '#16213e');

  const salesRows = [
    ['', '週間純売上 (円)',    _fmt(shop.netSales),     shop.salesWoW,  '—',     ''],
    ['', '注文数',            _fmt(shop.orders),        shop.ordersWoW, '—',     ''],
    ['', '客単価・新規 (円)', _fmt(shop.aovNew),        '—',            '—',     ''],
    ['', '客単価・リピ (円)', _fmt(shop.aovReturn),     '—',            '—',     ''],
    ['', 'CV率 (%)',          _pct(shop.cvr),           '—',            '—',     ''],
    ['', '新規 / リピ比',     shop.newVsReturn,         '—',            '—',     ''],
  ];
  row = _writeDataRows(sh, row, salesRows, '#1a1a2e', '#0d0d0d');

  // ─────────────────────────────────────────
  // セクション2: Meta広告
  // ─────────────────────────────────────────
  row = _writeSectionHeader(sh, row, '② Meta広告（先週）', '#16213e');

  const adsRows = [
    ['', '消化金額 (円)',      _fmt(ads.spend),       ads.spendWoW,    '—',                                         ''],
    ['', 'ROAS',               _num(ads.roas, 2),     ads.roasWoW,     `理想 ${TARGETS.ROAS_IDEAL*100}% / 最低 ${TARGETS.ROAS_MIN*100}%`, _roasStatus(ads.roas)],
    ['', 'CPA (円)',           _fmt(ads.cpa),         ads.cpaWoW,      `目標 ¥${TARGETS.CPA}以下`,                 _cpaStatus(ads.cpa)],
    ['', 'CTR (%)',            _pct(ads.ctr),         '—',             '—',                                         ''],
    ['', 'CPM (円)',           _fmt(ads.cpm),         '—',             '—',                                         ''],
    ['', 'CV数',               _num(ads.cv, 0),       '—',             '—',                                         ''],
    ['', 'CV売上 (円)',        _fmt(ads.cvSales),     '—',             '—',                                         ''],
    ['', '追加クリエイティブ数', '（手入力）',         '—',             `週${TARGETS.CREATIVE_PER_WEEK}本+`,         ''],
  ];
  row = _writeDataRows(sh, row, adsRows, '#1a1a2e', '#0d0d0d');

  // ─────────────────────────────────────────
  // セクション3: ギフティング
  // ─────────────────────────────────────────
  row = _writeSectionHeader(sh, row, '③ ギフティング（今月進捗）', '#16213e');

  const giftingRows = [
    ['', 'DM送信数',          _num(inf.dmCount, 0),       '—', `目標 ${TARGETS.DM_PER_MONTH}件`, ''],
    ['', '返答数',            _num(inf.replyCount, 0),    '—', '—',                              ''],
    ['', '返答率 (%)',        _pct(inf.replyRate),        '—', '—',                              ''],
    ['', 'ギフティング実施数', _num(inf.giftingDone, 0),  '—', `目標 ${TARGETS.GIFTING_PER_MONTH}件`, _giftingStatus(inf.giftingDone)],
    ['', '投稿済み数',        _num(inf.posted, 0),        '—', '—',                              ''],
    ['', '新規素材（広告転用可）', '（手入力）',          '—', '—',                              '広告転用した本数'],
  ];
  row = _writeDataRows(sh, row, giftingRows, '#1a1a2e', '#0d0d0d');

  // ─────────────────────────────────────────
  // セクション4: 今週の投稿予定（手入力）
  // ─────────────────────────────────────────
  row = _writeSectionHeader(sh, row, '④ 今週の投稿予定（ギフティング）', '#16213e');

  const postRows = [
    ['', '@インスタID', '投稿予定日', 'ジャンル', '品番', ''],
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ];
  row = _writeDataRows(sh, row, postRows, '#1a1a2e', '#0d0d0d');

  // ─────────────────────────────────────────
  // セクション5: 課題・懸念事項（手入力）
  // ─────────────────────────────────────────
  row = _writeSectionHeader(sh, row, '⑤ 課題・懸念事項', '#16213e');

  const issueRows = [
    ['', '課題内容', 'オーナー', '緊急度', '期限', 'ステータス'],
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ];
  row = _writeDataRows(sh, row, issueRows, '#1a1a2e', '#0d0d0d');

  // ─────────────────────────────────────────
  // セクション6: アクションアイテム（手入力）
  // ─────────────────────────────────────────
  row = _writeSectionHeader(sh, row, '⑥ アクションアイテム（今週）', '#16213e');

  const actionRows = [
    ['', 'アクション内容', '担当', 'Due', 'ステータス', ''],
    ['', '', 'yayoi', '',   '[ ] 未', ''],
    ['', '', 'shoi',  '',   '[ ] 未', ''],
    ['', '', 'MD',    '',   '[ ] 未', ''],
  ];
  row = _writeDataRows(sh, row, actionRows, '#1a1a2e', '#0d0d0d');

  // ─────────────────────────────────────────
  // セクション7: 施策実験ログ（今週試したこと・振り返り）
  // ─────────────────────────────────────────
  row = _writeSectionHeader(sh, row, '⑦ 施策実験ログ（今週試したこと・効果・振り返り）', '#1a0a2e');

  // ヘッダー行（施策ログ用）
  const expHeader = ['', '施策種別', '施策内容・仮説', '計測指標', '結果', '判定'];
  sh.getRange(row, 1, 1, 6).setValues([expHeader]);
  sh.getRange(row, 1, 1, 6)
    .setBackground('#2d1b4e')
    .setFontColor('#c9b1ff')
    .setFontWeight('bold')
    .setFontSize(9);
  row++;

  const expRows = [
    ['', 'ギフティング',  '', '再生数 / 保存数 / フォロワー転換率', '', ''],
    ['', 'ギフティング',  '', '再生数 / 保存数 / フォロワー転換率', '', ''],
    ['', 'Meta広告',      '', 'ROAS / CTR / CPA',                  '', ''],
    ['', '自社リール',    '', '再生数 / 保存数 / フォロワー増減',   '', ''],
    ['', 'MD/商品',       '', '売上 / CV率 / 返品率',               '', ''],
  ];

  expRows.forEach((r, i) => {
    sh.getRange(row + i, 1, 1, 6).setValues([r]);
    sh.getRange(row + i, 1, 1, 6)
      .setBackground(i % 2 === 0 ? '#1a0d2e' : '#130a22')
      .setFontColor('#e0e0e0')
      .setFontSize(9);
    // 判定列（F）はドロップダウン
    const judgeCell = sh.getRange(row + i, 6);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['✅ 継続', '🔄 改善して継続', '⏸ 一時停止', '❌ 中止', '📊 計測中'], true)
      .build();
    judgeCell.setDataValidation(rule);
    // 施策種別（B列）もドロップダウン
    const typeCell = sh.getRange(row + i, 2);
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ギフティング', 'Meta広告', '自社リール', 'MD/商品', 'サイト/LP', 'その他'], true)
      .build();
    typeCell.setDataValidation(typeRule);
  });
  row += expRows.length;

  // 施策ログシートへの誘導メモ
  sh.getRange(row, 2, 1, 5).merge()
    .setValue('▶ 詳細な施策履歴・効果一覧は「施策ログ」シートを参照')
    .setFontColor('#888888')
    .setFontStyle('italic')
    .setFontSize(8);
  row++;

  // 区切り線
  sh.getRange(row, 1, 1, 6).setBackground('#333333');
  row++;
}

// ============================================================
// 金曜進捗 エントリ本体
// ============================================================
function _appendFridayEntry(sh, date) {
  const lastRow = sh.getLastRow();
  let row = lastRow + (lastRow > 1 ? 2 : 1);

  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd (E)');

  sh.getRange(row, 1, 1, 5).merge()
    .setValue(`⚡ 金曜進捗ミーティング  ${dateStr}`)
    .setBackground('#0f3460')
    .setFontColor('#e94560')
    .setFontWeight('bold')
    .setFontSize(12);
  row++;

  row = _writeSectionHeader(sh, row, '① アクション進捗確認', '#16213e');

  const actionRows = [
    ['', 'アクション', '担当', 'ステータス', 'メモ'],
    ['', '（火曜から転記）', '', '✅ 完了 / 🔄 対応中 / ⏸ 保留', ''],
    ['', '', 'yayoi', '', ''],
    ['', '', 'shoi',  '', ''],
    ['', '', 'MD',    '', ''],
  ];
  row = _writeDataRows(sh, row, actionRows, '#1a1a2e', '#0d0d0d');

  row = _writeSectionHeader(sh, row, '② ブロッカー・課題', '#16213e');
  for (let i = 0; i < 3; i++) {
    sh.getRange(row, 2, 1, 4).merge().setValue('');
    sh.getRange(row, 1, 1, 5).setBackground('#0d0d0d');
    row++;
  }

  row = _writeSectionHeader(sh, row, '③ 翌週準備チェック', '#16213e');
  const checkRows = [
    ['', '来週のギフティングDM 準備済み？', '', '[ ] Yes  [ ] No', ''],
    ['', '広告クリエイティブ 2本以上確保？', '', '[ ] Yes  [ ] No', ''],
    ['', '在庫確認（オリジナル）完了？',    '', '[ ] Yes  [ ] No', ''],
    ['', '翌週ミーティングアジェンダ作成？', '', '[ ] Yes  [ ] No', ''],
  ];
  row = _writeDataRows(sh, row, checkRows, '#1a1a2e', '#0d0d0d');

  sh.getRange(row, 1, 1, 5).setBackground('#333333');
}

// ============================================================
// データ取得: Meta広告（先週分）
// ============================================================
function _fetchMetaAdsWeekly(fromDate, toDate) {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.META_ADS);
    const sh   = ss.getSheets()[0]; // 日次データシート
    const data = sh.getDataRange().getValues();

    // ヘッダー行を特定
    let headerRow = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).includes('日付') || String(data[i][0]).includes('date') || String(data[i][1]).includes('消化')) {
        headerRow = i;
        break;
      }
    }
    if (headerRow < 0) return _emptyAds();

    // 列インデックスを特定
    const headers = data[headerRow].map(h => String(h).trim());
    const col = {
      date:  _findCol(headers, ['日付', 'date']),
      spend: _findCol(headers, ['消化金額', 'spend']),
      roas:  _findCol(headers, ['ROAS', 'roas']),
      cpa:   _findCol(headers, ['CPA', 'cpa']),
      ctr:   _findCol(headers, ['CTR', 'ctr']),
      cpm:   _findCol(headers, ['CPM', 'cpm']),
      cv:    _findCol(headers, ['CV数', 'cv']),
      cvSales: _findCol(headers, ['CV売上', 'cv_sales', '売上']),
    };

    // 対象期間の行を集計
    let spend = 0, roasSum = 0, cpaSum = 0, ctrSum = 0, cpmSum = 0, cv = 0, cvSales = 0, count = 0;
    for (let i = headerRow + 1; i < data.length; i++) {
      const rowDate = _toDateStr(data[i][col.date]);
      if (rowDate >= fromDate && rowDate <= toDate) {
        spend   += _toNum(data[i][col.spend]);
        cv      += _toNum(data[i][col.cv]);
        cvSales += _toNum(data[i][col.cvSales]);
        roasSum += _toNum(data[i][col.roas]);
        cpaSum  += _toNum(data[i][col.cpa]);
        ctrSum  += _toNum(data[i][col.ctr]);
        cpmSum  += _toNum(data[i][col.cpm]);
        count++;
      }
    }
    if (count === 0) return _emptyAds();

    return {
      spend:    spend,
      roas:     roasSum / count,
      cpa:      spend > 0 && cv > 0 ? spend / cv : 0,
      ctr:      ctrSum / count,
      cpm:      cpmSum / count,
      cv:       cv,
      cvSales:  cvSales,
      spendWoW: '—',
      roasWoW:  '—',
      cpaWoW:   '—',
    };
  } catch (e) {
    Logger.log('Meta広告データ取得エラー: ' + e.message);
    return _emptyAds();
  }
}

function _emptyAds() {
  return { spend: null, roas: null, cpa: null, ctr: null, cpm: null, cv: null, cvSales: null,
           spendWoW: '—', roasWoW: '—', cpaWoW: '—' };
}

// ============================================================
// データ取得: Shopify（先週分）
// ============================================================
function _fetchShopifyWeekly(fromDate, toDate) {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.MARKETING);
    const sh   = ss.getSheets()[0];
    const data = sh.getDataRange().getValues();

    let headerRow = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).includes('日付') || String(data[i][1]).includes('売上')) {
        headerRow = i;
        break;
      }
    }
    if (headerRow < 0) return _emptyShop();

    const headers = data[headerRow].map(h => String(h).trim());
    const col = {
      date:      _findCol(headers, ['日付', 'date']),
      netSales:  _findCol(headers, ['純売上', 'net_sales', '売上高']),
      orders:    _findCol(headers, ['注文数', 'orders']),
      aovNew:    _findCol(headers, ['新規.*客単価', '新規顧客.*単価', 'aov_new']),
      aovReturn: _findCol(headers, ['リピ.*客単価', '既存.*単価', 'aov_return']),
      cvr:       _findCol(headers, ['CV率', 'cvr', 'conversion']),
      newOrders: _findCol(headers, ['新規.*注文', '新規顧客.*注文']),
      retOrders: _findCol(headers, ['リピ.*注文', '既存.*注文']),
    };

    let netSales = 0, orders = 0, aovNewSum = 0, aovRetSum = 0, cvrSum = 0, newOrd = 0, retOrd = 0, count = 0;
    for (let i = headerRow + 1; i < data.length; i++) {
      const rowDate = _toDateStr(data[i][col.date]);
      if (rowDate >= fromDate && rowDate <= toDate) {
        netSales  += _toNum(data[i][col.netSales]);
        orders    += _toNum(data[i][col.orders]);
        aovNewSum += _toNum(data[i][col.aovNew]);
        aovRetSum += _toNum(data[i][col.aovReturn]);
        cvrSum    += _toNum(data[i][col.cvr]);
        newOrd    += _toNum(data[i][col.newOrders]);
        retOrd    += _toNum(data[i][col.retOrders]);
        count++;
      }
    }
    if (count === 0) return _emptyShop();

    const totalOrd = newOrd + retOrd;
    return {
      netSales:    netSales,
      orders:      orders,
      aovNew:      count > 0 ? aovNewSum / count : null,
      aovReturn:   count > 0 ? aovRetSum / count : null,
      cvr:         count > 0 ? cvrSum / count    : null,
      newVsReturn: totalOrd > 0 ? `新規 ${Math.round(newOrd/totalOrd*100)}% / リピ ${Math.round(retOrd/totalOrd*100)}%` : '—',
      salesWoW:    '—',
      ordersWoW:   '—',
    };
  } catch (e) {
    Logger.log('Shopifyデータ取得エラー: ' + e.message);
    return _emptyShop();
  }
}

function _emptyShop() {
  return { netSales: null, orders: null, aovNew: null, aovReturn: null, cvr: null, newVsReturn: '—', salesWoW: '—', ordersWoW: '—' };
}

// ============================================================
// データ取得: ギフティング（今月分）
// ============================================================
function _fetchGiftingMonthly() {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.INFLUENCER);
    const sh   = ss.getSheetByName('マスター') || ss.getSheets()[0];
    const data = sh.getDataRange().getValues();

    // 今月の集計: ステータスカラムを確認
    const now   = new Date();
    const thisY = now.getFullYear();
    const thisM = now.getMonth();

    // ヘッダー行を特定
    let headerRow = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).includes('No') || String(data[i][3]).includes('ID')) {
        headerRow = i;
        break;
      }
    }
    if (headerRow < 0) return _emptyGifting();

    const headers = data[headerRow].map(h => String(h).trim());
    const col = {
      dmDate:   _findCol(headers, ['DM日', 'dm_date', 'dm日']),
      status:   _findCol(headers, ['ステータス', 'status']),
      posted:   _findCol(headers, ['投稿日', '掲載日']),
      gifted:   _findCol(headers, ['ギフティング', '提供', '発送']),
    };

    let dmCount = 0, replyCount = 0, giftingDone = 0, posted = 0;
    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      // DM日が今月かどうか確認
      const dmDateRaw = row[col.dmDate];
      if (dmDateRaw) {
        const dmDate = new Date(dmDateRaw);
        if (!isNaN(dmDate) && dmDate.getFullYear() === thisY && dmDate.getMonth() === thisM) {
          dmCount++;
          const status = String(row[col.status] || '');
          // 返答あり = ステータス1以上
          if (status.includes('1') || status.includes('2') || status.includes('3') || status.includes('4') ||
              status.includes('返答') || status.includes('承認') || status.includes('完了')) {
            replyCount++;
          }
        }
      }
      // ギフティング実施数・投稿数は全期間の今月分
      const giftedRaw = row[col.gifted];
      if (giftedRaw) {
        const giftedDate = new Date(giftedRaw);
        if (!isNaN(giftedDate) && giftedDate.getFullYear() === thisY && giftedDate.getMonth() === thisM) {
          giftingDone++;
        }
      }
      const postedRaw = row[col.posted];
      if (postedRaw) {
        const postedDate = new Date(postedRaw);
        if (!isNaN(postedDate) && postedDate.getFullYear() === thisY && postedDate.getMonth() === thisM) {
          posted++;
        }
      }
    }

    return {
      dmCount:    dmCount,
      replyCount: replyCount,
      replyRate:  dmCount > 0 ? replyCount / dmCount : 0,
      giftingDone: giftingDone,
      posted:     posted,
    };
  } catch (e) {
    Logger.log('ギフティングデータ取得エラー: ' + e.message);
    return _emptyGifting();
  }
}

function _emptyGifting() {
  return { dmCount: null, replyCount: null, replyRate: null, giftingDone: null, posted: null };
}

// ============================================================
// ステータス判定（色つきコメント）
// ============================================================
function _roasStatus(roas) {
  if (roas === null || isNaN(roas)) return '';
  if (roas >= TARGETS.ROAS_IDEAL)               return '🔵 理想達成（700%+）→ 予算拡大';
  if (roas >= (TARGETS.ROAS_IDEAL + TARGETS.ROAS_MIN) / 2) return '🟢 好調（600%+）→ 現状維持';
  if (roas >= TARGETS.ROAS_MIN)                 return '🟡 最低ライン達成（500%+）→ 改善余地あり';
  if (roas >= TARGETS.ROAS_MIN * 0.8)           return '🟠 要注意（400%台）→ 素材・ターゲ見直し';
  return '🔴 要対応（500%未満）→ 即改善アクション';
}

function _cpaStatus(cpa) {
  if (cpa === null || isNaN(cpa) || cpa === 0) return '';
  if (cpa <= TARGETS.CPA * 0.8)  return '🔵 優秀';
  if (cpa <= TARGETS.CPA)        return '🟢 目標内';
  if (cpa <= TARGETS.CPA * 1.3)  return '🟡 要改善';
  return '🔴 要対応（CPA超過）';
}

function _giftingStatus(done) {
  if (done === null || isNaN(done)) return '';
  const now  = new Date();
  const daysInMonth  = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const daysPassed   = now.getDate();
  const expected = Math.round(TARGETS.GIFTING_PER_MONTH * daysPassed / daysInMonth);
  if (done >= TARGETS.GIFTING_PER_MONTH) return '🟢 月目標達成';
  if (done >= expected) return '🟢 ペース良好';
  if (done >= expected * 0.7) return '🟡 やや遅れ';
  return '🔴 要加速';
}

// ============================================================
// UI ヘルパー
// ============================================================
function _writeSectionHeader(sh, row, title, bgColor) {
  sh.getRange(row, 1, 1, 6).merge()
    .setValue(title)
    .setBackground(bgColor || '#16213e')
    .setFontColor('#a0c4ff')
    .setFontWeight('bold')
    .setFontSize(10);
  return row + 1;
}

function _writeDataRows(sh, row, rows, evenBg, oddBg) {
  rows.forEach((r, i) => {
    const range = sh.getRange(row + i, 1, 1, 6);
    range.setValues([r]);
    range.setBackground(i % 2 === 0 ? (evenBg || '#1a1a2e') : (oddBg || '#0d0d0d'));
    range.setFontColor('#e0e0e0');
    range.setFontSize(9);

    // 目標値列（E列）は色を変える
    sh.getRange(row + i, 5).setFontColor('#ffd166');
  });
  return row + rows.length;
}

// ============================================================
// フォーマット ヘルパー
// ============================================================
function _fmt(val) {
  if (val === null || val === undefined || isNaN(val) || val === 0) return '取得中...';
  return '¥' + Math.round(val).toLocaleString();
}

function _pct(val) {
  if (val === null || val === undefined || isNaN(val)) return '取得中...';
  // 0〜1の小数 or すでに%表記の数値
  const v = val > 1 ? val : val * 100;
  return v.toFixed(2) + '%';
}

function _num(val, decimals) {
  if (val === null || val === undefined || isNaN(val)) return '取得中...';
  return val.toFixed(decimals !== undefined ? decimals : 0);
}

function _toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(/[,¥%]/g, ''));
  return isNaN(n) ? 0 : n;
}

function _toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v).substring(0, 10);
}

function _findCol(headers, candidates) {
  for (const cand of candidates) {
    const re = new RegExp(cand, 'i');
    const idx = headers.findIndex(h => re.test(h));
    if (idx >= 0) return idx;
  }
  return 0;
}

// ============================================================
// 対象週の計算（月〜日）
// ============================================================
function _getWeekRange(date) {
  // 直近の月曜日（先週月曜 = 今週の始まりを計算）
  const d   = new Date(date);
  const dow = d.getDay(); // 0=日
  // 今日が月曜(1)なら先週月曜 = -7日前
  const lastMon = new Date(d);
  lastMon.setDate(d.getDate() - (dow === 0 ? 7 : dow) - 7 + 1);
  const lastSun = new Date(lastMon);
  lastSun.setDate(lastMon.getDate() + 6);

  const fmt = dt => Utilities.formatDate(dt, 'Asia/Tokyo', 'MM/dd');
  return {
    from:  Utilities.formatDate(lastMon, 'Asia/Tokyo', 'yyyy-MM-dd'),
    to:    Utilities.formatDate(lastSun, 'Asia/Tokyo', 'yyyy-MM-dd'),
    label: `${fmt(lastMon)}（月）〜${fmt(lastSun)}（日）`,
  };
}

// ============================================================
// 施策ログシート セットアップ（蓄積型・全施策の一覧台帳）
// ============================================================
function setup施策LogSheet(ss) {
  const sh = ss.getSheetByName(SHEET_施策LOG);
  sh.clearContents();
  sh.clearFormats();

  // 列幅
  sh.setColumnWidth(1, 100);   // A: 記録日
  sh.setColumnWidth(2, 110);   // B: 施策種別
  sh.setColumnWidth(3, 240);   // C: 施策内容・仮説
  sh.setColumnWidth(4, 200);   // D: 計測指標・結果
  sh.setColumnWidth(5, 100);   // E: 数値（ROAS等）
  sh.setColumnWidth(6, 130);   // F: 判定
  sh.setColumnWidth(7, 280);   // G: 振り返り・学び
  sh.setColumnWidth(8, 180);   // H: 次のアクション

  // ヘッダー
  const headers = ['記録日', '施策種別', '施策内容・仮説', '計測指標・結果', '数値', '判定', '振り返り・学び', '次のアクション'];
  sh.getRange(1, 1, 1, 8).setValues([headers]);
  sh.getRange(1, 1, 1, 8)
    .setBackground('#2d1b4e')
    .setFontColor('#c9b1ff')
    .setFontWeight('bold')
    .setFontSize(10);
  sh.setFrozenRows(1);

  // サンプル行（使い方ガイド）
  const guideRows = [
    ['使い方', '毎週⑦で記録 → ここに転記 → 月次レビューで傾向を見る', '', '', '', '', '', ''],
    ['─── ギフティング例 ───', '', '', '', '', '', '', ''],
    [
      '2026/04/15', 'ギフティング',
      '【仮説】ファッション特化クリエイター（5〜10万）は購買転換率が高い',
      '対象5名 平均再生数 / CV数 / ROAS比較',
      'ROAS 6.2', '🔄 改善して継続',
      '再生10万超えの1本がCV全体の60%牽引。フォロワー数より動画クオリティが重要',
      '次月もファッション特化・動画クオリティ重視で選定継続'
    ],
    ['─── Meta広告例 ───', '', '', '', '', '', '', ''],
    [
      '2026/04/08', 'Meta広告',
      '【仮説】UGC素材 vs 自社制作素材でCTR比較',
      'CTR / CPA / ROAS（7日間テスト）',
      'UGC: CTR4.2% / 自社: CTR2.1%', '✅ 継続',
      'UGCのCTRが2倍。自社制作は商品説明に絞る',
      'ギフティング動画を積極的に2次利用。自社制作は商品ハイライトのみ'
    ],
  ];

  sh.getRange(2, 1, guideRows.length, 8).setValues(guideRows);
  // ガイド行はグレー
  sh.getRange(2, 1, 1, 8).setBackground('#1a0a2e').setFontColor('#666666').setFontSize(8);
  sh.getRange(3, 1, 1, 8).setBackground('#1a0a2e').setFontColor('#888888').setFontWeight('bold').setFontSize(8);
  sh.getRange(4, 1, 1, 8).setBackground('#1a0d2e').setFontColor('#d0d0d0').setFontSize(9);
  sh.getRange(5, 1, 1, 8).setBackground('#1a0a2e').setFontColor('#888888').setFontWeight('bold').setFontSize(8);
  sh.getRange(6, 1, 1, 8).setBackground('#130a22').setFontColor('#d0d0d0').setFontSize(9);

  // 入力行（7行目〜）にドロップダウン設定
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ギフティング', 'Meta広告', '自社リール', 'MD/商品', 'サイト/LP', 'その他'], true)
    .build();
  const judgeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ 継続', '🔄 改善して継続', '⏸ 一時停止', '❌ 中止', '📊 計測中'], true)
    .build();
  sh.getRange(7, 2, 50, 1).setDataValidation(typeRule);
  sh.getRange(7, 6, 50, 1).setDataValidation(judgeRule);
  sh.getRange(7, 1, 50, 8).setBackground('#0d0d0d').setFontColor('#e0e0e0').setFontSize(9);

  SpreadsheetApp.flush();
  Logger.log('✅ 施策ログシート セットアップ完了');
}

// ============================================================
// シート取得or作成
// ============================================================
function _getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  return sh;
}

// ============================================================
// トリガー設定
// ============================================================
function setupTriggers() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'createWeeklyMeeting' ||
        t.getHandlerFunction() === 'createFridayMeeting') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 月曜 8:00 → 火曜ミーティング用エントリ作成
  ScriptApp.newTrigger('createWeeklyMeeting')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  // 木曜 8:00 → 金曜進捗ミーティング用エントリ作成
  ScriptApp.newTrigger('createFridayMeeting')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(8)
    .create();

  Logger.log('✅ トリガー設定完了');
}
