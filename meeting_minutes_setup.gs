// ============================================================
// 週次ミーティング 議事録スクリプト v3.0（月別タブ方式）
//
// タブ構成:
//   2026年4月, 2026年5月 ... → 月ごとに自動作成（火曜+金曜を縦積み）
//   施策ログ                 → 全期間蓄積型
//   月次サマリー             → 月ごとのKPIまとめ
//
// トリガー:
//   月曜 8:00 → createWeeklyMeeting()  火曜ミーティング枠を追記
//   木曜 8:00 → createFridayMeeting()  金曜進捗枠を追記
//   毎月1日 8:00 → createMonthSheet()  翌月タブを先行作成
// ============================================================

// ─── 他シートのスプレッドシートID ───
const SPREADSHEET_IDS = {
  META_ADS:   '1kd6HD4NOECKNl4knsklsrcUDWAOs-ve3xhULEtKrEt8',
  INFLUENCER: '1prt8F8EAoHlGlkIv6uLkTGXZwIUD13-_QDtIdvEOMIE',
  MARKETING:  '1fQg_01ViGrZEoT38LifM_zh-U4vah_8ORE4xTodz708',
};

// ─── 目標値 ───
const TARGETS = {
  ROAS_IDEAL:        7.0,   // 理想（700%）
  ROAS_MIN:          5.0,   // 最低ライン（500%）
  ROAS:              5.0,
  CPA:               2000,
  GIFTING_PER_MONTH: 10,
  DM_PER_MONTH:      35,
  CREATIVE_PER_WEEK: 2,
};

// ─── 固定シート名 ───
const SHEET_施策LOG  = '施策ログ';
const SHEET_MONTHLY = '月次サマリー';

// ============================================================
// 初期セットアップ（一度だけ手動実行）
// ============================================================
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 今月タブ作成
  const monthSh = _getOrCreateMonthSheet(ss, new Date());

  // 固定シート作成
  _getOrCreateSheet(ss, SHEET_施策LOG);
  _getOrCreateSheet(ss, SHEET_MONTHLY);

  // 各シートを初期化
  _initMonthSheet(monthSh, new Date());
  setup施策LogSheet(ss);
  setupMonthlySheet(ss);

  // 今月の第1週エントリを挿入
  const week = _getWeekRange(new Date());
  _appendWeeklyEntry(monthSh, week);

  setupTriggers();
  Logger.log('✅ セットアップ完了: ' + _monthLabel(new Date()) + 'タブ + 施策ログ + 月次サマリー');
}

// ============================================================
// 毎週月曜 8:00: 火曜ミーティング枠を追記
// ============================================================
function createWeeklyMeeting() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = _getOrCreateMonthSheet(ss, new Date());
  const week = _getWeekRange(new Date());
  _appendWeeklyEntry(sh, week);
  Logger.log('✅ 火曜ミーティング枠追加: ' + week.label);
}

// ============================================================
// 毎週木曜 8:00: 金曜進捗枠を追記
// ============================================================
function createFridayMeeting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = _getOrCreateMonthSheet(ss, new Date());
  _appendFridayEntry(sh, new Date());
  Logger.log('✅ 金曜進捗枠追加');
}

// ============================================================
// 毎月1日 8:00: 翌月タブを先行作成（月初に慌てないよう）
// ============================================================
function createMonthSheet() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const next = new Date();
  next.setMonth(next.getMonth() + 1);
  const sh = _getOrCreateMonthSheet(ss, next);
  Logger.log('✅ 翌月タブ作成: ' + _monthLabel(next));
}

// ============================================================
// 月別タブ: 初期ヘッダー設定
// ============================================================
function _initMonthSheet(sh, date) {
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 28);
  sh.setColumnWidth(2, 185);
  sh.setColumnWidth(3, 125);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 160);
  sh.setColumnWidth(6, 280);

  // 月タイトル行
  const label = _monthLabel(date);
  sh.getRange(1, 1, 1, 6).merge()
    .setValue('📅 ' + label + '  議事録')
    .setBackground('#0a0a1a')
    .setFontColor('#e94560')
    .setFontWeight('bold')
    .setFontSize(14);

  // 使い方ガイド
  sh.getRange(2, 1, 1, 6).merge()
    .setValue('火曜ミーティング（毎週月曜8:00に自動追記）と金曜進捗（毎週木曜8:00に自動追記）が下に積み上がります')
    .setBackground('#0a0a1a')
    .setFontColor('#666666')
    .setFontSize(8);

  sh.setFrozenRows(2);
  SpreadsheetApp.flush();
}

// ============================================================
// 火曜ミーティング エントリ
// ============================================================
function _appendWeeklyEntry(sh, week) {
  let row = Math.max(sh.getLastRow() + 2, 3);

  // ── データ取得 ──
  const ads  = _fetchMetaAdsWeekly(week.from, week.to);
  const shop = _fetchShopifyWeekly(week.from, week.to);
  const inf  = _fetchGiftingMonthly();

  // ─── タイトル ───
  sh.getRange(row, 1, 1, 6).merge()
    .setValue('🗓️ 火曜ミーティング   ' + week.label)
    .setBackground('#0f3460')
    .setFontColor('#e94560')
    .setFontWeight('bold')
    .setFontSize(11);
  row++;

  // ─── ① 売上サマリー ───
  row = _writeSectionHeader(sh, row, '① 売上サマリー（先週）');
  row = _writeDataRows(sh, row, [
    ['', '週間純売上',       _fmt(shop.netSales),   shop.salesWoW,  '—',  ''],
    ['', '注文数',           _fmt0(shop.orders),    shop.ordersWoW, '—',  ''],
    ['', '客単価・新規',     _fmt(shop.aovNew),     '—',            '—',  ''],
    ['', '客単価・リピ',     _fmt(shop.aovReturn),  '—',            '—',  ''],
    ['', 'CV率',             _pct(shop.cvr),        '—',            '—',  ''],
    ['', '新規 / リピ比',    shop.newVsReturn,      '—',            '—',  ''],
  ]);

  // ─── ② Meta広告 ───
  row = _writeSectionHeader(sh, row, '② Meta広告（先週）');
  row = _writeDataRows(sh, row, [
    ['', '消化金額',           _fmt(ads.spend),      ads.spendWoW, '—',                                              ''],
    ['', 'ROAS',               _num(ads.roas, 2),    ads.roasWoW,  `理想 ${TARGETS.ROAS_IDEAL*100}% / 最低 ${TARGETS.ROAS_MIN*100}%`, _roasStatus(ads.roas)],
    ['', 'CPA',                _fmt(ads.cpa),        ads.cpaWoW,   `目標 ¥${TARGETS.CPA}以下`,                       _cpaStatus(ads.cpa)],
    ['', 'CTR',                _pct(ads.ctr),        '—',          '—',                                              ''],
    ['', 'CPM',                _fmt(ads.cpm),        '—',          '—',                                              ''],
    ['', 'CV数',               _fmt0(ads.cv),        '—',          '—',                                              ''],
    ['', 'CV売上',             _fmt(ads.cvSales),    '—',          '—',                                              ''],
    ['', '追加クリエイティブ数', '（手入力）',        '—',          `週${TARGETS.CREATIVE_PER_WEEK}本+`,              ''],
  ]);

  // ─── ③ ギフティング ───
  row = _writeSectionHeader(sh, row, '③ ギフティング（今月進捗）');
  row = _writeDataRows(sh, row, [
    ['', 'DM送信数',            _fmt0(inf.dmCount),      '—', `目標 ${TARGETS.DM_PER_MONTH}件`,      ''],
    ['', '返答数',              _fmt0(inf.replyCount),   '—', '—',                                   ''],
    ['', '返答率',              _pct(inf.replyRate),     '—', '—',                                   ''],
    ['', 'ギフティング実施数',  _fmt0(inf.giftingDone),  '—', `目標 ${TARGETS.GIFTING_PER_MONTH}件`,  _giftingStatus(inf.giftingDone)],
    ['', '投稿済み数',          _fmt0(inf.posted),       '—', '—',                                   ''],
    ['', '新規素材（広告転用可）', '（手入力）',         '—', '—',                                   '広告転用した本数'],
  ]);

  // ─── ④ 今週の投稿予定 ───
  row = _writeSectionHeader(sh, row, '④ 今週の投稿予定');
  sh.getRange(row, 2, 1, 5).setValues([['@インスタID', '投稿予定日', 'ジャンル', '品番', '備考']])
    .setFontColor('#888888').setFontSize(8).setBackground('#111111');
  row++;
  row = _writeDataRows(sh, row, [
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ]);

  // ─── ⑤ 課題・懸念事項 ───
  row = _writeSectionHeader(sh, row, '⑤ 課題・懸念事項');
  sh.getRange(row, 2, 1, 5).setValues([['課題内容', 'オーナー', '緊急度', '期限', 'ステータス']])
    .setFontColor('#888888').setFontSize(8).setBackground('#111111');
  row++;
  row = _writeDataRows(sh, row, [
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ]);

  // ─── ⑥ アクションアイテム ───
  row = _writeSectionHeader(sh, row, '⑥ アクションアイテム（今週）');
  sh.getRange(row, 2, 1, 5).setValues([['アクション内容', '担当', 'Due', 'ステータス', '']])
    .setFontColor('#888888').setFontSize(8).setBackground('#111111');
  row++;
  row = _writeDataRows(sh, row, [
    ['', '', 'yayoi', '', '[ ] 未', ''],
    ['', '', 'shoi',  '', '[ ] 未', ''],
    ['', '', 'MD',    '', '[ ] 未', ''],
  ]);

  // ─── ⑦ 施策実験ログ ───
  row = _writeSectionHeader(sh, row, '⑦ 施策実験ログ（今週試したこと）', '#1a0a2e', '#c9b1ff');
  sh.getRange(row, 2, 1, 5).setValues([['施策種別', '施策内容・仮説', '計測指標', '結果', '判定']])
    .setFontColor('#888888').setFontSize(8).setBackground('#130a22');
  row++;

  const expRows = [
    ['', 'ギフティング', '', '再生数 / 保存数', '', ''],
    ['', 'ギフティング', '', '再生数 / 保存数', '', ''],
    ['', 'Meta広告',    '', 'ROAS / CTR',       '', ''],
    ['', '自社リール',  '', '再生数 / 保存数',  '', ''],
  ];
  expRows.forEach((r, i) => {
    sh.getRange(row + i, 1, 1, 6).setValues([r])
      .setBackground(i % 2 === 0 ? '#1a0d2e' : '#130a22')
      .setFontColor('#e0e0e0').setFontSize(9);
    // 施策種別ドロップダウン
    sh.getRange(row + i, 2).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['ギフティング', 'Meta広告', '自社リール', 'MD/商品', 'その他'], true).build()
    );
    // 判定ドロップダウン
    sh.getRange(row + i, 6).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['✅ 継続', '🔄 改善して継続', '⏸ 一時停止', '❌ 中止', '📊 計測中'], true).build()
    );
  });
  row += expRows.length;

  sh.getRange(row, 2, 1, 5).merge()
    .setValue('▶ 詳細な施策履歴は「施策ログ」シートへ転記')
    .setFontColor('#555555').setFontStyle('italic').setFontSize(8).setBackground('#0d0d0d');
  row++;

  // 区切り線
  sh.getRange(row, 1, 1, 6).setBackground('#1e1e3a').setValues([['', '', '', '', '', '']]);
  row += 2; // 次エントリとの間隔
}

// ============================================================
// 金曜進捗 エントリ
// ============================================================
function _appendFridayEntry(sh, date) {
  let row = Math.max(sh.getLastRow() + 2, 3);
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd (E)');

  // タイトル
  sh.getRange(row, 1, 1, 6).merge()
    .setValue('⚡ 金曜進捗ミーティング   ' + dateStr)
    .setBackground('#0f3460')
    .setFontColor('#ffd166')
    .setFontWeight('bold')
    .setFontSize(11);
  row++;

  // ① アクション進捗
  row = _writeSectionHeader(sh, row, '① アクション進捗確認（火曜から転記）');
  sh.getRange(row, 2, 1, 4).setValues([['アクション内容', '担当', 'ステータス', 'メモ']])
    .setFontColor('#888888').setFontSize(8).setBackground('#111111');
  row++;
  row = _writeDataRows(sh, row, [
    ['', '', 'yayoi', '✅ 完了 / 🔄 対応中 / ⏸ 保留', ''],
    ['', '', 'shoi',  '', ''],
    ['', '', 'MD',    '', ''],
  ]);

  // ② ブロッカー
  row = _writeSectionHeader(sh, row, '② ブロッカー・課題');
  row = _writeDataRows(sh, row, [
    ['', '', '', '', ''],
    ['', '', '', '', ''],
  ]);

  // ③ 翌週準備チェック
  row = _writeSectionHeader(sh, row, '③ 翌週準備チェック');
  row = _writeDataRows(sh, row, [
    ['', '来週ギフティングDM 準備済み？',    '', '[ ] Yes  [ ] No', ''],
    ['', '広告クリエイティブ 2本以上確保？',  '', '[ ] Yes  [ ] No', ''],
    ['', '在庫確認（オリジナル）完了？',      '', '[ ] Yes  [ ] No', ''],
    ['', '翌週ミーティングアジェンダ作成？',  '', '[ ] Yes  [ ] No', ''],
  ]);

  // 区切り線
  sh.getRange(row, 1, 1, 6).setBackground('#1e1e3a').setValues([['', '', '', '', '', '']]);
}

// ============================================================
// 施策ログシート（蓄積型）
// ============================================================
function setup施策LogSheet(ss) {
  const sh = ss.getSheetByName(SHEET_施策LOG);
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 100);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 240);
  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 140);
  sh.setColumnWidth(7, 280);
  sh.setColumnWidth(8, 180);

  const headers = ['記録日', '施策種別', '施策内容・仮説', '計測指標・結果', '数値', '判定', '振り返り・学び', '次のアクション'];
  sh.getRange(1, 1, 1, 8).setValues([headers])
    .setBackground('#2d1b4e').setFontColor('#c9b1ff').setFontWeight('bold').setFontSize(10);
  sh.setFrozenRows(1);

  // サンプル行
  sh.getRange(2, 1, 1, 8).setValues([['使い方', '毎週⑦で記録 → ここに転記 → 月次レビューで傾向分析', '', '', '', '', '', '']])
    .setBackground('#1a0a2e').setFontColor('#666666').setFontSize(8);

  const samples = [
    ['2026/04/15', 'ギフティング', '【仮説】ファッション特化5〜10万フォロワーは購買転換率高い',
     '対象5名の平均再生数/CV数/ROAS比較', 'ROAS 6.2', '🔄 改善して継続',
     '再生10万超の1本がCV全体60%牽引。フォロワー数より動画クオリティが重要',
     '次月もファッション特化・クオリティ重視で選定'],
    ['2026/04/08', 'Meta広告', '【仮説】インフルエンサーUGC vs 自社制作でCTR比較',
     'CTR/CPA/ROAS（7日間）', 'UGC:CTR4.2% / 自社:2.1%', '✅ 継続',
     'UGCのCTRが2倍。自社制作は商品説明に絞る',
     'ギフティング動画を積極的に2次利用。自社制作は商品ハイライトのみ'],
  ];
  sh.getRange(3, 1, samples.length, 8).setValues(samples)
    .setBackground('#1a0d2e').setFontColor('#d0d0d0').setFontSize(9);

  // 入力行（5行目〜）にドロップダウン
  const typeRule  = SpreadsheetApp.newDataValidation().requireValueInList(['ギフティング', 'Meta広告', '自社リール', 'MD/商品', 'サイト/LP', 'その他'], true).build();
  const judgeRule = SpreadsheetApp.newDataValidation().requireValueInList(['✅ 継続', '🔄 改善して継続', '⏸ 一時停止', '❌ 中止', '📊 計測中'], true).build();
  sh.getRange(5, 2, 100, 1).setDataValidation(typeRule);
  sh.getRange(5, 6, 100, 1).setDataValidation(judgeRule);
  sh.getRange(5, 1, 100, 8).setBackground('#0d0d0d').setFontColor('#e0e0e0').setFontSize(9);

  Logger.log('✅ 施策ログシート セットアップ完了');
}

// ============================================================
// 月次サマリーシート
// ============================================================
function setupMonthlySheet(ss) {
  const sh = ss.getSheetByName(SHEET_MONTHLY);
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 100);

  const headers = ['指標', '今月実績', '先月実績', '前月比', '目標', '達成'];
  sh.getRange(1, 1, 1, 6).setValues([headers])
    .setBackground('#16213e').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);

  const kpis = [
    ['── 売上 ──', '', '', '', '', ''],
    ['月間純売上', '', '', '', '—', ''],
    ['注文数',     '', '', '', '—', ''],
    ['客単価・新規', '', '', '', '—', ''],
    ['新規/リピ比', '', '', '', '—', ''],
    ['── Meta広告 ──', '', '', '', '', ''],
    ['月間消化金額', '', '', '', '—', ''],
    ['月間ROAS',    '', '', '', `理想${TARGETS.ROAS_IDEAL*100}% / 最低${TARGETS.ROAS_MIN*100}%`, ''],
    ['月間CPA',     '', '', '', `¥${TARGETS.CPA}以下`, ''],
    ['── ギフティング ──', '', '', '', '', ''],
    ['DM送信数',        '', '', '', `${TARGETS.DM_PER_MONTH}件`, ''],
    ['ギフティング実施数', '', '', '', `${TARGETS.GIFTING_PER_MONTH}件`, ''],
    ['投稿済み数',       '', '', '', '—', ''],
    ['広告転用数',       '', '', '', `月${TARGETS.CREATIVE_PER_WEEK*4}本+`, ''],
    ['── 施策振り返り ──', '', '', '', '', ''],
    ['今月の勝ちパターン', '', '', '', '', ''],
    ['今月の改善点',     '', '', '', '', ''],
    ['来月に試すこと',   '', '', '', '', ''],
  ];
  kpis.forEach((r, i) => {
    sh.getRange(i + 2, 1, 1, 6).setValues([r]);
    if (String(r[0]).startsWith('──')) {
      sh.getRange(i + 2, 1, 1, 6).merge()
        .setBackground('#16213e').setFontColor('#a0c4ff').setFontWeight('bold').setFontSize(9);
    } else {
      sh.getRange(i + 2, 1, 1, 6)
        .setBackground(i % 2 === 0 ? '#1a1a2e' : '#0d0d0d')
        .setFontColor('#e0e0e0').setFontSize(9);
      sh.getRange(i + 2, 5).setFontColor('#ffd166');
    }
  });

  Logger.log('✅ 月次サマリーシート セットアップ完了');
}

// ============================================================
// データ取得: Meta広告（先週分）
// ============================================================
function _fetchMetaAdsWeekly(fromDate, toDate) {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.META_ADS);
    const sh   = ss.getSheets()[0];
    const data = sh.getDataRange().getValues();

    let headerRow = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).includes('日付') || String(data[i][1]).includes('消化')) { headerRow = i; break; }
    }
    if (headerRow < 0) return _emptyAds();

    const h = data[headerRow].map(v => String(v).trim());
    const c = {
      date:    _fc(h, ['日付']),
      spend:   _fc(h, ['消化金額', 'spend']),
      roas:    _fc(h, ['ROAS']),
      cpa:     _fc(h, ['CPA']),
      ctr:     _fc(h, ['CTR']),
      cpm:     _fc(h, ['CPM']),
      cv:      _fc(h, ['CV数']),
      cvSales: _fc(h, ['CV売上']),
    };

    let spend = 0, roasSum = 0, ctrSum = 0, cpmSum = 0, cv = 0, cvSales = 0, cnt = 0;
    for (let i = headerRow + 1; i < data.length; i++) {
      const d = _toDateStr(data[i][c.date]);
      if (d >= fromDate && d <= toDate) {
        spend   += _n(data[i][c.spend]);
        cv      += _n(data[i][c.cv]);
        cvSales += _n(data[i][c.cvSales]);
        roasSum += _n(data[i][c.roas]);
        ctrSum  += _n(data[i][c.ctr]);
        cpmSum  += _n(data[i][c.cpm]);
        cnt++;
      }
    }
    if (cnt === 0) return _emptyAds();
    return {
      spend: spend, roas: roasSum / cnt,
      cpa: cv > 0 ? spend / cv : 0,
      ctr: ctrSum / cnt, cpm: cpmSum / cnt,
      cv: cv, cvSales: cvSales,
      spendWoW: '—', roasWoW: '—', cpaWoW: '—',
    };
  } catch (e) {
    Logger.log('Meta広告取得エラー: ' + e.message);
    return _emptyAds();
  }
}
function _emptyAds() {
  return { spend:null, roas:null, cpa:null, ctr:null, cpm:null, cv:null, cvSales:null, spendWoW:'—', roasWoW:'—', cpaWoW:'—' };
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
      if (String(data[i][0]).includes('日付') || String(data[i][1]).includes('売上')) { headerRow = i; break; }
    }
    if (headerRow < 0) return _emptyShop();

    const h = data[headerRow].map(v => String(v).trim());
    const c = {
      date:      _fc(h, ['日付']),
      netSales:  _fc(h, ['純売上', '売上高']),
      orders:    _fc(h, ['注文数']),
      aovNew:    _fc(h, ['新規.*単価', '新規.*客単']),
      aovReturn: _fc(h, ['リピ.*単価', '既存.*単価']),
      cvr:       _fc(h, ['CV率', 'cvr']),
      newOrders: _fc(h, ['新規.*注文']),
      retOrders: _fc(h, ['リピ.*注文', '既存.*注文']),
    };

    let netSales = 0, orders = 0, aovNS = 0, aovRS = 0, cvrS = 0, newO = 0, retO = 0, cnt = 0;
    for (let i = headerRow + 1; i < data.length; i++) {
      const d = _toDateStr(data[i][c.date]);
      if (d >= fromDate && d <= toDate) {
        netSales += _n(data[i][c.netSales]);
        orders   += _n(data[i][c.orders]);
        aovNS    += _n(data[i][c.aovNew]);
        aovRS    += _n(data[i][c.aovReturn]);
        cvrS     += _n(data[i][c.cvr]);
        newO     += _n(data[i][c.newOrders]);
        retO     += _n(data[i][c.retOrders]);
        cnt++;
      }
    }
    if (cnt === 0) return _emptyShop();
    const tot = newO + retO;
    return {
      netSales: netSales, orders: orders,
      aovNew: cnt > 0 ? aovNS / cnt : null,
      aovReturn: cnt > 0 ? aovRS / cnt : null,
      cvr: cnt > 0 ? cvrS / cnt : null,
      newVsReturn: tot > 0 ? `新規 ${Math.round(newO/tot*100)}% / リピ ${Math.round(retO/tot*100)}%` : '—',
      salesWoW: '—', ordersWoW: '—',
    };
  } catch (e) {
    Logger.log('Shopify取得エラー: ' + e.message);
    return _emptyShop();
  }
}
function _emptyShop() {
  return { netSales:null, orders:null, aovNew:null, aovReturn:null, cvr:null, newVsReturn:'—', salesWoW:'—', ordersWoW:'—' };
}

// ============================================================
// データ取得: ギフティング（今月分）
// ============================================================
function _fetchGiftingMonthly() {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.INFLUENCER);
    const sh   = ss.getSheetByName('マスター') || ss.getSheets()[0];
    const data = sh.getDataRange().getValues();
    const now  = new Date();
    const thisY = now.getFullYear(), thisM = now.getMonth();

    let headerRow = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).includes('No') || String(data[i][3]).includes('ID')) { headerRow = i; break; }
    }
    if (headerRow < 0) return _emptyGifting();

    const h = data[headerRow].map(v => String(v).trim());
    const c = {
      dmDate:  _fc(h, ['DM日']),
      status:  _fc(h, ['ステータス']),
      gifted:  _fc(h, ['ギフティング', '提供', '発送']),
      posted:  _fc(h, ['投稿日', '掲載日']),
    };

    let dm = 0, reply = 0, gifted = 0, posted = 0;
    for (let i = headerRow + 1; i < data.length; i++) {
      const r = data[i];
      const dmD = r[c.dmDate] ? new Date(r[c.dmDate]) : null;
      if (dmD && !isNaN(dmD) && dmD.getFullYear() === thisY && dmD.getMonth() === thisM) {
        dm++;
        const st = String(r[c.status] || '');
        if (/[1234]|返答|承認|完了/.test(st)) reply++;
      }
      const gD = r[c.gifted] ? new Date(r[c.gifted]) : null;
      if (gD && !isNaN(gD) && gD.getFullYear() === thisY && gD.getMonth() === thisM) gifted++;
      const pD = r[c.posted] ? new Date(r[c.posted]) : null;
      if (pD && !isNaN(pD) && pD.getFullYear() === thisY && pD.getMonth() === thisM) posted++;
    }
    return { dmCount: dm, replyCount: reply, replyRate: dm > 0 ? reply / dm : 0, giftingDone: gifted, posted };
  } catch (e) {
    Logger.log('ギフティング取得エラー: ' + e.message);
    return _emptyGifting();
  }
}
function _emptyGifting() {
  return { dmCount:null, replyCount:null, replyRate:null, giftingDone:null, posted:null };
}

// ============================================================
// ステータス判定
// ============================================================
function _roasStatus(roas) {
  if (!roas) return '';
  if (roas >= TARGETS.ROAS_IDEAL)                                    return '🔵 理想達成（700%+）→ 予算拡大';
  if (roas >= (TARGETS.ROAS_IDEAL + TARGETS.ROAS_MIN) / 2)          return '🟢 好調（600%+）→ 現状維持';
  if (roas >= TARGETS.ROAS_MIN)                                      return '🟡 最低ライン達成（500%+）→ 改善余地あり';
  if (roas >= TARGETS.ROAS_MIN * 0.8)                                return '🟠 要注意（400%台）→ 素材・ターゲ見直し';
  return '🔴 要対応（500%未満）→ 即改善アクション';
}
function _cpaStatus(cpa) {
  if (!cpa) return '';
  if (cpa <= TARGETS.CPA * 0.8) return '🔵 優秀';
  if (cpa <= TARGETS.CPA)       return '🟢 目標内';
  if (cpa <= TARGETS.CPA * 1.3) return '🟡 要改善';
  return '🔴 要対応（CPA超過）';
}
function _giftingStatus(done) {
  if (done === null || done === undefined) return '';
  const now = new Date();
  const days = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const exp  = Math.round(TARGETS.GIFTING_PER_MONTH * now.getDate() / days);
  if (done >= TARGETS.GIFTING_PER_MONTH) return '🟢 月目標達成';
  if (done >= exp)                       return '🟢 ペース良好';
  if (done >= exp * 0.7)                 return '🟡 やや遅れ';
  return '🔴 要加速';
}

// ============================================================
// UIヘルパー
// ============================================================
function _writeSectionHeader(sh, row, title, bg, color) {
  sh.getRange(row, 1, 1, 6).merge()
    .setValue(title)
    .setBackground(bg    || '#16213e')
    .setFontColor(color  || '#a0c4ff')
    .setFontWeight('bold')
    .setFontSize(9);
  return row + 1;
}

function _writeDataRows(sh, row, rows, evenBg, oddBg) {
  if (!rows || rows.length === 0) return row;
  const numCols = rows[0].length;
  rows.forEach((r, i) => {
    sh.getRange(row + i, 1, 1, numCols)
      .setValues([r])
      .setBackground(i % 2 === 0 ? (evenBg || '#1a1a2e') : (oddBg || '#0d0d0d'))
      .setFontColor('#e0e0e0')
      .setFontSize(9);
    if (numCols >= 5) sh.getRange(row + i, 5).setFontColor('#ffd166');
  });
  return row + rows.length;
}

// ============================================================
// フォーマットヘルパー
// ============================================================
function _fmt(v)  { return v == null || isNaN(v) || v === 0 ? '取得中...' : '¥' + Math.round(v).toLocaleString(); }
function _fmt0(v) { return v == null || isNaN(v) ? '取得中...' : String(Math.round(v)); }
function _pct(v)  { if (v == null || isNaN(v)) return '取得中...'; const x = v > 1 ? v : v * 100; return x.toFixed(2) + '%'; }
function _num(v, d) { return v == null || isNaN(v) ? '取得中...' : v.toFixed(d || 0); }
function _n(v)    { const x = parseFloat(String(v || '').replace(/[,¥%]/g, '')); return isNaN(x) ? 0 : x; }
function _toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v).substring(0, 10);
}
function _fc(headers, cands) {
  for (const c of cands) {
    const re = new RegExp(c, 'i');
    const i  = headers.findIndex(h => re.test(h));
    if (i >= 0) return i;
  }
  return 0;
}

// ============================================================
// 月ラベル・タブ管理
// ============================================================
function _monthLabel(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年M月');
}

function _getOrCreateMonthSheet(ss, date) {
  const name = _monthLabel(date);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    _initMonthSheet(sh, date);
    Logger.log('✅ 新規月タブ作成: ' + name);
  }
  return sh;
}

function _getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ============================================================
// 先週（月〜日）の範囲を返す
// ============================================================
function _getWeekRange(date) {
  const d   = new Date(date);
  const dow = d.getDay();
  const lastMon = new Date(d);
  lastMon.setDate(d.getDate() - (dow === 0 ? 7 : dow) - 6);
  const lastSun = new Date(lastMon);
  lastSun.setDate(lastMon.getDate() + 6);
  const fmt = dt => Utilities.formatDate(dt, 'Asia/Tokyo', 'MM/dd');
  return {
    from:  Utilities.formatDate(lastMon, 'Asia/Tokyo', 'yyyy-MM-dd'),
    to:    Utilities.formatDate(lastSun, 'Asia/Tokyo', 'yyyy-MM-dd'),
    label: `${fmt(lastMon)}（月）〜 ${fmt(lastSun)}（日）`,
  };
}

// ============================================================
// トリガー設定
// ============================================================
function setupTriggers() {
  const fns = ['createWeeklyMeeting', 'createFridayMeeting', 'createMonthSheet'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (fns.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  // 月曜 8:00 → 火曜ミーティング枠
  ScriptApp.newTrigger('createWeeklyMeeting').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  // 木曜 8:00 → 金曜進捗枠
  ScriptApp.newTrigger('createFridayMeeting').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(8).create();
  // 毎月1日 8:00 → 翌月タブ先行作成
  ScriptApp.newTrigger('createMonthSheet').timeBased().onMonthDay(1).atHour(8).create();

  Logger.log('✅ トリガー設定完了（月曜/木曜/毎月1日）');
}
