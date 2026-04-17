// ============================================================
// 週次ミーティング 議事録スクリプト v4.0（週次+月次PDCA対応）
//
// タブ構成:
//   2026年4月, 2026年5月 ... → 月ごとに自動作成（火曜+金曜を縦積み）
//   PDCA_2026年4月 ...       → 月次PDCAサイクル管理シート
//   施策ログ                 → 全期間蓄積型
//   月次サマリー             → 月ごとのKPIまとめ
//
// トリガー:
//   月曜 8:00 → createWeeklyMeeting()  火曜ミーティング枠を追記
//   木曜 8:00 → createFridayMeeting()  金曜進捗枠を追記
//   毎月1日 8:00 → createMonthSheet()  翌月タブ + PDCAタブを先行作成
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
const SHEET_施策LOG   = '施策ログ';
const SHEET_MONTHLY  = '月次サマリー';
const SHEET_PDCA_PFX = 'PDCA_'; // + 月ラベル → "PDCA_2026年4月"

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
  setup月次PDCASheet(ss, new Date());
  setup施策LogSheet(ss);
  setupMonthlySheet(ss);

  // 今月の第1週エントリを挿入
  const week = _getWeekRange(new Date());
  _appendWeeklyEntry(monthSh, week);

  setupTriggers();
  Logger.log('✅ セットアップ完了: ' + _monthLabel(new Date()) + 'タブ + PDCAタブ + 施策ログ + 月次サマリー');
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
  _getOrCreateMonthSheet(ss, next);
  setup月次PDCASheet(ss, next);
  Logger.log('✅ 翌月タブ作成: ' + _monthLabel(next) + ' + PDCAタブ');
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

  // ─── 週次PDCAサイクル ───
  row = _writeSectionHeader(sh, row, '🔄 今週のPDCAサイクル', '#0a2a1a', '#86efac');
  const pdcaWeekDefs = [
    { stage: '📋 P（計画）', bg: '#0d2b0d', hint: '今週やること・狙い・仮説' },
    { stage: '⚡ D（実行）', bg: '#2b2000', hint: 'やったこと・実施した施策' },
    { stage: '📊 C（評価）', bg: '#0d0d2b', hint: '数値で確認したこと（ROAS/売上/ギフティング）' },
    { stage: '🔧 A（改善）', bg: '#2b0d0d', hint: '来週変えること・次に試すこと' },
  ];
  pdcaWeekDefs.forEach(def => {
    // ステージラベル（B列）
    sh.getRange(row, 2).setValue(def.stage)
      .setBackground(def.bg).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
    // 入力エリア（C〜F列をマージ）
    sh.getRange(row, 3, 1, 4).merge()
      .setValue(def.hint)
      .setBackground(def.bg).setFontColor('#555555').setFontStyle('italic').setFontSize(9);
    sh.getRange(row, 1).setBackground(def.bg);
    row++;
  });
  // PDCAタブへの誘導
  sh.getRange(row, 2, 1, 5).merge()
    .setValue('▶ 月次PDCAの全体計画・進捗は「' + SHEET_PDCA_PFX + _monthLabel(new Date()) + '」シートを参照')
    .setFontColor('#555555').setFontStyle('italic').setFontSize(8).setBackground('#0d0d0d');
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
// 月次PDCAシート（画像フォーマット準拠）
// ============================================================
function setup月次PDCASheet(ss, date) {
  const name = SHEET_PDCA_PFX + _monthLabel(date);
  const sh   = _getOrCreateSheet(ss, name);
  sh.clearContents();
  sh.clearFormats();

  // 列幅
  sh.setColumnWidth(1, 90);   // A: PDCAステージ
  sh.setColumnWidth(2, 120);  // B: エリア
  sh.setColumnWidth(3, 260);  // C: アクション内容
  sh.setColumnWidth(4, 80);   // D: 担当
  sh.setColumnWidth(5, 85);   // E: 期日
  sh.setColumnWidth(6, 70);   // F: 完了率
  sh.setColumnWidth(7, 110);  // G: ステータス
  sh.setColumnWidth(8, 220);  // H: 重要事項/備考

  const label = _monthLabel(date);

  // ─── 行1: メインタイトル ───
  sh.getRange(1, 1, 1, 8).merge()
    .setValue('PDCA 月次サイクル  ▌ ' + label)
    .setBackground('#0a0a1a').setFontColor('#e94560')
    .setFontWeight('bold').setFontSize(14);

  // ─── 行2〜5: 左=プロジェクト情報 / 右=PDCAビジュアル ───
  const infoData = [
    ['ブランド', 'gypsophilia（セレクト×オリジナル）', '', '', '改善', '計画'],
    ['サイクル', label + ' #1', '', '', '評価', '実行'],
    ['リーダー', 'yayoi', '', '', '', ''],
    ['更新日', '', '', '', '', ''],
  ];
  infoData.forEach((r, i) => {
    // 左側ラベル（A列）
    sh.getRange(2 + i, 1).setValue(r[0])
      .setBackground('#111122').setFontColor('#a0c4ff').setFontWeight('bold').setFontSize(9);
    // 左側値（B〜D列マージ）
    sh.getRange(2 + i, 2, 1, 3).merge().setValue(r[1])
      .setBackground('#0d0d1a').setFontColor('#e0e0e0').setFontSize(9);
    // 右側: PDCAビジュアル（E〜H列を2×2グリッド）
    if (i < 2) {
      const col1Label = r[4], col2Label = r[5];
      const colors = {
        '改善': { bg: '#e8f5e9', fg: '#1b5e20' },
        '計画': { bg: '#fce4ec', fg: '#880e4f' },
        '評価': { bg: '#e8eaf6', fg: '#1a237e' },
        '実行': { bg: '#fff9c4', fg: '#f57f17' },
      };
      [col1Label, col2Label].forEach((lbl, j) => {
        const c = colors[lbl] || { bg: '#333333', fg: '#ffffff' };
        sh.getRange(2 + i, 5 + j * 2, 1, 2).merge()
          .setValue(lbl).setBackground(c.bg).setFontColor(c.fg)
          .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
      });
    }
  });

  // ─── 行6: 進捗バー見出し ───
  sh.getRange(6, 1, 1, 8).merge()
    .setValue('─── 月間アクション計画・進捗 ───')
    .setBackground('#111122').setFontColor('#666666').setFontSize(8);

  // ─── 行7: テーブルヘッダー ───
  const headers = ['PDCAステージ', 'エリア', 'アクション内容 / 仮説', '担当', '期日', '完了率', 'ステータス', '重要事項/備考'];
  sh.getRange(7, 1, 1, 8).setValues([headers])
    .setBackground('#16213e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
  sh.setFrozenRows(7);

  // ─── 行8〜: アクション行（デフォルト） ───
  const PDCA_STAGES = {
    '計画': { bg: '#fce4ec', fg: '#880e4f' },
    '実行': { bg: '#fff9c4', fg: '#7f4f00' },
    '評価': { bg: '#e8eaf6', fg: '#1a237e' },
    '改善': { bg: '#e8f5e9', fg: '#1b5e20' },
  };

  const defaultActions = [
    // 計画
    ['計画', 'ギフティング', '月間ターゲット選定（フォロワー・ジャンル・投稿クオリティ基準）', 'yayoi', '月初', '0%', '未開始', 'DM文面も更新すること'],
    ['計画', 'ギフティング', 'DM送信スケジュール（目標35件）＋品番リスト確定',              'yayoi', '月初', '0%', '未開始', ''],
    ['計画', 'Meta広告',    '月間予算配分・キャンペーン構成見直し',                         'yayoi', '月初', '0%', '未開始', `目標ROAS ${TARGETS.ROAS_IDEAL*100}%`],
    ['計画', 'Meta広告',    '月間クリエイティブ計画（週2本確保のスケジュール）',             'yayoi', '月初', '0%', '未開始', ''],
    ['計画', 'MD',          '新商品仕入れ・在庫計画（無在庫/有在庫）',                      'MD',    '月初', '0%', '未開始', ''],
    // 実行
    ['実行', 'ギフティング', 'DM送信・返答フォロー・発送対応',  'yayoi', '月中', '0%', '未開始', '週次で進捗確認'],
    ['実行', 'ギフティング', '投稿確認・2次利用許可取得・広告転用', 'yayoi', '随時', '0%', '未開始', ''],
    ['実行', 'Meta広告',    'クリエイティブ追加入稿（週2本）',  'yayoi', '随時', '0%', '未開始', ''],
    ['実行', 'Meta広告',    '予算調整・低パフォーマンス素材停止', 'yayoi', '随時', '0%', '未開始', ''],
    ['実行', '自社リール',  'リール撮影・編集・投稿',           'shoi',  '随時', '0%', '未開始', ''],
    ['実行', 'MD',          '発注・納品・在庫登録',             'MD',    '随時', '0%', '未開始', ''],
    // 評価
    ['評価', '売上',        '月間純売上・注文数・客単価（新規/リピ）の振り返り', 'yayoi', '月末', '0%', '未開始', ''],
    ['評価', 'Meta広告',    '月間ROAS・CPA・CTR・クリエイティブ別パフォーマンス分析', 'yayoi', '月末', '0%', '未開始', `最低${TARGETS.ROAS_MIN*100}%達成確認`],
    ['評価', 'ギフティング', '実施数・投稿数・広告転用数・再生数の分析', 'yayoi', '月末', '0%', '未開始', ''],
    ['評価', 'コンテンツ',  '自社リール再生数・保存数・フォロワー増減', 'shoi',  '月末', '0%', '未開始', ''],
    // 改善
    ['改善', 'Meta広告',    '勝ちクリエイティブの横展開・負けパターン特定・停止', 'yayoi', '月末', '0%', '未開始', ''],
    ['改善', 'ギフティング', '翌月のターゲット属性・ジャンル・選定基準の見直し',  'yayoi', '月末', '0%', '未開始', ''],
    ['改善', 'MD',          '売れ筋・死に筋の整理・翌月仕入れ方針',              'MD',    '月末', '0%', '未開始', ''],
    ['改善', '全体',        '翌月のPDCA計画（計画フェーズへ）',                   'yayoi', '月末', '0%', '未開始', ''],
  ];

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未開始', '進行中', '完了済み', '保留'], true).build();

  defaultActions.forEach((r, i) => {
    const rowNum = 8 + i;
    const style  = PDCA_STAGES[r[0]] || { bg: '#1a1a2e', fg: '#ffffff' };
    sh.getRange(rowNum, 1, 1, 8).setValues([r]);
    // ステージ列（A）をカラー
    sh.getRange(rowNum, 1)
      .setBackground(style.bg).setFontColor(style.fg).setFontWeight('bold').setFontSize(9);
    // 残列
    sh.getRange(rowNum, 2, 1, 7)
      .setBackground(i % 2 === 0 ? '#1a1a2e' : '#111122')
      .setFontColor('#e0e0e0').setFontSize(9);
    // ステータスドロップダウン
    sh.getRange(rowNum, 7).setDataValidation(statusRule);
    // 完了率列（F）を数値色
    sh.getRange(rowNum, 6).setFontColor('#ffd166');
  });

  // 空白入力行を5行追加
  for (let i = 0; i < 5; i++) {
    const rowNum = 8 + defaultActions.length + i;
    sh.getRange(rowNum, 1, 1, 8).setBackground('#0d0d0d').setFontColor('#e0e0e0').setFontSize(9);
    sh.getRange(rowNum, 7).setDataValidation(statusRule);
  }

  SpreadsheetApp.flush();
  Logger.log('✅ 月次PDCAシート作成: ' + name);
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
