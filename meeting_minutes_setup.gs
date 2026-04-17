// ============================================================
// 週次ミーティング 議事録スクリプト v5.0（Apple-style ライトテーマ）
//
// タブ構成:
//   2026年4月, 2026年5月 ... → 月ごとに自動作成（火曜+金曜を縦積み）
//   PDCA_2026年4月 ...       → 月次PDCAサイクル管理シート
//   施策ログ                 → 全期間蓄積型
//   月次サマリー             → 月ごとのKPIまとめ
//
// トリガー:
//   月曜 8:00 → createWeeklyMeeting()
//   木曜 8:00 → createFridayMeeting()
//   毎月1日 8:00 → createMonthSheet()
// ============================================================

// ─── カラーテーマ（ここだけ変えれば全体に反映） ───
const T = {
  bg:        '#ffffff',   // ページ背景
  bgAlt:     '#f5f5f7',   // 交互行（Apple light gray）
  bgSection: '#f5f5f7',   // セクションヘッダー背景
  bgSubHdr:  '#fafafa',   // サブヘッダー（列名行）

  text:      '#1d1d1f',   // メインテキスト（Apple near-black）
  textSec:   '#6e6e73',   // サブテキスト
  textHint:  '#aeaeb2',   // ヒント・プレースホルダー

  accent:    '#0071e3',   // Apple Blue
  accentBg:  '#e8f4fe',   // 薄ブルー背景

  titleBg:   '#1d1d1f',   // ミーティングタイトルバー（near-black）
  titleFg:   '#ffffff',
  fridayFg:  '#f5f5f7',   // 金曜タイトル文字

  target:    '#0071e3',   // 目標値カラー
  targetBg:  '#e8f4fe',

  sep:       '#e5e5e5',   // 区切り線

  // PDCA ステージ（パステル）
  pdcaP:  { bg: '#fce4ec', fg: '#c2185b', label: '計画' },
  pdcaD:  { bg: '#fff8e1', fg: '#f57f17', label: '実行' },
  pdcaC:  { bg: '#e8eaf6', fg: '#3949ab', label: '評価' },
  pdcaA:  { bg: '#e8f5e9', fg: '#2e7d32', label: '改善' },

  // 週次PDCA ステージ行
  weekP:  { bg: '#fce4ec', fg: '#c2185b' },
  weekD:  { bg: '#fff8e1', fg: '#f57f17' },
  weekC:  { bg: '#e8eaf6', fg: '#3949ab' },
  weekA:  { bg: '#e8f5e9', fg: '#2e7d32' },
};

// ─── 他シートのスプレッドシートID ───
const SPREADSHEET_IDS = {
  META_ADS:   '1kd6HD4NOECKNl4knsklsrcUDWAOs-ve3xhULEtKrEt8',
  INFLUENCER: '1prt8F8EAoHlGlkIv6uLkTGXZwIUD13-_QDtIdvEOMIE',
  MARKETING:  '1fQg_01ViGrZEoT38LifM_zh-U4vah_8ORE4xTodz708',
};

// ─── 目標値 ───
const TARGETS = {
  ROAS_IDEAL:        7.0,
  ROAS_MIN:          5.0,
  ROAS:              5.0,
  CPA:               2000,
  GIFTING_PER_MONTH: 10,
  DM_PER_MONTH:      35,
  CREATIVE_PER_WEEK: 2,
};

// ─── 固定シート名 ───
const SHEET_施策LOG   = '施策ログ';
const SHEET_MONTHLY  = '月次サマリー';
const SHEET_PDCA_PFX = 'PDCA_';

// ============================================================
// 初期セットアップ（一度だけ手動実行）
// ============================================================
function setupAllSheets() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const monthSh = _getOrCreateMonthSheet(ss, new Date());

  _getOrCreateSheet(ss, SHEET_施策LOG);
  _getOrCreateSheet(ss, SHEET_MONTHLY);

  _initMonthSheet(monthSh, new Date());
  setup月次PDCASheet(ss, new Date());
  setup施策LogSheet(ss);
  setupMonthlySheet(ss);

  const week = _getWeekRange(new Date());
  _appendWeeklyEntry(monthSh, week);

  setupTriggers();
  Logger.log('✅ セットアップ完了: ' + _monthLabel(new Date()));
}

// ============================================================
// 毎週月曜 8:00
// ============================================================
function createWeeklyMeeting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = _getOrCreateMonthSheet(ss, new Date());
  _appendWeeklyEntry(sh, _getWeekRange(new Date()));
  Logger.log('✅ 火曜ミーティング枠追加');
}

// ============================================================
// 毎週木曜 8:00
// ============================================================
function createFridayMeeting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = _getOrCreateMonthSheet(ss, new Date());
  _appendFridayEntry(sh, new Date());
  Logger.log('✅ 金曜進捗枠追加');
}

// ============================================================
// 毎月1日 8:00: 翌月タブ先行作成
// ============================================================
function createMonthSheet() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const next = new Date();
  next.setMonth(next.getMonth() + 1);
  _getOrCreateMonthSheet(ss, next);
  setup月次PDCASheet(ss, next);
  Logger.log('✅ 翌月タブ作成: ' + _monthLabel(next));
}

// ============================================================
// 月別タブ 初期ヘッダー
// ============================================================
function _initMonthSheet(sh, date) {
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 24);
  sh.setColumnWidth(2, 190);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 165);
  sh.setColumnWidth(6, 290);

  // タイトル
  sh.getRange(1, 1, 1, 6).merge()
    .setValue(_monthLabel(date) + '  議事録')
    .setBackground(T.titleBg).setFontColor(T.titleFg)
    .setFontWeight('bold').setFontSize(14);

  // サブタイトル
  sh.getRange(2, 1, 1, 6).merge()
    .setValue('火曜ミーティング（毎週月曜 8:00 自動追記）　　金曜進捗（毎週木曜 8:00 自動追記）')
    .setBackground(T.bgAlt).setFontColor(T.textSec).setFontSize(8);

  sh.setFrozenRows(2);
  SpreadsheetApp.flush();
}

// ============================================================
// 火曜ミーティング エントリ
// ============================================================
function _appendWeeklyEntry(sh, week) {
  let row = Math.max(sh.getLastRow() + 2, 3);

  const ads  = _fetchMetaAdsWeekly(week.from, week.to);
  const shop = _fetchShopifyWeekly(week.from, week.to);
  const inf  = _fetchGiftingMonthly();

  // ─── タイトルバー ───
  sh.getRange(row, 1, 1, 6).merge()
    .setValue('火曜ミーティング   ' + week.label)
    .setBackground(T.titleBg).setFontColor(T.titleFg)
    .setFontWeight('bold').setFontSize(11);
  row++;

  // ─── 週次PDCAサイクル ───
  row = _writeSectionHeader(sh, row, '🔄  今週のPDCAサイクル');
  const pdcaWeek = [
    { label: 'P  計画', style: T.weekP, hint: '今週やること・狙い・仮説' },
    { label: 'D  実行', style: T.weekD, hint: 'やったこと・実施した施策' },
    { label: 'C  評価', style: T.weekC, hint: '数値で確認したこと（ROAS / 売上 / ギフティング）' },
    { label: 'A  改善', style: T.weekA, hint: '来週変えること・次に試すこと' },
  ];
  pdcaWeek.forEach(def => {
    sh.getRange(row, 1).setBackground(def.style.bg);
    sh.getRange(row, 2)
      .setValue(def.label)
      .setBackground(def.style.bg).setFontColor(def.style.fg)
      .setFontWeight('bold').setFontSize(9);
    sh.getRange(row, 3, 1, 4).merge()
      .setValue(def.hint)
      .setBackground(def.style.bg).setFontColor(T.textHint)
      .setFontStyle('italic').setFontSize(9);
    row++;
  });
  sh.getRange(row, 2, 1, 5).merge()
    .setValue('月次PDCAの全体計画・進捗 → 「' + SHEET_PDCA_PFX + _monthLabel(new Date()) + '」シート')
    .setBackground(T.bg).setFontColor(T.textHint).setFontStyle('italic').setFontSize(8);
  row++;

  // ─── ① 売上サマリー ───
  row = _writeSectionHeader(sh, row, '①  売上サマリー（先週）');
  row = _writeSubHeader(sh, row, ['', '指標', '実績', '前週比', '目標', 'メモ']);
  row = _writeDataRows(sh, row, [
    ['', '週間純売上',      _fmt(shop.netSales),  shop.salesWoW,  '—', ''],
    ['', '注文数',          _fmt0(shop.orders),   shop.ordersWoW, '—', ''],
    ['', '客単価・新規',    _fmt(shop.aovNew),    '—',            '—', ''],
    ['', '客単価・リピ',    _fmt(shop.aovReturn), '—',            '—', ''],
    ['', 'CV率',            _pct(shop.cvr),       '—',            '—', ''],
    ['', '新規 / リピ比',   shop.newVsReturn,     '—',            '—', ''],
  ]);

  // ─── ② Meta広告 ───
  row = _writeSectionHeader(sh, row, '②  Meta広告（先週）');
  row = _writeSubHeader(sh, row, ['', '指標', '実績', '前週比', '目標', 'ステータス']);
  row = _writeDataRows(sh, row, [
    ['', '消化金額',             _fmt(ads.spend),    ads.spendWoW, '—',                                               ''],
    ['', 'ROAS',                 _num(ads.roas, 2),  ads.roasWoW,  `理想 ${TARGETS.ROAS_IDEAL*100}% / 最低 ${TARGETS.ROAS_MIN*100}%`, _roasStatus(ads.roas)],
    ['', 'CPA',                  _fmt(ads.cpa),      ads.cpaWoW,   `¥${TARGETS.CPA}以下`,                             _cpaStatus(ads.cpa)],
    ['', 'CTR',                  _pct(ads.ctr),      '—',          '—',                                               ''],
    ['', 'CPM',                  _fmt(ads.cpm),      '—',          '—',                                               ''],
    ['', 'CV数',                 _fmt0(ads.cv),      '—',          '—',                                               ''],
    ['', 'CV売上',               _fmt(ads.cvSales),  '—',          '—',                                               ''],
    ['', '追加クリエイティブ数',  '（手入力）',       '—',          `週${TARGETS.CREATIVE_PER_WEEK}本+`,               ''],
  ]);

  // ─── ③ ギフティング ───
  row = _writeSectionHeader(sh, row, '③  ギフティング（今月進捗）');
  row = _writeSubHeader(sh, row, ['', '指標', '実績', '—', '目標', 'ステータス']);
  row = _writeDataRows(sh, row, [
    ['', 'DM送信数',              _fmt0(inf.dmCount),     '—', `${TARGETS.DM_PER_MONTH}件`,      ''],
    ['', '返答数',                _fmt0(inf.replyCount),  '—', '—',                              ''],
    ['', '返答率',                _pct(inf.replyRate),    '—', '—',                              ''],
    ['', 'ギフティング実施数',    _fmt0(inf.giftingDone), '—', `${TARGETS.GIFTING_PER_MONTH}件`,  _giftingStatus(inf.giftingDone)],
    ['', '投稿済み数',            _fmt0(inf.posted),      '—', '—',                              ''],
    ['', '新規素材（広告転用可）', '（手入力）',          '—', '—',                              ''],
  ]);

  // ─── ④ 今週の投稿予定 ───
  row = _writeSectionHeader(sh, row, '④  今週の投稿予定');
  row = _writeSubHeader(sh, row, ['', '@インスタID', '投稿予定日', 'ジャンル', '品番', '備考']);
  row = _writeDataRows(sh, row, [
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ]);

  // ─── ⑤ 課題・懸念事項 ───
  row = _writeSectionHeader(sh, row, '⑤  課題・懸念事項');
  row = _writeSubHeader(sh, row, ['', '課題内容', 'オーナー', '緊急度', '期限', 'ステータス']);
  row = _writeDataRows(sh, row, [
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ]);

  // ─── ⑥ アクションアイテム ───
  row = _writeSectionHeader(sh, row, '⑥  アクションアイテム（今週）');
  row = _writeSubHeader(sh, row, ['', 'アクション内容', '担当', 'Due', 'ステータス', '']);
  row = _writeDataRows(sh, row, [
    ['', '', 'yayoi', '', '[ ] 未', ''],
    ['', '', 'shoi',  '', '[ ] 未', ''],
    ['', '', 'MD',    '', '[ ] 未', ''],
  ]);

  // ─── ⑦ 施策実験ログ ───
  row = _writeSectionHeader(sh, row, '⑦  施策実験ログ（今週試したこと）');
  row = _writeSubHeader(sh, row, ['', '施策種別', '施策内容・仮説', '計測指標', '結果', '判定']);
  const expRows = [
    ['', 'ギフティング', '', '再生数 / 保存数', '', ''],
    ['', 'ギフティング', '', '再生数 / 保存数', '', ''],
    ['', 'Meta広告',    '', 'ROAS / CTR',       '', ''],
    ['', '自社リール',  '', '再生数 / 保存数',  '', ''],
  ];
  const typeRule  = SpreadsheetApp.newDataValidation().requireValueInList(['ギフティング', 'Meta広告', '自社リール', 'MD/商品', 'その他'], true).build();
  const judgeRule = SpreadsheetApp.newDataValidation().requireValueInList(['✅ 継続', '🔄 改善して継続', '⏸ 一時停止', '❌ 中止', '📊 計測中'], true).build();
  expRows.forEach((r, i) => {
    sh.getRange(row + i, 1, 1, 6).setValues([r])
      .setBackground(i % 2 === 0 ? T.bg : T.bgAlt)
      .setFontColor(T.text).setFontSize(9);
    sh.getRange(row + i, 2).setDataValidation(typeRule);
    sh.getRange(row + i, 6).setDataValidation(judgeRule);
  });
  row += expRows.length;
  sh.getRange(row, 2, 1, 5).merge()
    .setValue('詳細な施策履歴は「施策ログ」シートへ転記')
    .setBackground(T.bg).setFontColor(T.textHint).setFontStyle('italic').setFontSize(8);
  row++;

  // 区切り線
  sh.getRange(row, 1, 1, 6)
    .setBackground(T.sep).setValues([['', '', '', '', '', '']]);
  row += 2;
}

// ============================================================
// 金曜進捗 エントリ
// ============================================================
function _appendFridayEntry(sh, date) {
  let row = Math.max(sh.getLastRow() + 2, 3);
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd (E)');

  sh.getRange(row, 1, 1, 6).merge()
    .setValue('金曜進捗ミーティング   ' + dateStr)
    .setBackground(T.titleBg).setFontColor(T.fridayFg)
    .setFontWeight('bold').setFontSize(11);
  row++;

  row = _writeSectionHeader(sh, row, '①  アクション進捗確認（火曜から転記）');
  row = _writeSubHeader(sh, row, ['', 'アクション内容', '担当', 'ステータス', 'メモ', '']);
  row = _writeDataRows(sh, row, [
    ['', '', 'yayoi', '', '', ''],
    ['', '', 'shoi',  '', '', ''],
    ['', '', 'MD',    '', '', ''],
  ]);

  row = _writeSectionHeader(sh, row, '②  ブロッカー・課題');
  row = _writeDataRows(sh, row, [
    ['', '', '', '', '', ''],
    ['', '', '', '', '', ''],
  ]);

  row = _writeSectionHeader(sh, row, '③  翌週準備チェック');
  row = _writeDataRows(sh, row, [
    ['', '来週ギフティングDM 準備済み？',    '', '[ ] Yes  [ ] No', '', ''],
    ['', '広告クリエイティブ 2本以上確保？',  '', '[ ] Yes  [ ] No', '', ''],
    ['', '在庫確認（オリジナル）完了？',      '', '[ ] Yes  [ ] No', '', ''],
    ['', '翌週ミーティングアジェンダ作成？',  '', '[ ] Yes  [ ] No', '', ''],
  ]);

  sh.getRange(row, 1, 1, 6).setBackground(T.sep).setValues([['', '', '', '', '', '']]);
}

// ============================================================
// 月次PDCAシート
// ============================================================
function setup月次PDCASheet(ss, date) {
  const name = SHEET_PDCA_PFX + _monthLabel(date);
  const sh   = _getOrCreateSheet(ss, name);
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 260);
  sh.setColumnWidth(4, 80);
  sh.setColumnWidth(5, 85);
  sh.setColumnWidth(6, 70);
  sh.setColumnWidth(7, 110);
  sh.setColumnWidth(8, 220);

  const label = _monthLabel(date);

  // ─── タイトル ───
  sh.getRange(1, 1, 1, 8).merge()
    .setValue('PDCA  月次サイクル   ' + label)
    .setBackground(T.titleBg).setFontColor(T.titleFg)
    .setFontWeight('bold').setFontSize(14);

  // ─── プロジェクト情報 + PDCAビジュアル（行2〜5） ───
  const infoRows = [
    { label: 'ブランド', value: 'gypsophilia（セレクト × オリジナル）' },
    { label: 'サイクル', value: label + '  #1' },
    { label: 'リーダー', value: 'yayoi' },
    { label: '更新日',   value: '' },
  ];
  const pdcaVisual = [
    [T.pdcaA, T.pdcaP],   // 行2: 改善 | 計画
    [T.pdcaC, T.pdcaD],   // 行3: 評価 | 実行
  ];

  infoRows.forEach((info, i) => {
    sh.getRange(2 + i, 1)
      .setValue(info.label)
      .setBackground(T.bgAlt).setFontColor(T.textSec)
      .setFontWeight('bold').setFontSize(9);
    sh.getRange(2 + i, 2, 1, 3).merge()
      .setValue(info.value)
      .setBackground(T.bg).setFontColor(T.text).setFontSize(9);

    if (i < 2) {
      pdcaVisual[i].forEach((stage, j) => {
        sh.getRange(2 + i, 6 + j, 1, 1)
          .setValue(stage.label)
          .setBackground(stage.bg).setFontColor(stage.fg)
          .setFontWeight('bold').setFontSize(14)
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
      });
    }
    if (i >= 2) {
      sh.getRange(2 + i, 6, 1, 3).merge()
        .setBackground(T.bgAlt);
    }
  });

  // ─── 区切り ───
  sh.getRange(6, 1, 1, 8).merge()
    .setValue('月間アクション計画・進捗管理')
    .setBackground(T.bgAlt).setFontColor(T.textSec)
    .setFontWeight('bold').setFontSize(9);

  // ─── テーブルヘッダー ───
  sh.getRange(7, 1, 1, 8)
    .setValues([['PDCAステージ', 'エリア', 'アクション内容 / 仮説', '担当', '期日', '完了率', 'ステータス', '重要事項 / 備考']])
    .setBackground(T.titleBg).setFontColor(T.titleFg)
    .setFontWeight('bold').setFontSize(9);
  sh.setFrozenRows(7);

  // ─── アクション行 ───
  const STAGE = {
    '計画': T.pdcaP,
    '実行': T.pdcaD,
    '評価': T.pdcaC,
    '改善': T.pdcaA,
  };

  const actions = [
    ['計画', 'ギフティング', '月間ターゲット選定（フォロワー・ジャンル・動画クオリティ基準）', 'yayoi', '月初', '0%', '未開始', 'DM文面も更新すること'],
    ['計画', 'ギフティング', 'DM送信スケジュール（目標35件）＋品番リスト確定',               'yayoi', '月初', '0%', '未開始', ''],
    ['計画', 'Meta広告',    '月間予算配分・キャンペーン構成見直し',                          'yayoi', '月初', '0%', '未開始', `理想ROAS ${TARGETS.ROAS_IDEAL*100}%`],
    ['計画', 'Meta広告',    '月間クリエイティブ計画（週2本確保スケジュール）',                'yayoi', '月初', '0%', '未開始', ''],
    ['計画', 'MD',          '新商品仕入れ・在庫計画（無在庫 / 有在庫）',                     'MD',    '月初', '0%', '未開始', ''],
    ['実行', 'ギフティング', 'DM送信・返答フォロー・発送対応',                               'yayoi', '月中', '0%', '未開始', '週次で進捗確認'],
    ['実行', 'ギフティング', '投稿確認・2次利用許可取得・広告転用',                          'yayoi', '随時', '0%', '未開始', ''],
    ['実行', 'Meta広告',    'クリエイティブ追加入稿（週2本）',                               'yayoi', '随時', '0%', '未開始', ''],
    ['実行', 'Meta広告',    '予算調整・低パフォーマンス素材停止',                            'yayoi', '随時', '0%', '未開始', ''],
    ['実行', '自社リール',  'リール撮影・編集・投稿',                                        'shoi',  '随時', '0%', '未開始', ''],
    ['実行', 'MD',          '発注・納品・在庫登録',                                          'MD',    '随時', '0%', '未開始', ''],
    ['評価', '売上',        '月間純売上・注文数・客単価（新規 / リピ）の振り返り',            'yayoi', '月末', '0%', '未開始', ''],
    ['評価', 'Meta広告',    '月間ROAS・CPA・CTR・クリエイティブ別分析',                      'yayoi', '月末', '0%', '未開始', `最低${TARGETS.ROAS_MIN*100}%達成確認`],
    ['評価', 'ギフティング', '実施数・投稿数・広告転用数・再生数の分析',                     'yayoi', '月末', '0%', '未開始', ''],
    ['評価', 'コンテンツ',  '自社リール 再生数・保存数・フォロワー増減',                     'shoi',  '月末', '0%', '未開始', ''],
    ['改善', 'Meta広告',    '勝ちクリエイティブの横展開・負けパターン特定・停止',            'yayoi', '月末', '0%', '未開始', ''],
    ['改善', 'ギフティング', '翌月ターゲット属性・ジャンル・選定基準の見直し',               'yayoi', '月末', '0%', '未開始', ''],
    ['改善', 'MD',          '売れ筋・死に筋の整理・翌月仕入れ方針',                         'MD',    '月末', '0%', '未開始', ''],
    ['改善', '全体',        '翌月のPDCA計画策定（計画フェーズへ）',                          'yayoi', '月末', '0%', '未開始', ''],
  ];

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未開始', '進行中', '完了済み', '保留'], true).build();

  actions.forEach((r, i) => {
    const rowNum = 8 + i;
    const stage  = STAGE[r[0]] || { bg: T.bgAlt, fg: T.text };
    sh.getRange(rowNum, 1, 1, 8).setValues([r]);
    sh.getRange(rowNum, 1)
      .setBackground(stage.bg).setFontColor(stage.fg)
      .setFontWeight('bold').setFontSize(9);
    sh.getRange(rowNum, 2, 1, 7)
      .setBackground(i % 2 === 0 ? T.bg : T.bgAlt)
      .setFontColor(T.text).setFontSize(9);
    sh.getRange(rowNum, 6).setFontColor(T.accent);
    sh.getRange(rowNum, 7).setDataValidation(statusRule);
  });

  // 空白入力行
  for (let i = 0; i < 5; i++) {
    const rowNum = 8 + actions.length + i;
    sh.getRange(rowNum, 1, 1, 8)
      .setBackground(i % 2 === 0 ? T.bg : T.bgAlt)
      .setFontColor(T.text).setFontSize(9);
    sh.getRange(rowNum, 7).setDataValidation(statusRule);
  }

  SpreadsheetApp.flush();
  Logger.log('✅ 月次PDCAシート: ' + name);
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

  // ヘッダー
  sh.getRange(1, 1, 1, 8)
    .setValues([['記録日', '施策種別', '施策内容・仮説', '計測指標・結果', '数値', '判定', '振り返り・学び', '次のアクション']])
    .setBackground(T.titleBg).setFontColor(T.titleFg)
    .setFontWeight('bold').setFontSize(9);
  sh.setFrozenRows(1);

  // 使い方ガイド行
  sh.getRange(2, 1, 1, 8)
    .setValues([['使い方', '毎週⑦で記録 → ここに転記 → 月次レビューで傾向分析', '', '', '', '', '', '']])
    .setBackground(T.bgAlt).setFontColor(T.textSec).setFontSize(8);

  // サンプルデータ
  const samples = [
    ['2026/04/15', 'ギフティング',
     '【仮説】ファッション特化5〜10万フォロワーは購買転換率が高い',
     '対象5名 平均再生数 / CV数 / ROAS比較', 'ROAS 6.2', '🔄 改善して継続',
     '再生10万超の1本がCV全体60%を牽引。フォロワー数より動画クオリティが重要',
     '次月もファッション特化・クオリティ重視で選定継続'],
    ['2026/04/08', 'Meta広告',
     '【仮説】インフルエンサーUGC vs 自社制作でCTR比較',
     'CTR / CPA / ROAS（7日間テスト）', 'UGC: CTR 4.2% / 自社: 2.1%', '✅ 継続',
     'UGCのCTRが2倍。自社制作は商品説明動画に絞る',
     'ギフティング動画を積極的に2次利用。自社制作は商品ハイライトのみ'],
  ];
  sh.getRange(3, 1, samples.length, 8).setValues(samples)
    .setBackground(T.bg).setFontColor(T.text).setFontSize(9);
  sh.getRange(4, 1, 1, 8).setBackground(T.bgAlt);

  // 入力行（5行目〜）ドロップダウン + 背景
  const typeRule  = SpreadsheetApp.newDataValidation().requireValueInList(['ギフティング', 'Meta広告', '自社リール', 'MD/商品', 'サイト/LP', 'その他'], true).build();
  const judgeRule = SpreadsheetApp.newDataValidation().requireValueInList(['✅ 継続', '🔄 改善して継続', '⏸ 一時停止', '❌ 中止', '📊 計測中'], true).build();
  for (let i = 0; i < 50; i++) {
    sh.getRange(5 + i, 1, 1, 8)
      .setBackground(i % 2 === 0 ? T.bg : T.bgAlt)
      .setFontColor(T.text).setFontSize(9);
  }
  sh.getRange(5, 2, 50, 1).setDataValidation(typeRule);
  sh.getRange(5, 6, 50, 1).setDataValidation(judgeRule);

  Logger.log('✅ 施策ログシート完了');
}

// ============================================================
// 月次サマリーシート
// ============================================================
function setupMonthlySheet(ss) {
  const sh = ss.getSheetByName(SHEET_MONTHLY);
  sh.clearContents();
  sh.clearFormats();

  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 80);

  sh.getRange(1, 1, 1, 6)
    .setValues([['指標', '今月実績', '先月実績', '前月比', '目標', '達成']])
    .setBackground(T.titleBg).setFontColor(T.titleFg)
    .setFontWeight('bold').setFontSize(9);
  sh.setFrozenRows(1);

  const kpis = [
    ['── 売上 ──'],
    ['月間純売上', '', '', '', '—'],
    ['注文数',     '', '', '', '—'],
    ['客単価・新規', '', '', '', '—'],
    ['新規 / リピ比', '', '', '', '—'],
    ['── Meta広告 ──'],
    ['月間消化金額', '', '', '', '—'],
    ['月間ROAS',    '', '', '', `理想 ${TARGETS.ROAS_IDEAL*100}% / 最低 ${TARGETS.ROAS_MIN*100}%`],
    ['月間CPA',     '', '', '', `¥${TARGETS.CPA}以下`],
    ['── ギフティング ──'],
    ['DM送信数',         '', '', '', `${TARGETS.DM_PER_MONTH}件`],
    ['ギフティング実施数', '', '', '', `${TARGETS.GIFTING_PER_MONTH}件`],
    ['投稿済み数',        '', '', '', '—'],
    ['広告転用数',        '', '', '', `月${TARGETS.CREATIVE_PER_WEEK*4}本+`],
    ['── 施策振り返り ──'],
    ['今月の勝ちパターン', '', '', '', ''],
    ['今月の改善点',      '', '', '', ''],
    ['来月に試すこと',    '', '', '', ''],
  ];

  kpis.forEach((r, i) => {
    const fullRow = r.length < 6 ? [...r, ...Array(6 - r.length).fill('')] : r;
    sh.getRange(i + 2, 1, 1, 6).setValues([fullRow]);
    if (String(r[0]).startsWith('──')) {
      sh.getRange(i + 2, 1, 1, 6).merge()
        .setBackground(T.bgSection).setFontColor(T.accent)
        .setFontWeight('bold').setFontSize(9);
    } else {
      sh.getRange(i + 2, 1, 1, 6)
        .setBackground(i % 2 === 0 ? T.bg : T.bgAlt)
        .setFontColor(T.text).setFontSize(9);
      sh.getRange(i + 2, 5).setFontColor(T.accent).setFontWeight('bold');
    }
  });

  Logger.log('✅ 月次サマリーシート完了');
}

// ============================================================
// データ取得: Meta広告
// ============================================================
function _fetchMetaAdsWeekly(fromDate, toDate) {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.META_ADS);
    const sh   = ss.getSheets()[0];
    const data = sh.getDataRange().getValues();
    let hRow = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).includes('日付') || String(data[i][1]).includes('消化')) { hRow = i; break; }
    }
    if (hRow < 0) return _emptyAds();
    const h = data[hRow].map(v => String(v).trim());
    const c = {
      date: _fc(h,['日付']), spend: _fc(h,['消化金額']), roas: _fc(h,['ROAS']),
      cpa: _fc(h,['CPA']), ctr: _fc(h,['CTR']), cpm: _fc(h,['CPM']),
      cv: _fc(h,['CV数']), cvSales: _fc(h,['CV売上']),
    };
    let spend=0, rS=0, ctrS=0, cpmS=0, cv=0, cvSales=0, cnt=0;
    for (let i = hRow+1; i < data.length; i++) {
      const d = _toDateStr(data[i][c.date]);
      if (d >= fromDate && d <= toDate) {
        spend += _n(data[i][c.spend]); cv += _n(data[i][c.cv]); cvSales += _n(data[i][c.cvSales]);
        rS += _n(data[i][c.roas]); ctrS += _n(data[i][c.ctr]); cpmS += _n(data[i][c.cpm]); cnt++;
      }
    }
    if (cnt===0) return _emptyAds();
    return { spend, roas: rS/cnt, cpa: cv>0?spend/cv:0, ctr: ctrS/cnt, cpm: cpmS/cnt, cv, cvSales, spendWoW:'—', roasWoW:'—', cpaWoW:'—' };
  } catch(e) { Logger.log('Meta広告エラー: '+e.message); return _emptyAds(); }
}
function _emptyAds() {
  return { spend:null,roas:null,cpa:null,ctr:null,cpm:null,cv:null,cvSales:null,spendWoW:'—',roasWoW:'—',cpaWoW:'—' };
}

// ============================================================
// データ取得: Shopify
// ============================================================
function _fetchShopifyWeekly(fromDate, toDate) {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.MARKETING);
    const sh   = ss.getSheets()[0];
    const data = sh.getDataRange().getValues();
    let hRow = -1;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      if (String(data[i][0]).includes('日付') || String(data[i][1]).includes('売上')) { hRow = i; break; }
    }
    if (hRow < 0) return _emptyShop();
    const h = data[hRow].map(v => String(v).trim());
    const c = {
      date: _fc(h,['日付']), netSales: _fc(h,['純売上','売上高']), orders: _fc(h,['注文数']),
      aovNew: _fc(h,['新規.*単価']), aovReturn: _fc(h,['リピ.*単価','既存.*単価']),
      cvr: _fc(h,['CV率']), newOrders: _fc(h,['新規.*注文']), retOrders: _fc(h,['リピ.*注文','既存.*注文']),
    };
    let ns=0,ord=0,aN=0,aR=0,cvrS=0,nO=0,rO=0,cnt=0;
    for (let i = hRow+1; i < data.length; i++) {
      const d = _toDateStr(data[i][c.date]);
      if (d >= fromDate && d <= toDate) {
        ns+=_n(data[i][c.netSales]); ord+=_n(data[i][c.orders]);
        aN+=_n(data[i][c.aovNew]); aR+=_n(data[i][c.aovReturn]); cvrS+=_n(data[i][c.cvr]);
        nO+=_n(data[i][c.newOrders]); rO+=_n(data[i][c.retOrders]); cnt++;
      }
    }
    if (cnt===0) return _emptyShop();
    const tot = nO+rO;
    return {
      netSales:ns, orders:ord, aovNew:cnt>0?aN/cnt:null, aovReturn:cnt>0?aR/cnt:null, cvr:cnt>0?cvrS/cnt:null,
      newVsReturn: tot>0?`新規 ${Math.round(nO/tot*100)}% / リピ ${Math.round(rO/tot*100)}%`:'—',
      salesWoW:'—', ordersWoW:'—',
    };
  } catch(e) { Logger.log('Shopifyエラー: '+e.message); return _emptyShop(); }
}
function _emptyShop() {
  return { netSales:null,orders:null,aovNew:null,aovReturn:null,cvr:null,newVsReturn:'—',salesWoW:'—',ordersWoW:'—' };
}

// ============================================================
// データ取得: ギフティング
// ============================================================
function _fetchGiftingMonthly() {
  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_IDS.INFLUENCER);
    const sh   = ss.getSheetByName('マスター') || ss.getSheets()[0];
    const data = sh.getDataRange().getValues();
    const now=new Date(), thisY=now.getFullYear(), thisM=now.getMonth();
    let hRow=-1;
    for (let i=0;i<Math.min(5,data.length);i++) {
      if (String(data[i][0]).includes('No')||String(data[i][3]).includes('ID')) { hRow=i; break; }
    }
    if (hRow<0) return _emptyGifting();
    const h = data[hRow].map(v=>String(v).trim());
    const c = { dmDate:_fc(h,['DM日']), status:_fc(h,['ステータス']), gifted:_fc(h,['ギフティング','提供','発送']), posted:_fc(h,['投稿日','掲載日']) };
    let dm=0,reply=0,gifted=0,posted=0;
    for (let i=hRow+1;i<data.length;i++) {
      const r=data[i];
      const dmD=r[c.dmDate]?new Date(r[c.dmDate]):null;
      if (dmD&&!isNaN(dmD)&&dmD.getFullYear()===thisY&&dmD.getMonth()===thisM) {
        dm++; if (/[1234]|返答|承認|完了/.test(String(r[c.status]||''))) reply++;
      }
      const gD=r[c.gifted]?new Date(r[c.gifted]):null;
      if (gD&&!isNaN(gD)&&gD.getFullYear()===thisY&&gD.getMonth()===thisM) gifted++;
      const pD=r[c.posted]?new Date(r[c.posted]):null;
      if (pD&&!isNaN(pD)&&pD.getFullYear()===thisY&&pD.getMonth()===thisM) posted++;
    }
    return { dmCount:dm, replyCount:reply, replyRate:dm>0?reply/dm:0, giftingDone:gifted, posted };
  } catch(e) { Logger.log('ギフティングエラー: '+e.message); return _emptyGifting(); }
}
function _emptyGifting() {
  return { dmCount:null,replyCount:null,replyRate:null,giftingDone:null,posted:null };
}

// ============================================================
// ステータス判定
// ============================================================
function _roasStatus(roas) {
  if (!roas) return '';
  if (roas >= TARGETS.ROAS_IDEAL)                           return '理想達成（700%+）→ 予算拡大';
  if (roas >= (TARGETS.ROAS_IDEAL+TARGETS.ROAS_MIN)/2)      return '好調（600%+）→ 現状維持';
  if (roas >= TARGETS.ROAS_MIN)                             return '最低ライン達成（500%+）';
  if (roas >= TARGETS.ROAS_MIN*0.8)                         return '⚠️ 要注意（400%台）';
  return '🚨 要対応（500%未満）';
}
function _cpaStatus(cpa) {
  if (!cpa) return '';
  if (cpa <= TARGETS.CPA*0.8) return '優秀';
  if (cpa <= TARGETS.CPA)     return '目標内';
  if (cpa <= TARGETS.CPA*1.3) return '⚠️ 要改善';
  return '🚨 CPA超過';
}
function _giftingStatus(done) {
  if (done==null) return '';
  const now=new Date(), days=new Date(now.getFullYear(),now.getMonth()+1,0).getDate();
  const exp=Math.round(TARGETS.GIFTING_PER_MONTH*now.getDate()/days);
  if (done>=TARGETS.GIFTING_PER_MONTH) return '月目標達成';
  if (done>=exp)   return 'ペース良好';
  if (done>=exp*0.7) return '⚠️ やや遅れ';
  return '🚨 要加速';
}

// ============================================================
// UIヘルパー
// ============================================================
function _writeSectionHeader(sh, row, title) {
  sh.getRange(row, 1, 1, 6).merge()
    .setValue(title)
    .setBackground(T.bgSection).setFontColor(T.text)
    .setFontWeight('bold').setFontSize(9);
  return row + 1;
}

function _writeSubHeader(sh, row, cols) {
  sh.getRange(row, 1, 1, cols.length).setValues([cols])
    .setBackground(T.bgAlt).setFontColor(T.textSec)
    .setFontSize(8).setFontStyle('normal');
  return row + 1;
}

function _writeDataRows(sh, row, rows) {
  if (!rows||rows.length===0) return row;
  const nc = rows[0].length;
  rows.forEach((r, i) => {
    sh.getRange(row+i, 1, 1, nc)
      .setValues([r])
      .setBackground(i%2===0 ? T.bg : T.bgAlt)
      .setFontColor(T.text).setFontSize(9);
    if (nc>=5) sh.getRange(row+i, 5).setFontColor(T.accent).setFontWeight('bold');
  });
  return row + rows.length;
}

// ============================================================
// フォーマットヘルパー
// ============================================================
function _fmt(v)    { return v==null||isNaN(v)||v===0?'—':'¥'+Math.round(v).toLocaleString(); }
function _fmt0(v)   { return v==null||isNaN(v)?'—':String(Math.round(v)); }
function _pct(v)    { if(v==null||isNaN(v)) return '—'; const x=v>1?v:v*100; return x.toFixed(2)+'%'; }
function _num(v,d)  { return v==null||isNaN(v)?'—':v.toFixed(d||0); }
function _n(v)      { const x=parseFloat(String(v||'').replace(/[,¥%]/g,'')); return isNaN(x)?0:x; }
function _toDateStr(v) {
  if(!v) return '';
  if(v instanceof Date) return Utilities.formatDate(v,'Asia/Tokyo','yyyy-MM-dd');
  return String(v).substring(0,10);
}
function _fc(headers, cands) {
  for (const c of cands) {
    const i = headers.findIndex(h => new RegExp(c,'i').test(h));
    if (i>=0) return i;
  }
  return 0;
}

// ============================================================
// 月ラベル・タブ管理
// ============================================================
function _monthLabel(date) {
  return Utilities.formatDate(date,'Asia/Tokyo','yyyy年M月');
}
function _getOrCreateMonthSheet(ss, date) {
  const name = _monthLabel(date);
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); _initMonthSheet(sh, date); Logger.log('✅ 新規月タブ: '+name); }
  return sh;
}
function _getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ============================================================
// 先週（月〜日）の範囲
// ============================================================
function _getWeekRange(date) {
  const d=new Date(date), dow=d.getDay();
  const mon=new Date(d); mon.setDate(d.getDate()-(dow===0?7:dow)-6);
  const sun=new Date(mon); sun.setDate(mon.getDate()+6);
  const fmt=dt=>Utilities.formatDate(dt,'Asia/Tokyo','MM/dd');
  return {
    from: Utilities.formatDate(mon,'Asia/Tokyo','yyyy-MM-dd'),
    to:   Utilities.formatDate(sun,'Asia/Tokyo','yyyy-MM-dd'),
    label:`${fmt(mon)}（月）〜 ${fmt(sun)}（日）`,
  };
}

// ============================================================
// トリガー設定
// ============================================================
function setupTriggers() {
  const fns = ['createWeeklyMeeting','createFridayMeeting','createMonthSheet'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (fns.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('createWeeklyMeeting').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ScriptApp.newTrigger('createFridayMeeting').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(8).create();
  ScriptApp.newTrigger('createMonthSheet').timeBased().onMonthDay(1).atHour(8).create();
  Logger.log('✅ トリガー設定完了');
}
