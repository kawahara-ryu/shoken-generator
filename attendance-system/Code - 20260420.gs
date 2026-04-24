// =======================================================================
//  🛸 サイバー・コマンドセンター出欠管理システム v4.0
//  -----------------------------------------------------------------------
//  機能一覧:
//    A) ジャンプメニュー (アンチグラビティ・スクロール)
//    B) 3日連続欠席アラート
//    C) ホログラム・ダッシュボード
//    D) デイリー・オートメーション
//    E) 初期セットアップ
//    F) クラス別人数設定（設定シートで管理）
//    G) 転入生・転校生管理
// =======================================================================

// ─── 定数定義 ───────────────────────────────────────────────────────────

/** メインシート名 */
var SHEET_MAIN     = 'メイン';
/** ログシート名 */
var SHEET_LOG      = 'ログ';
/** 分析シート名 */
var SHEET_ANALYSIS = '分析';
/** 設定シート名 */
var SHEET_CONFIG   = '設定';
/** 運用フローシート名 */
var SHEET_FLOW     = '運用フロー';
/** 一覧シート名 */
var SHEET_OVERVIEW = '一覧';

/** クラス数 */
var NUM_CLASSES = 6;

/** デフォルトの1クラスあたりの生徒数（設定シートが未設定の場合） */
var DEFAULT_STUDENTS = 42;

/** ヘッダー行数（タイトル + ナビ + 日付） */
var HEADER_ROWS = 3;

/** 転校ステータス */
var STATUS_TRANSFERRED = '転校';

/** カラーパレット */
var COLOR = {
  BG_BLACK:      '#000000',
  NEON_BLUE:     '#00f2ff',
  CYBER_GREEN:   '#39ff14',
  MAGENTA:       '#ff00ff',
  ORANGE:        '#ff6600',
  DARK_NAVY:     '#1a1a2e',
  DARK_RED:      '#330000',
  DARK_YELLOW:   '#333300',
  DARK_CYAN:     '#003333',
  DARK_ORANGE:   '#331a00',
  DARK_MAGENTA:  '#1a001a',
  DARK_LIME:     '#1a3300',
  HEADER_BG:     '#0a0a1a',
  SEPARATOR_BG:  '#001a00',
  GRAY:          '#2a2a2a',
  GRAY_TEXT:     '#555555',
  WHITE:         '#ffffff'
};

/** データ列: A=No, B=氏名, C=🚨, D=ステータス, E=理由, F=備考, G=時刻 */
var NUM_COLS = 7;

/** ステータス省略記号マッピング */
var STATUS_SHORT = {
  '出席': '',
  '欠席': '欠',
  '遅刻': 'チ',
  '早退': 'ソ',
  '遅刻早退': 'チソ',
  '保健室': '保',
  '出停': '停',
  '転校': '転'
};

/** 全ステータスリスト（プルダウン用、転校含む） */
var ALL_STATUSES = ['出席', '欠席', '遅刻', '早退', '遅刻早退', '保健室', '出停', '転校'];

/** 集計対象ステータス（転校を除く） */
var ACTIVE_STATUSES = ['出席', '欠席', '遅刻', '早退', '遅刻早退', '保健室', '出停'];

/** 理由リスト（プルダウン用） */
var ALL_REASONS = [
  '体調不良',
  '家事都合',
  '通院',
  '入院',

  '公欠',
  '新型コロナウイルス',
  'インフルエンザ',
  'マイコプラズマ肺炎',
  '進学試験',
  '就職試験',
  '大雨による交通機関不通',
  '大雪による交通機関不通',
  '台風による交通機関不通',
  '忌引（○○の死亡）',
  '学級閉鎖',
  'その他'
];

// ─── ユーティリティ ──────────────────────────────────────────────────────

/**
 * 設定シートから各クラスの生徒数を読み取る
 * @param {number} classNum - クラス番号 (1〜6)
 * @return {number} そのクラスの生徒数
 */
function getStudentsPerClass(classNum) {
  if (!classNum || classNum < 1 || classNum > NUM_CLASSES) return DEFAULT_STUDENTS;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) return DEFAULT_STUDENTS;

  // 設定シート: 行1=タイトル, 行2=ヘッダー, 行3〜8=クラス1〜6のデータ
  var targetRow = classNum + 2;
  if (targetRow < 1 || targetRow > configSheet.getMaxRows()) return DEFAULT_STUDENTS;

  try {
    var value = configSheet.getRange(targetRow, 2).getValue();
    if (value && !isNaN(value) && value > 0) {
      return Number(value);
    }
  } catch (e) {
    // シートの行が不足している場合はデフォルト値を返す
  }
  return DEFAULT_STUDENTS;
}

/**
 * 全クラスの生徒数を配列で返す
 * @return {number[]} 各クラスの生徒数 (index 0 = 1組)
 */
function getAllClassSizes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEET_CONFIG);
  var sizes = [];
  for (var i = 0; i < NUM_CLASSES; i++) {
    if (configSheet) {
      // 行3〜8 = クラス1〜6のデータ (i=0→行3, i=5→行8)
      var value = configSheet.getRange(i + 3, 2).getValue();
      sizes.push((value && !isNaN(value) && value > 0) ? Number(value) : DEFAULT_STUDENTS);
    } else {
      sizes.push(DEFAULT_STUDENTS);
    }
  }
  return sizes;
}

/**
 * 指定クラス番号のデータ開始行を返す（1-indexed）
 * @param {number} classNum - クラス番号 (1〜6)
 * @return {number} データ開始行
 */
function getClassStartRow(classNum) {
  var row = 6; // 行4=セパレータ, 行5=ヘッダー, 行6=最初の生徒
  var sizes = getAllClassSizes();
  for (var i = 0; i < classNum - 1; i++) {
    row += sizes[i];
    row += 2; // 区切り行 + 列ヘッダー
  }
  return row;
}

/**
 * 指定クラスの列ヘッダー行を返す
 */
function getClassHeaderRow(classNum) {
  return getClassStartRow(classNum) - 1;
}

/**
 * 指定クラスの区切り行を返す（全クラス対応）
 */
function getSeparatorRow(classNum) {
  return getClassHeaderRow(classNum) - 1;
}

/**
 * 翌日の日付を返す
 * @return {Date} 翌日の日付
 */
function _getNextSchoolDay() {
  var d = new Date();
  d.setDate(d.getDate() + 1);
  return d;
}

/**
 * 文字列を指定文字数で切り詰める（一覧シート用）
 * @param {string} text - 元の文字列
 * @param {number} maxLen - 最大文字数
 * @return {string} 切り詰められた文字列
 */
function _truncate(text, maxLen) {
  if (!text) return '';
  text = String(text).trim();
  if (text.length <= maxLen) return text;
  return text.substring(0, maxLen) + '…';
}

/**
 * 時刻値を HH:mm 形式の文字列に変換（一覧シート用）
 * GASのgetValues()はDateオブジェクトを返すため、フォーマットが必要
 * @param {Date|string} timeVal - 時刻値
 * @return {string} HH:mm 形式の文字列
 */
function _formatTime(timeVal) {
  if (!timeVal) return '';
  if (timeVal instanceof Date) {
    return Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm');
  }
  // 文字列の場合はそのまま返す
  return String(timeVal).trim();
}

/**
 * 行番号からクラス番号と生徒インデックス(0-based)を判定する
 * @param {number} row - スプレッドシートの行番号
 * @return {Object|null} {classNum, studentIndex} or null
 */
function getClassAndStudentFromRow(row) {
  var sizes = getAllClassSizes();
  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];
    if (row >= startRow && row < startRow + numStudents) {
      return { classNum: cls, studentIndex: row - startRow };
    }
  }
  return null;
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能A】ジャンプメニュー (アンチグラビティ・スクロール)
// ═══════════════════════════════════════════════════════════════════════

function jumpToClass1() { _jumpToClass(1); }
function jumpToClass2() { _jumpToClass(2); }
function jumpToClass3() { _jumpToClass(3); }
function jumpToClass4() { _jumpToClass(4); }
function jumpToClass5() { _jumpToClass(5); }
function jumpToClass6() { _jumpToClass(6); }

function jumpToAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ANALYSIS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠ 「分析」シートが見つかりません。');
    return;
  }
  ss.setActiveSheet(sheet);
  sheet.setActiveSelection('A1');
}

function _jumpToClass(classNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MAIN);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠ 「メイン」シートが見つかりません。');
    return;
  }
  ss.setActiveSheet(sheet);
  var targetRow = getClassStartRow(classNum);
  sheet.setActiveSelection('A' + targetRow);
  SpreadsheetApp.flush();
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能B】3日連続欠席アラート
// ═══════════════════════════════════════════════════════════════════════

/**
 * ログシートを参照し、今日を含めて3日連続で欠席している生徒の
 * メインシートC列に「🚨」を表示する。転校生は除外。
 */
function checkConsecutiveAbsence() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var logSheet  = ss.getSheetByName(SHEET_LOG);

  if (!mainSheet || !logSheet) {
    SpreadsheetApp.getUi().alert('⚠ 「メイン」または「ログ」シートが見つかりません。');
    return;
  }

  var sizes = getAllClassSizes();

  var logData = logSheet.getDataRange().getValues();
  if (logData.length <= 1) return;

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var targetDates = [];
  for (var d = 0; d < 3; d++) {
    var dt = new Date(today);
    dt.setDate(dt.getDate() - d);
    targetDates.push(dt.getTime());
  }

  var absenceMap = {};

  for (var i = 1; i < logData.length; i++) {
    var row = logData[i];
    var logDate = new Date(row[0]);
    logDate.setHours(0, 0, 0, 0);
    var logTime = logDate.getTime();

    if (targetDates.indexOf(logTime) === -1) continue;

    var classNum = row[1];
    var studentNo = row[2];
    var status = String(row[4]).trim();
    var reason = String(row[5] || '').trim();

    if (status === '欠席' && (reason === '体調不良' || reason === '家事都合')) {
      var key = classNum + '_' + studentNo;
      if (!absenceMap[key]) absenceMap[key] = [];
      if (absenceMap[key].indexOf(logTime) === -1) {
        absenceMap[key].push(logTime);
      }
    }
  }

  // 今日のメインシートのデータも加味（転校生は除外）
  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];
    for (var s = 0; s < numStudents; s++) {
      var rowNum = startRow + s;
      var status = String(mainSheet.getRange(rowNum, 4).getValue()).trim();
      var reason = String(mainSheet.getRange(rowNum, 5).getValue()).trim();

      if (status === STATUS_TRANSFERRED) continue; // 転校生は除外

      if (status === '欠席' && (reason === '体調不良' || reason === '家事都合')) {
        var key = cls + '_' + (s + 1);
        if (!absenceMap[key]) absenceMap[key] = [];
        var todayTime = today.getTime();
        if (absenceMap[key].indexOf(todayTime) === -1) {
          absenceMap[key].push(todayTime);
        }
      }
    }
  }

  // メインシートC列のクリアとアラート設定
  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];
    mainSheet.getRange(startRow, 3, numStudents, 1).clearContent();

    for (var s = 0; s < numStudents; s++) {
      // 転校生にはアラートを付けない
      var currentStatus = String(mainSheet.getRange(startRow + s, 4).getValue()).trim();
      if (currentStatus === STATUS_TRANSFERRED) continue;

      var key = cls + '_' + (s + 1);
      if (absenceMap[key] && absenceMap[key].length >= 3) {
        mainSheet.getRange(startRow + s, 3).setValue('🚨');
      }
    }
  }

  SpreadsheetApp.flush();
  ss.toast('🚨 連続欠席チェック完了', 'CYBER ALERT', 3);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能C】ホログラム・ダッシュボード
// ═══════════════════════════════════════════════════════════════════════

/**
 * 分析シートに出席データのグラフを描画（転校生は集計から除外）
 */
function refreshHologramCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var analysisSheet = ss.getSheetByName(SHEET_ANALYSIS);

  if (!mainSheet || !analysisSheet) {
    SpreadsheetApp.getUi().alert('⚠ シートが見つかりません。');
    return;
  }

  var sizes = getAllClassSizes();

  // 既存グラフを削除
  var existingCharts = analysisSheet.getCharts();
  for (var i = 0; i < existingCharts.length; i++) {
    analysisSheet.removeChart(existingCharts[i]);
  }

  // 集計データ作成（転校生を除外）
  var summaryData = [];
  var statusCounts = { '出席': 0, '欠席': 0, '遅刻': 0, '早退': 0, '遅刻早退': 0, '保健室': 0, '出停': 0 };
  var transferCounts = []; // クラスごとの転校者数

  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];
    var classCount = { '出席': 0, '欠席': 0, '遅刻': 0, '早退': 0, '遅刻早退': 0, '保健室': 0, '出停': 0, '合計': 0, '転校': 0 };

    for (var s = 0; s < numStudents; s++) {
      var status = String(mainSheet.getRange(startRow + s, 4).getValue()).trim();
      if (status === '' || status === undefined) status = '出席';

      if (status === STATUS_TRANSFERRED) {
        classCount['転校']++;
        continue; // 転校生は集計から除外
      }

      if (classCount.hasOwnProperty(status)) {
        classCount[status]++;
        statusCounts[status]++;
      }
      classCount['合計']++;
    }

    transferCounts.push(classCount['転校']);

    var attendRate = classCount['合計'] > 0
        ? Math.round((classCount['出席'] / classCount['合計']) * 100)
        : 0;
    summaryData.push([cls + '組', attendRate, classCount['出席'], classCount['欠席'],
                       classCount['遅刻'], classCount['早退'], classCount['遅刻早退'], classCount['保健室'], classCount['出停'],
                       classCount['合計'], classCount['転校']]);
  }

  // 分析シートの背景色設定
  analysisSheet.getRange(1, 1, analysisSheet.getMaxRows(), analysisSheet.getMaxColumns())
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE);

  // タイトル
  analysisSheet.getRange('A1').setValue('🔮 ホログラム・ダッシュボード')
    .setFontSize(16).setFontWeight('bold').setFontColor(COLOR.NEON_BLUE);
  analysisSheet.getRange('A2').setValue('📅 更新: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '　（※ 転校生は集計から除外）')
    .setFontSize(10).setFontColor(COLOR.CYBER_GREEN);

  // クラス別集計テーブル（転校者数・在籍数カラム追加）
  var tableHeaders = ['クラス', '出席率(%)', '出席', '欠席', '遅刻', '早退', '遅刻早退', '保健室', '出停', '在籍数', '転校'];
  for (var h = 0; h < tableHeaders.length; h++) {
    analysisSheet.getRange(4, h + 1).setValue(tableHeaders[h]).setFontWeight('bold');
  }
  analysisSheet.getRange(4, 1, 1, 11)
    .setBackground(COLOR.DARK_NAVY)
    .setFontColor(COLOR.CYBER_GREEN);

  for (var r = 0; r < summaryData.length; r++) {
    analysisSheet.getRange(5 + r, 1, 1, 11).setValues([summaryData[r]]);
  }

  // ステータス分布テーブル（円グラフ用）
  analysisSheet.getRange('A13').setValue('ステータス').setFontWeight('bold').setFontColor(COLOR.CYBER_GREEN);
  analysisSheet.getRange('B13').setValue('人数').setFontWeight('bold').setFontColor(COLOR.CYBER_GREEN);
  analysisSheet.getRange(13, 1, 1, 2).setBackground(COLOR.DARK_NAVY);

  var statusKeys = Object.keys(statusCounts);
  for (var k = 0; k < statusKeys.length; k++) {
    analysisSheet.getRange(14 + k, 1).setValue(statusKeys[k]);
    analysisSheet.getRange(14 + k, 2).setValue(statusCounts[statusKeys[k]]);
  }

  // グラフ1: クラス別出席率（縦棒グラフ）
  var barChart = analysisSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(analysisSheet.getRange('A4:B10'))
    .setPosition(4, 12, 0, 0)
    .setOption('title', '⚡ クラス別出席率')
    .setOption('titleTextStyle', { color: COLOR.NEON_BLUE, fontSize: 14, bold: true })
    .setOption('backgroundColor', { fill: COLOR.BG_BLACK })
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { textStyle: { color: COLOR.NEON_BLUE } })
    .setOption('vAxis', {
      textStyle: { color: COLOR.NEON_BLUE },
      gridlines: { color: '#1a1a2e' },
      minValue: 0, maxValue: 100
    })
    .setOption('colors', [COLOR.NEON_BLUE])
    .setOption('chartArea', { backgroundColor: { fill: COLOR.BG_BLACK } })
    .setOption('width', 500).setOption('height', 350)
    .build();
  analysisSheet.insertChart(barChart);

  // グラフ2: ステータス分布（円グラフ）
  var pieChart = analysisSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(analysisSheet.getRange('A13:B20'))
    .setPosition(4, 19, 0, 0)
    .setOption('title', '🔮 ステータス分布')
    .setOption('titleTextStyle', { color: COLOR.MAGENTA, fontSize: 14, bold: true })
    .setOption('backgroundColor', { fill: COLOR.BG_BLACK })
    .setOption('legend', { textStyle: { color: COLOR.NEON_BLUE }, position: 'right' })
    .setOption('pieSliceBorderColor', COLOR.BG_BLACK)
    .setOption('colors', [COLOR.CYBER_GREEN, '#ff0040', '#ffff00', COLOR.NEON_BLUE, '#aaff44', COLOR.ORANGE, COLOR.MAGENTA])
    .setOption('chartArea', { backgroundColor: { fill: COLOR.BG_BLACK } })
    .setOption('width', 500).setOption('height', 350)
    .build();
  analysisSheet.insertChart(pieChart);

  // グラフ3: クラス別詳細（積み上げ棒グラフ）
  var stackedChart = analysisSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(analysisSheet.getRange('A4:I10'))
    .setPosition(22, 12, 0, 0)
    .setOption('title', '📊 クラス別ステータス詳細')
    .setOption('titleTextStyle', { color: COLOR.CYBER_GREEN, fontSize: 14, bold: true })
    .setOption('backgroundColor', { fill: COLOR.BG_BLACK })
    .setOption('isStacked', true)
    .setOption('legend', { textStyle: { color: COLOR.NEON_BLUE }, position: 'top' })
    .setOption('hAxis', { textStyle: { color: COLOR.NEON_BLUE }, gridlines: { color: '#1a1a2e' } })
    .setOption('vAxis', { textStyle: { color: COLOR.NEON_BLUE } })
    .setOption('colors', [COLOR.NEON_BLUE, COLOR.CYBER_GREEN, '#ff0040', '#ffff00', '#aaff44', COLOR.ORANGE, COLOR.MAGENTA])
    .setOption('chartArea', { backgroundColor: { fill: COLOR.BG_BLACK } })
    .setOption('width', 900).setOption('height', 350)
    .build();
  analysisSheet.insertChart(stackedChart);

  SpreadsheetApp.flush();
  ss.toast('✨ ピコン！ ダッシュボード更新完了', 'HOLOGRAM SYSTEM', 3);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能D】デイリー・オートメーション
// ═══════════════════════════════════════════════════════════════════════

/**
 * 【手動実行】本日の出欠データをログシートに転記（転校生は除外）
 * 確認ダイアログ付き＋二重実行チェック
 */
function dailyArchive() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var logSheet  = ss.getSheetByName(SHEET_LOG);

  if (!mainSheet || !logSheet) {
    ui.alert('⚠ シートが見つかりません。');
    return;
  }

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  // ── 二重実行チェック ──
  if (logSheet.getLastRow() > 1) {
    var lastLogDate = logSheet.getRange(logSheet.getLastRow(), 1).getValue();
    if (String(lastLogDate) === today) {
      var overwrite = ui.alert(
        '⚠ 本日分は既にアーカイブ済みです',
        '本日（' + today + '）のデータは既にログに記録されています。\n再度アーカイブすると重複データが作成されます。\n\n続行しますか？',
        ui.ButtonSet.YES_NO
      );
      if (overwrite !== ui.Button.YES) return;
    }
  }

  // ── 確認ダイアログ ──
  var confirm = ui.alert(
    '📦 アーカイブ確認',
    '本日（' + today + '）の出欠データをログに保存しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  _dailyArchiveInternal(ss, mainSheet);
}

/**
 * アーカイブの内部実装（ダイアログなし）
 * dailyArchiveAndResetから直接呼び出せる
 */
function _dailyArchiveInternal(ss, mainSheet) {
  var logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) return;

  var sizes = getAllClassSizes();
  // メインシートのA3セルから日付を取得（ログにはその日の日付を記録）
  var mainDateRaw = String(mainSheet.getRange('A3').getValue()).replace('📅 ', '');
  var today = mainDateRaw.replace(/\s*\(.*\)/, '');  // 曜日部分を除去 → yyyy/MM/dd

  // ── ログヘッダー ──
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['日付', 'クラス', 'No', '氏名', 'ステータス', '理由', '備考', '時刻']);
    logSheet.getRange(1, 1, 1, 8)
      .setBackground(COLOR.DARK_NAVY)
      .setFontColor(COLOR.CYBER_GREEN)
      .setFontWeight('bold');
  }

  // ── データ転記 ──
  var archivedCount = 0;
  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];

    for (var s = 0; s < numStudents; s++) {
      var rowData = mainSheet.getRange(startRow + s, 1, 1, NUM_COLS).getValues()[0];
      var studentNo = rowData[0];
      var name      = rowData[1];
      var status    = rowData[3];
      var reason    = rowData[4];
      var note      = rowData[5];
      var time      = rowData[6];

      if (!name || String(name).trim() === '') continue;
      if (String(status).trim() === STATUS_TRANSFERRED) continue;
      if (!status || String(status).trim() === '') status = '出席';

      logSheet.appendRow([today, cls, studentNo, name, status, reason, note, time]);
      archivedCount++;
    }
  }

  var lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    logSheet.getRange(2, 1, lastRow - 1, 8)
      .setBackground(COLOR.BG_BLACK)
      .setFontColor(COLOR.NEON_BLUE);
  }

  SpreadsheetApp.flush();
  ss.toast('📦 ' + archivedCount + '名分のデータをアーカイブしました（' + today + '）', 'DAILY ARCHIVE', 5);
}

/**
 * 【手動実行】メインシートの入力欄をクリア（転校生はスキップ）
 * 新しい日の出欠記録を始める前に実行する
 */
function dailyReset() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  // ── 確認ダイアログ ──
  // メインシートの現在の日付から翌日を計算
  var currentDateStr = String(mainSheet.getRange('A3').getValue()).replace('📅 ', '');
  var currentDate = new Date(currentDateStr.replace(/\s*\(.*\)/, '').replace(/\//g, '-'));
  currentDate.setDate(currentDate.getDate() + 1);
  var todayStr = Utilities.formatDate(currentDate, 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  var confirm = ui.alert(
    '🔄 デイリーリセット確認',
    '全生徒のステータスを「出席」に戻し、理由・備考欄をクリアします。\n日付を「' + todayStr + '」に更新します。\n\n⚠ アーカイブは済んでいますか？\n（未保存のデータは失われます）\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  _dailyResetInternal(ss, mainSheet);
}

/**
 * リセットの内部実装（ダイアログなし）
 * dailyArchiveAndResetから直接呼び出せる
 */
function _dailyResetInternal(ss, mainSheet) {
  // メインシートの現在の日付から翌日を計算
  var currentDateStr = String(mainSheet.getRange('A3').getValue()).replace('📅 ', '');
  var currentDate = new Date(currentDateStr.replace(/\s*\(.*\)/, '').replace(/\//g, '-'));
  currentDate.setDate(currentDate.getDate() + 1);
  var nextDayStr = Utilities.formatDate(currentDate, 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  var sizes = getAllClassSizes();

  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];

    for (var s = 0; s < numStudents; s++) {
      var rowNum = startRow + s;
      var currentStatus = String(mainSheet.getRange(rowNum, 4).getValue()).trim();

      if (currentStatus === STATUS_TRANSFERRED) continue;

      mainSheet.getRange(rowNum, 5, 1, 3).clearContent();
      mainSheet.getRange(rowNum, 4).setValue('出席');
      mainSheet.getRange(rowNum, 3).clearContent();
    }
  }

  mainSheet.getRange('A3').setValue('📅 ' + nextDayStr);

  // 一覧シートもリセット状態に更新
  var overviewSheet = ss.getSheetByName(SHEET_OVERVIEW);
  if (overviewSheet) {
    overviewSheet.getRange(1, 1).setValue('出欠表　　' + nextDayStr);
    _refreshOverviewData(overviewSheet);
  }

  SpreadsheetApp.flush();
  ss.toast('🔄 リセット完了！日付: ' + nextDayStr, 'SYSTEM RESET', 3);
}

/**
 * アーカイブ + リセットを一括実行（確認ダイアログは1回だけ）
 */
function dailyArchiveAndReset() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  // メインシートの現在の日付から翌日を計算
  var currentDateStr = String(mainSheet.getRange('A3').getValue()).replace('📅 ', '');
  var currentDate = new Date(currentDateStr.replace(/\s*\(.*\)/, '').replace(/\//g, '-'));
  currentDate.setDate(currentDate.getDate() + 1);
  var nextDayStr = Utilities.formatDate(currentDate, 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  var confirm = ui.alert(
    '📦+🔄 アーカイブ＆リセット',
    '以下を一括実行します：\n\n1️⃣ 本日分をログにアーカイブ\n2️⃣ 全生徒を「出席」にリセット\n3️⃣ 日付を「' + nextDayStr + '」に更新\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // アーカイブ（ダイアログなし版）
  _dailyArchiveInternal(ss, mainSheet);

  // リセット（ダイアログなし版）
  _dailyResetInternal(ss, mainSheet);
}

// ═══════════════════════════════════════════════════════════════════════
//  【機能H】アーカイブから復元
// ═══════════════════════════════════════════════════════════════════════

/**
 * ログシートの最新アーカイブからメインシートのデータを復元する
 * 誤ってアーカイブ＆リセットを実行した場合の救済機能
 */
function restoreFromArchive() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var logSheet  = ss.getSheetByName(SHEET_LOG);

  if (!mainSheet || !logSheet) {
    ui.alert('⚠ シートが見つかりません。');
    return;
  }

  var lastRow = logSheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('⚠ ログにアーカイブデータがありません。');
    return;
  }

  // ログの最新日付を取得
  var allLogData = logSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var latestDate = String(allLogData[allLogData.length - 1][0]);

  // 確認ダイアログ
  var confirm = ui.alert(
    '↩ アーカイブから復元',
    '「' + latestDate + '」のアーカイブデータをメインシートに復元します。\n\n❗ 現在のメインシートのデータは上書きされます。\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // 最新日付のデータだけ抽出
  var restoreData = [];
  for (var i = 0; i < allLogData.length; i++) {
    if (String(allLogData[i][0]) === latestDate) {
      restoreData.push(allLogData[i]);
    }
  }

  // クラスごとにデータを復元
  var sizes = getAllClassSizes();
  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];

    // まずクラスのデータをクリア（転校生以外）
    for (var s = 0; s < numStudents; s++) {
      var rowNum = startRow + s;
      var currentStatus = String(mainSheet.getRange(rowNum, 4).getValue()).trim();
      if (currentStatus === STATUS_TRANSFERRED) continue;

      mainSheet.getRange(rowNum, 4).setValue('出席');
      mainSheet.getRange(rowNum, 5, 1, 3).clearContent();
      mainSheet.getRange(rowNum, 3).clearContent();
    }

    // ログから該当クラスのデータを復元
    for (var r = 0; r < restoreData.length; r++) {
      var logRow = restoreData[r];
      var logClass  = Number(logRow[1]);
      var logNo     = Number(logRow[2]);
      var logStatus = logRow[4];
      var logReason = logRow[5];
      var logNote   = logRow[6];
      var logTime   = logRow[7];

      if (logClass !== cls) continue;
      if (logNo < 1 || logNo > numStudents) continue;

      var targetRow = startRow + logNo - 1;
      mainSheet.getRange(targetRow, 4).setValue(logStatus || '出席');
      mainSheet.getRange(targetRow, 5).setValue(logReason || '');
      mainSheet.getRange(targetRow, 6).setValue(logNote || '');
      if (logTime) {
        mainSheet.getRange(targetRow, 7).setValue(logTime);
      }
    }
  }

  // 日付を復元
  var dateForDisplay = latestDate.replace(/-/g, '/');
  var restoreDate = new Date(latestDate);
  var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
  var dayOfWeek = dayNames[restoreDate.getDay()];
  mainSheet.getRange('A3').setValue('📅 ' + dateForDisplay + ' (' + dayOfWeek + ')');

  SpreadsheetApp.flush();
  ss.toast('↩ 「' + latestDate + '」のデータを復元しました', 'RESTORE', 5);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能G】転入生・転校生管理
// ═══════════════════════════════════════════════════════════════════════

/**
 * 転入生を追加する
 * ダイアログでクラス番号と氏名を入力 → そのクラスの末尾に新しい番号で追加
 * 設定シートの人数も自動で+1
 */
function addTransferStudent() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var configSheet = ss.getSheetByName(SHEET_CONFIG);

  if (!mainSheet || !configSheet) {
    ui.alert('⚠ シートが見つかりません。先にセットアップを実行してください。');
    return;
  }

  // クラス選択
  var classResponse = ui.prompt(
    '👤 転入生追加',
    'クラス番号を入力してください（1〜' + NUM_CLASSES + '）:',
    ui.ButtonSet.OK_CANCEL
  );
  if (classResponse.getSelectedButton() !== ui.Button.OK) return;

  var classNum = parseInt(classResponse.getResponseText().trim(), 10);
  if (isNaN(classNum) || classNum < 1 || classNum > NUM_CLASSES) {
    ui.alert('⚠ 無効なクラス番号です。');
    return;
  }

  // 氏名入力
  var nameResponse = ui.prompt(
    '👤 転入生追加 (' + classNum + '組)',
    '転入生の氏名を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;

  var studentName = nameResponse.getResponseText().trim();
  if (!studentName) {
    ui.alert('⚠ 氏名が入力されていません。');
    return;
  }

  // 現在のクラス人数を取得
  var currentSize = getStudentsPerClass(classNum);
  var newNo = currentSize + 1;

  // 設定シートの人数を+1（行3〜8がクラス1〜6）
  configSheet.getRange(classNum + 2, 2).setValue(newNo);

  // ── ここからシートを再構築 ──
  // 新しい人数に基づいてメインシートを再構築する代わりに、
  // 現在の該当クラスの最終行の直後に1行挿入する方式を使う

  // 現在のクラスデータ最終行（旧人数ベース）
  var oldStartRow = _getClassStartRowDirect(classNum, getAllClassSizes());
  // ※ 注意: getAllClassSizesは既にconfigSheetを更新した後なので、
  //   直前の人数でのStartRowを計算するために-1する必要がある
  var sizes = getAllClassSizes();
  sizes[classNum - 1] = currentSize; // 元の人数で計算
  var insertAfterRow = _getClassStartRowWithSizes(classNum, sizes) + currentSize - 1;

  // 行を挿入
  mainSheet.insertRowAfter(insertAfterRow);
  var newRow = insertAfterRow + 1;

  // 新しい行のデータを設定
  mainSheet.getRange(newRow, 1).setValue(newNo);
  mainSheet.getRange(newRow, 2).setValue(studentName);
  mainSheet.getRange(newRow, 3).setValue(''); // 🚨
  mainSheet.getRange(newRow, 4).setValue('出席');
  mainSheet.getRange(newRow, 5).setValue(''); // 理由
  mainSheet.getRange(newRow, 6).setValue('転入'); // 備考に「転入」と記載
  mainSheet.getRange(newRow, 7).setValue(''); // 時刻

  // スタイルを設定
  mainSheet.getRange(newRow, 1, 1, NUM_COLS)
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono');

  // プルダウン設定
  var statusOptions = SpreadsheetApp.newDataValidation()
    .requireValueInList(ALL_STATUSES, true)
    .setAllowInvalid(false)
    .build();
  mainSheet.getRange(newRow, 4).setDataValidation(statusOptions);

  // グリッド罫線
  mainSheet.getRange(newRow, 1, 1, NUM_COLS)
    .setBorder(null, null, null, null, true, true, '#1a3a1a', SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.flush();

  // 条件付き書式を再適用
  _applyConditionalFormatting(mainSheet, getAllClassSizes());

  ss.toast('👤 ' + classNum + '組に転入生「' + studentName + '」(No.' + newNo + ') を追加しました', 'TRANSFER IN', 5);

  // 追加した行にジャンプ
  mainSheet.setActiveSelection('B' + newRow);
}

/**
 * 指定サイズ配列を使ってクラスのデータ開始行を計算（内部用）
 */
function _getClassStartRowWithSizes(classNum, sizes) {
  var row = 6; // getClassStartRowと同じ起点
  for (var i = 0; i < classNum - 1; i++) {
    row += sizes[i];
    row += 2;
  }
  return row;
}

/**
 * 直接サイズ指定でクラス開始行を取得（内部用）
 */
function _getClassStartRowDirect(classNum, sizes) {
  return _getClassStartRowWithSizes(classNum, sizes);
}

/**
 * 選択中のセルの生徒を「転校」扱いにする
 * ステータスを「転校」に変更し、行をグレーアウトする
 */
function markAsTransferred() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);

  if (!mainSheet) {
    ui.alert('⚠ 「メイン」シートが見つかりません。');
    return;
  }

  // 現在のアクティブシートがメインシートか確認
  if (ss.getActiveSheet().getName() !== SHEET_MAIN) {
    ui.alert('⚠ メインシートを開いた状態で実行してください。');
    return;
  }

  var activeRange = mainSheet.getActiveRange();
  var row = activeRange.getRow();

  // どのクラスの何番目の生徒か判定
  var info = getClassAndStudentFromRow(row);
  if (!info) {
    ui.alert('⚠ 生徒のデータ行を選択してください。\n（ヘッダーや区切り行は対象外です）');
    return;
  }

  var studentNo = mainSheet.getRange(row, 1).getValue();
  var studentName = mainSheet.getRange(row, 2).getValue();

  if (!studentName || String(studentName).trim() === '') {
    ui.alert('⚠ この行には生徒が登録されていません。');
    return;
  }

  // 確認ダイアログ
  var confirmResult = ui.alert(
    '🚪 転校処理の確認',
    info.classNum + '組 No.' + studentNo + ' 「' + studentName + '」を転校扱いにしますか？\n\n' +
    '※ 行がグレーアウトされ、出欠集計から除外されます。\n' +
    '※ 元に戻すにはD列のプルダウンを「出席」等に変更してください。',
    ui.ButtonSet.YES_NO
  );

  if (confirmResult !== ui.Button.YES) return;

  // ステータスを「転校」に変更
  mainSheet.getRange(row, 4).setValue(STATUS_TRANSFERRED);

  // 理由・備考・時刻をクリア、C列(🚨)もクリア
  mainSheet.getRange(row, 3).clearContent();
  mainSheet.getRange(row, 5, 1, 3).clearContent();

  // 行をグレーアウト（条件付き書式でも対応するが、即座に見た目を変える）
  mainSheet.getRange(row, 1, 1, NUM_COLS)
    .setBackground(COLOR.GRAY)
    .setFontColor(COLOR.GRAY_TEXT);

  SpreadsheetApp.flush();
  ss.toast('🚪 ' + info.classNum + '組 No.' + studentNo + ' 「' + studentName + '」を転校処理しました', 'TRANSFER OUT', 5);
}

/**
 * 転校処理を取り消す（元に戻す）
 * 選択中のセルの生徒が「転校」ステータスの場合、「出席」に戻す
 */
function undoTransfer() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);

  if (!mainSheet || ss.getActiveSheet().getName() !== SHEET_MAIN) {
    ui.alert('⚠ メインシートを開いた状態で実行してください。');
    return;
  }

  var row = mainSheet.getActiveRange().getRow();
  var info = getClassAndStudentFromRow(row);
  if (!info) {
    ui.alert('⚠ 生徒のデータ行を選択してください。');
    return;
  }

  var currentStatus = String(mainSheet.getRange(row, 4).getValue()).trim();
  if (currentStatus !== STATUS_TRANSFERRED) {
    ui.alert('⚠ この生徒は転校ステータスではありません。');
    return;
  }

  var studentNo = mainSheet.getRange(row, 1).getValue();
  var studentName = mainSheet.getRange(row, 2).getValue();

  // ステータスを「出席」に戻す
  mainSheet.getRange(row, 4).setValue('出席');

  // 背景色を元に戻す
  var bgColor = (info.studentIndex % 2 === 1) ? '#050510' : COLOR.BG_BLACK;
  mainSheet.getRange(row, 1, 1, NUM_COLS)
    .setBackground(bgColor)
    .setFontColor(COLOR.NEON_BLUE);

  SpreadsheetApp.flush();
  ss.toast('↩ ' + info.classNum + '組 No.' + studentNo + ' 「' + studentName + '」の転校を取り消しました', 'UNDO TRANSFER', 5);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能E】初期セットアップ
// ═══════════════════════════════════════════════════════════════════════

/**
 * スプレッドシートの初期設定を一括実行
 */
function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var configSheet   = ss.getSheetByName(SHEET_CONFIG)   || ss.insertSheet(SHEET_CONFIG);
  var mainSheet     = ss.getSheetByName(SHEET_MAIN)     || ss.insertSheet(SHEET_MAIN);
  var logSheet      = ss.getSheetByName(SHEET_LOG)      || ss.insertSheet(SHEET_LOG);
  var analysisSheet = ss.getSheetByName(SHEET_ANALYSIS) || ss.insertSheet(SHEET_ANALYSIS);
  var flowSheet     = ss.getSheetByName(SHEET_FLOW)     || ss.insertSheet(SHEET_FLOW);
  var overviewSheet = ss.getSheetByName(SHEET_OVERVIEW) || ss.insertSheet(SHEET_OVERVIEW);

  // デフォルトの「シート1」があれば削除
  var defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // 設定シートの構築
  _setupConfigSheet(configSheet);

  // クラス人数を取得
  var sizes = getAllClassSizes();

  // メインシートの行数を確保
  var totalRows = 3; // 行1-3: タイトル, メニュー案内, 日付
  for (var i = 0; i < NUM_CLASSES; i++) {
    totalRows += sizes[i] + 2; // セパレータ + ヘッダー + 生徒数
  }
  totalRows += 10; // バッファ
  if (mainSheet.getMaxRows() < totalRows) {
    mainSheet.insertRowsAfter(mainSheet.getMaxRows(), totalRows - mainSheet.getMaxRows());
  }
  if (mainSheet.getMaxColumns() < NUM_COLS) {
    mainSheet.insertColumnsAfter(mainSheet.getMaxColumns(), NUM_COLS - mainSheet.getMaxColumns());
  }

  // 全体の背景色・フォント色
  mainSheet.getRange(1, 1, mainSheet.getMaxRows(), mainSheet.getMaxColumns())
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono');

  // データエリア（行4以降）を完全クリア（旧セパレータの残骸防止）
  var lastRow = mainSheet.getMaxRows();
  if (lastRow >= 4) {
    mainSheet.getRange(4, 1, lastRow - 3, mainSheet.getMaxColumns()).breakApart();
    mainSheet.getRange(4, 1, lastRow - 3, NUM_COLS).clearContent();
    mainSheet.getRange(4, 1, lastRow - 3, NUM_COLS).clearDataValidations();
  }

  // 列幅設定
  mainSheet.setColumnWidth(1, 45);
  mainSheet.setColumnWidth(2, 130);
  mainSheet.setColumnWidth(3, 45);
  mainSheet.setColumnWidth(4, 110);
  mainSheet.setColumnWidth(5, 160);
  mainSheet.setColumnWidth(6, 160);
  mainSheet.setColumnWidth(7, 90);

  // 行1: メインタイトル
  mainSheet.getRange('A1').setValue('🛸出欠管理システム 例の「アレ」');
  mainSheet.getRange(1, 1, 1, NUM_COLS)
    .merge()
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor(COLOR.NEON_BLUE)
    .setBackground(COLOR.HEADER_BG)
    .setHorizontalAlignment('center');

  // 行2: メニュー案内
  mainSheet.getRange('A2').setValue('📋 メニュー「🛸 コマンドセンター」から全機能にアクセスできます');
  mainSheet.getRange(2, 1, 1, NUM_COLS)
    .merge()
    .setFontSize(9)
    .setFontColor(COLOR.CYBER_GREEN)
    .setBackground(COLOR.HEADER_BG)
    .setHorizontalAlignment('center');

  // 行3: 日付
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  mainSheet.getRange('A3').setValue('📅 ' + todayStr);
  mainSheet.getRange(3, 1, 1, NUM_COLS)
    .merge()
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor(COLOR.MAGENTA)
    .setBackground(COLOR.HEADER_BG);

  // プルダウン選択肢（転校を含む）
  var statusOptions = SpreadsheetApp.newDataValidation()
    .requireValueInList(ALL_STATUSES, true)
    .setAllowInvalid(false)
    .build();

  // 理由プルダウン
  var reasonOptions = SpreadsheetApp.newDataValidation()
    .requireValueInList(ALL_REASONS, true)
    .setAllowInvalid(true) // 手入力も許可するか（今回はリスト選択推奨だが、自由記述も残すならtrue。指定はプルダウンにしてくれとのことなのでfalseが基本だが、その他があるので詳細を書くならtrueの方がいいかも？いや、リストにするという要望なのでfalseか、あるいは「その他」を選んで備考に書く運用か。とりあえずfalseにするが、もしリスト外を許容したければtrue）
    // 要望は「下記のプルダウンにしてください」なのでリスト内選択のみ(false)が無難。
    .setAllowInvalid(false)
    .build();

  // 各クラスの枠を生成
  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var headerRow = getClassHeaderRow(cls);
    var startRow  = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];

    // 区切りライン（全クラス共通）
    var sepRow = getSeparatorRow(cls);
    mainSheet.getRange(sepRow, 1, 1, NUM_COLS)
      .merge()
      .setValue('▓▓▓▓▓▓▓▓▓▓▓▓ ★ ' + cls + ' 組 (' + numStudents + '名) ★ ▓▓▓▓▓▓▓▓▓▓▓▓')
      .setBackground(COLOR.SEPARATOR_BG)
      .setFontColor(COLOR.CYBER_GREEN)
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    mainSheet.getRange(sepRow, 1, 1, NUM_COLS)
      .setBorder(true, null, true, null, null, null, COLOR.CYBER_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);

    // 列ヘッダー
    var headers = ['No', '氏名', '🚨', 'ステータス', '理由', '備考', '時刻'];
    mainSheet.getRange(headerRow, 1, 1, NUM_COLS).setValues([headers]);
    mainSheet.getRange(headerRow, 1, 1, NUM_COLS)
      .setBackground(COLOR.DARK_NAVY)
      .setFontColor(COLOR.CYBER_GREEN)
      .setFontWeight('bold')
      .setFontSize(10)
      .setBorder(true, null, true, null, null, null, COLOR.CYBER_GREEN, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // データ行の初期設定
    for (var s = 0; s < numStudents; s++) {
      var row = startRow + s;
      // 既にデータがある行はスキップ（再セットアップ時にデータを保持）
      var existingName = mainSheet.getRange(row, 2).getValue();
      var existingStatus = String(mainSheet.getRange(row, 4).getValue()).trim();

      mainSheet.getRange(row, 1).setValue(s + 1);
      mainSheet.getRange(row, 4).setDataValidation(statusOptions);
      mainSheet.getRange(row, 5).setDataValidation(reasonOptions); // 理由プルダウンを設定
      // フォントサイズを確実に10ptにする（セパレータの影響防止）
      mainSheet.getRange(row, 1, 1, NUM_COLS).setFontSize(10);
      // 水平配置を明示的に設定（ズレ防止）
      // No:中, 氏名:左, 🚨:中, ステータス:中, 理由:左, 備考:左, 時刻:中
      mainSheet.getRange(row, 1, 1, NUM_COLS).setHorizontalAlignments([['center', 'left', 'center', 'center', 'left', 'left', 'center']]);

      if (!existingName || String(existingName).trim() === '') {
        // 新規行: デフォルスト「出席」を設定
        mainSheet.getRange(row, 4).setValue('出席');
      }

      // 転校生の行はグレーアウトを維持
      if (existingStatus === STATUS_TRANSFERRED) {
        mainSheet.getRange(row, 1, 1, NUM_COLS)
          .setBackground(COLOR.GRAY)
          .setFontColor(COLOR.GRAY_TEXT);
      }
    }

    // G列（時刻）の表示形式を HH:mm に設定
    mainSheet.getRange(startRow, 7, numStudents, 1).setNumberFormat('HH:mm');

    // データ行にグリッド罫線
    mainSheet.getRange(startRow, 1, numStudents, NUM_COLS)
      .setBorder(null, null, null, null, true, true, '#1a3a1a', SpreadsheetApp.BorderStyle.SOLID);

    // 偶数行に微妙な背景色差（転校生以外）
    for (var s = 0; s < numStudents; s++) {
      var existingStatus = String(mainSheet.getRange(startRow + s, 4).getValue()).trim();
      if (existingStatus === STATUS_TRANSFERRED) continue;
      if (s % 2 === 1) {
        mainSheet.getRange(startRow + s, 1, 1, NUM_COLS)
          .setBackground('#050510');
      }
    }
  }

  // 条件付き書式の設定
  _applyConditionalFormatting(mainSheet, sizes);

  // ログシートの初期設定
  logSheet.getRange(1, 1, logSheet.getMaxRows(), logSheet.getMaxColumns())
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono');

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['日付', 'クラス', 'No', '氏名', 'ステータス', '理由', '備考', '時刻']);
    logSheet.getRange(1, 1, 1, 8)
      .setBackground(COLOR.DARK_NAVY)
      .setFontColor(COLOR.CYBER_GREEN)
      .setFontWeight('bold');
  }

  // 分析シートの初期設定
  analysisSheet.getRange(1, 1, analysisSheet.getMaxRows(), analysisSheet.getMaxColumns())
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono');

  analysisSheet.getRange('A1').setValue('🔮 ホログラム・ダッシュボード')
    .setFontSize(16).setFontWeight('bold').setFontColor(COLOR.NEON_BLUE);
  analysisSheet.getRange('A2').setValue('📊 「ダッシュボード更新」')
    .setFontSize(10).setFontColor(COLOR.CYBER_GREEN);

  // 運用フローシートの構築
  _setupFlowSheet(flowSheet);

  // 一覧シートの初期構築
  _setupOverviewSheet(overviewSheet, sizes);

  SpreadsheetApp.flush();
  ss.toast('🚀 セットアップ完了！ 人数変更は「設定」シートで。', 'SYSTEM READY', 5);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能F】設定シート管理
// ═══════════════════════════════════════════════════════════════════════

/**
 * 設定シートの初期構築（内部関数）
 */
function _setupConfigSheet(configSheet) {
  configSheet.getRange(1, 1, configSheet.getMaxRows(), configSheet.getMaxColumns())
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono');

  configSheet.getRange('A1').setValue('⚙ クラス別人数設定');
  configSheet.getRange(1, 1, 1, 3).merge()
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor(COLOR.NEON_BLUE)
    .setBackground(COLOR.HEADER_BG);

  configSheet.getRange('A2').setValue('クラス');
  configSheet.getRange('B2').setValue('人数');
  configSheet.getRange('C2').setValue('メモ');
  configSheet.getRange(2, 1, 1, 3)
    .setBackground(COLOR.DARK_NAVY)
    .setFontColor(COLOR.CYBER_GREEN)
    .setFontWeight('bold');

  for (var i = 1; i <= NUM_CLASSES; i++) {
    // 行1=タイトル, 行2=ヘッダー, 行3〜8=クラス1〜6のデータ
    var row = i + 2;
    configSheet.getRange(row, 1).setValue(i + '組');

    var currentValue = configSheet.getRange(row, 2).getValue();
    if (!currentValue || currentValue === '') {
      configSheet.getRange(row, 2).setValue(DEFAULT_STUDENTS);
    }

    var numValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 50)
      .setHelpText('1〜50の数値を入力してください')
      .setAllowInvalid(false)
      .build();
    configSheet.getRange(row, 2).setDataValidation(numValidation);
  }

  configSheet.getRange('A10').setValue('📝 人数変更後は「🛸 コマンドセンター」→「🔄 初期セットアップ」を再実行');
  configSheet.getRange(10, 1, 1, 3).merge()
    .setFontSize(9)
    .setFontColor(COLOR.ORANGE);

  configSheet.getRange('A11').setValue('📝 転入生は「🛸 コマンドセンター」→「👤 転入生追加」で自動追加されます');
  configSheet.getRange(11, 1, 1, 3).merge()
    .setFontSize(9)
    .setFontColor(COLOR.ORANGE);

  configSheet.setColumnWidth(1, 80);
  configSheet.setColumnWidth(2, 80);
  configSheet.setColumnWidth(3, 200);

  configSheet.getRange(2, 1, NUM_CLASSES + 1, 3)
    .setBorder(true, true, true, true, true, true, COLOR.CYBER_GREEN, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * 条件付き書式を適用（内部関数）
 * 転校 → グレーアウト の条件付き書式を追加
 */
function _applyConditionalFormatting(sheet, sizes) {
  sheet.clearConditionalFormatRules();
  var rules = [];

  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];
    var range = sheet.getRange(startRow, 1, numStudents, NUM_COLS);

    // 転校 → グレーアウト（最優先）
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D' + startRow + '="転校"')
      .setBackground(COLOR.GRAY)
      .setFontColor(COLOR.GRAY_TEXT)
      .setStrikethrough(true)
      .setRanges([range])
      .build());

    // 欠席 → ダークレッド
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D' + startRow + '="欠席"')
      .setBackground(COLOR.DARK_RED)
      .setFontColor('#ff4444')
      .setRanges([range])
      .build());

    // 遅刻 → ダークイエロー
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D' + startRow + '="遅刻"')
      .setBackground(COLOR.DARK_YELLOW)
      .setFontColor('#ffff44')
      .setRanges([range])
      .build());

    // 早退 → ダークシアン
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D' + startRow + '="早退"')
      .setBackground(COLOR.DARK_CYAN)
      .setFontColor('#44ffff')
      .setRanges([range])
      .build());

    // 遅刻早退 → ダークライム
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D' + startRow + '="遅刻早退"')
      .setBackground(COLOR.DARK_LIME)
      .setFontColor('#aaff44')
      .setRanges([range])
      .build());

    // 保健室 → ダークオレンジ
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D' + startRow + '="保健室"')
      .setBackground(COLOR.DARK_ORANGE)
      .setFontColor(COLOR.ORANGE)
      .setRanges([range])
      .build());

    // 出停 → ダークマゼンタ
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D' + startRow + '="出停"')
      .setBackground(COLOR.DARK_MAGENTA)
      .setFontColor(COLOR.MAGENTA)
      .setRanges([range])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
}


// ═══════════════════════════════════════════════════════════════════════
//  カスタムメニュー
// ═══════════════════════════════════════════════════════════════════════

/**
 * スプレッドシートを開いた時に自動でカスタムメニューを追加し、日付を更新する
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // ジャンプサブメニュー
  var jumpMenu = ui.createMenu('🚀 クラスにジャンプ')
    .addItem('1組へ', 'jumpToClass1')
    .addItem('2組へ', 'jumpToClass2')
    .addItem('3組へ', 'jumpToClass3')
    .addItem('4組へ', 'jumpToClass4')
    .addItem('5組へ', 'jumpToClass5')
    .addItem('6組へ', 'jumpToClass6');

  ui.createMenu('🛸 コマンドセンター')
    .addSubMenu(jumpMenu)
    .addItem('📊 分析シートへ', 'jumpToAnalysis')
    .addItem('📖 運用フローへ', 'jumpToFlow')
    .addItem('🗂 一覧シートへ', 'jumpToOverview')
    .addSeparator()
    .addItem('🔄 一覧を更新', 'refreshOverview')
    .addItem('⏱ 一覧の自動更新ON', 'setupAutoRefresh')
    .addItem('⏹ 一覧の自動更新OFF', 'removeAutoRefresh')
    .addSeparator()
    .addItem('📦 本日分をアーカイブ', 'dailyArchive')
    .addItem('🔄 翌日用にリセット', 'dailyReset')
    .addItem('📦+🔄 アーカイブ＆リセット', 'dailyArchiveAndReset')
    .addItem('↩ アーカイブから復元', 'restoreFromArchive')
    .addSeparator()
    .addItem('🚨 連続欠席チェック', 'checkConsecutiveAbsence')
    .addItem('📊 ダッシュボード更新', 'refreshHologramCharts')
    .addSeparator()
    .addItem('📋 本日の欠席者リスト', 'showTodayAbsentList')
    .addItem('📋 日付指定で欠席者リスト', 'showDateAbsentList')
    .addItem('🔍 生徒の出欠履歴', 'showStudentHistory')
    .addItem('📊 月間出欠集計', 'monthlyAttendanceSummary')
    .addItem('📅 学期末集計', 'semesterSummary')
    .addSeparator()
    .addItem('👤 転入生追加', 'addTransferStudent')
    .addItem('🚪 転校処理（選択行）', 'markAsTransferred')
    .addItem('↩ 転校取り消し（選択行）', 'undoTransfer')
    .addSeparator()
    .addItem('🔄 初期セットアップ', 'setupSpreadsheet')
    .addItem('⚙ 人数設定シートを開く', 'jumpToConfig')
    .addToUi();

  // 日付の自動更新は廃止（ユーザー要望により手動「アーカイブ＆リセット」で更新）
  // _updateDateDisplay();

  // 一覧シートの日付も更新（メインシートのA3セルから取得）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var overviewSheet = ss.getSheetByName(SHEET_OVERVIEW);
  if (overviewSheet) {
    var mainSheet = ss.getSheetByName(SHEET_MAIN);
    var mainDateStr = mainSheet ? String(mainSheet.getRange('A3').getValue()).replace('📅 ', '') : '';
    if (mainDateStr) {
      overviewSheet.getRange(1, 1).setValue('出欠表　　' + mainDateStr);
    }
  }

  ss.toast('🛸 出欠管理システム « 例の"アレ" » 稼働中...', 'SYSTEM INITIALIZED', 5);
}

/**
 * メインシートの日付表示を今日の日付に更新する（内部関数）
 */
function _updateDateDisplay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  // 日付を更新
  mainSheet.getRange('A3').setValue('📅 ' + todayStr);
}

function jumpToConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠ 「設定」シートが見つかりません。\n先に「初期セットアップ」を実行してください。');
    return;
  }
  ss.setActiveSheet(sheet);
  sheet.setActiveSelection('B3');
}

function jumpToFlow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_FLOW);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠ 「運用フロー」シートが見つかりません。\n先に「初期セットアップ」を実行してください。');
    return;
  }
  ss.setActiveSheet(sheet);
  sheet.setActiveSelection('A1');
}

function jumpToOverview() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_OVERVIEW);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠ 「一覧」シートが見つかりません。\n先に「初期セットアップ」を実行してください。');
    return;
  }
  ss.setActiveSheet(sheet);
  sheet.setActiveSelection('A1');
}


// ═══════════════════════════════════════════════════════════════════════
//  一覧シート（全クラス一目表示）
// ═══════════════════════════════════════════════════════════════════════

/** 一覧シートの1クラスあたりの列数: No, 氏名, 状態, 備考 */
var OVERVIEW_COLS_PER_CLASS = 4;

/**
 * 一覧シートの初期構築（内部関数）
 * 画像のように6クラスを横並びで一目で全体把握できるシート
 */
function _setupOverviewSheet(sheet, sizes) {
  // 既存のマージセルをすべて解除（再セットアップ時の表示バグ防止）
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();

  // 全体クリア＆スタイル
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setFontFamily('Meiryo')
    .setFontLine('none');

  // 総列数: 6クラス × 4列 = 24列
  var totalCols = NUM_CLASSES * OVERVIEW_COLS_PER_CLASS;
  if (sheet.getMaxColumns() < totalCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), totalCols - sheet.getMaxColumns());
  }

  // 最大行数を確保
  var maxStudents = 0;
  for (var i = 0; i < sizes.length; i++) {
    if (sizes[i] > maxStudents) maxStudents = sizes[i];
  }
  var totalRows = maxStudents + 6;
  if (sheet.getMaxRows() < totalRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  }

  // 列幅設定（各クラスの列）
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    sheet.setColumnWidth(baseCol, 30);      // No
    sheet.setColumnWidth(baseCol + 1, 95);  // 氏名
    sheet.setColumnWidth(baseCol + 2, 25);  // 状態
    sheet.setColumnWidth(baseCol + 3, 100); // 備考
  }

  // 行の高さを詰める
  for (var row = 1; row <= totalRows; row++) {
    sheet.setRowHeight(row, 20);
  }

  // ── 行1: タイトル ──
  sheet.getRange(1, 1, 1, totalCols).merge();
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  sheet.getRange(1, 1)
    .setValue('出欠表　　' + todayStr)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e0e0e0')
    .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ── 行2: 各クラスのヘッダー（○組） ──
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    sheet.getRange(2, baseCol, 1, OVERVIEW_COLS_PER_CLASS).merge();
    sheet.getRange(2, baseCol)
      .setValue((cls + 1) + '組 (' + sizes[cls] + '名)')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#d0d0d0')
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }

  // ── 行3: 列ヘッダー（No, 氏名, 状態, 備考）× 6 ──
  var colHeaders = ['No', '氏名', '', '備考'];
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    sheet.getRange(3, baseCol, 1, OVERVIEW_COLS_PER_CLASS).setValues([colHeaders]);
    sheet.getRange(3, baseCol, 1, OVERVIEW_COLS_PER_CLASS)
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#e8e8e8')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }

  // ── データ行の罫線とスタイルを設定 ──
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    var numStudents = sizes[cls];
    if (numStudents > 0) {
      sheet.getRange(4, baseCol, numStudents, OVERVIEW_COLS_PER_CLASS)
        .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID)
        .setFontSize(8)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      sheet.getRange(4, baseCol, numStudents, 1).setHorizontalAlignment('center');     // No: 中央
      sheet.getRange(4, baseCol + 1, numStudents, 1).setHorizontalAlignment('left');   // 氏名: 左寄せ
      sheet.getRange(4, baseCol + 2, numStudents, 1).setHorizontalAlignment('center'); // 状態: 中央
      sheet.getRange(4, baseCol + 3, numStudents, 1).setHorizontalAlignment('left');   // 備考: 左寄せ
    }
  }

  // ── フッター: 記入凡例 ──
  var footerRow = maxStudents + 5;
  sheet.getRange(footerRow, 1, 1, totalCols).merge();
  sheet.getRange(footerRow, 1)
    .setValue('【凡例】 欠=欠席（理由）　チ=遅刻（時間）　ソ=早退　チソ=遅刻早退　保=保健室　停=出停　転=転校（グレー）　空白=出席')
    .setFontSize(9)
    .setFontWeight('bold')
    .setBackground('#f0f0f0')
    .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // 初回データ投入
  _refreshOverviewData(sheet);
}

/**
 * 一覧シートを更新する（メニューから手動実行）
 * メインシートの最新データを読み取って一覧シートに反映
 */
function refreshOverview() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var overviewSheet = ss.getSheetByName(SHEET_OVERVIEW);
  if (!overviewSheet) {
    SpreadsheetApp.getUi().alert('⚠ 「一覧」シートが見つかりません。\n先に「初期セットアップ」を実行してください。');
    return;
  }

  var totalCols = NUM_CLASSES * OVERVIEW_COLS_PER_CLASS;

  // 更新中表示
  var originalTitle = overviewSheet.getRange(1, 1).getValue();
  overviewSheet.getRange(1, 1).setValue('🔄 一覧の自動更新中...')
    .setFontColor('#ff6600');
  SpreadsheetApp.flush();

  // タイトルの日付を更新（メインシートのA3セルから取得）
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var mainDateStr = mainSheet ? String(mainSheet.getRange('A3').getValue()).replace('📅 ', '') : '';
  var newTitle = mainDateStr ? '出欠表　　' + mainDateStr : originalTitle;

  // データ更新
  _refreshOverviewData(overviewSheet);

  // 更新完了→タイトルを元に戻す
  overviewSheet.getRange(1, 1).setValue(newTitle)
    .setFontColor('#000000');
  SpreadsheetApp.flush();
  ss.toast('🗂 一覧シートを更新しました', 'OVERVIEW', 3);
}

/**
 * 一覧シートのデータ部分を更新する（内部関数）
 */
function _refreshOverviewData(overviewSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  var sizes = getAllClassSizes();

  // 最大人数を求める
  var maxStudents = 0;
  for (var i = 0; i < sizes.length; i++) {
    if (sizes[i] > maxStudents) maxStudents = sizes[i];
  }

  // まず全データエリアをクリア（マージも解除して確実にリセット）
  if (maxStudents > 0) {
    var totalCols = NUM_CLASSES * OVERVIEW_COLS_PER_CLASS;
    var allDataRange = overviewSheet.getRange(4, 1, maxStudents, totalCols);
    allDataRange.breakApart();  // マージセル解除
    allDataRange.clearContent();
    allDataRange.setBackground('#ffffff').setFontColor('#000000').setFontLine('none');

    // 配置を再設定（クリアで崩れるのを防ぐ）
    for (var cls = 0; cls < NUM_CLASSES; cls++) {
      var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
      var numStudents = sizes[cls];
      if (numStudents > 0) {
        overviewSheet.getRange(4, baseCol, numStudents, 1).setHorizontalAlignment('center');     // No
        overviewSheet.getRange(4, baseCol + 1, numStudents, 1).setHorizontalAlignment('left');   // 氏名
        overviewSheet.getRange(4, baseCol + 2, numStudents, 1).setHorizontalAlignment('center'); // 状態
        overviewSheet.getRange(4, baseCol + 3, numStudents, 1).setHorizontalAlignment('left');   // 備考
      }
    }
  }

  // 各クラスのデータを読み取って一覧に書き込む
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var classNum = cls + 1;
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    var startRow = getClassStartRow(classNum);
    var numStudents = sizes[cls];

    if (numStudents <= 0) continue;

    // メインシートから一括読み取り
    var mainData = mainSheet.getRange(startRow, 1, numStudents, NUM_COLS).getValues();

    // 書き込みデータを配列で構築（高速化）
    var writeData = [];
    var colorInfo = [];  // {row, bg, fg, strikethrough}

    for (var s = 0; s < numStudents; s++) {
      var dataRow = mainData[s];
      var no = dataRow[0];         // A: No
      var name = dataRow[1];       // B: 氏名
      var status = String(dataRow[3]).trim(); // D: ステータス
      var reason = dataRow[4];     // E: 理由
      var note = dataRow[5];       // F: 備考
      var time = dataRow[6];       // G: 時刻

      // No が数値でない場合はスキップ（セパレータ行の誤読防止）
      if (no === '' || isNaN(Number(no))) {
        writeData.push(['', '', '', '']);
        continue;
      }

      // ステータス省略記号
      var shortStatus = STATUS_SHORT.hasOwnProperty(status) ? STATUS_SHORT[status] : '';

      // 備考: 省略表示（最大6文字）
      var noteText = '';
      if (status === '欠席' && reason) {
        noteText = _truncate(String(reason), 6);
      } else if (status === '遅刻' && time) {
        noteText = _formatTime(time);
      } else if (status === '早退' && time) {
        noteText = _formatTime(time);
      } else if (status === '遅刻早退' && time) {
        noteText = _formatTime(time);
      } else if (status === '保健室' && time) {
        noteText = _formatTime(time);
      } else if (status === '早退' && note) {
        noteText = _truncate(String(note), 6);
      } else if (note) {
        noteText = _truncate(String(note), 6);
      }

      writeData.push([no, name, shortStatus, noteText]);

      // 色情報を記録
      var overviewRow = 4 + s;
      if (status === '転校') {
        colorInfo.push({row: overviewRow, bg: '#d0d0d0', fg: '#888888', strike: true});
      } else if (status === '欠席') {
        colorInfo.push({row: overviewRow, bg: '#ffe0e0', fg: '#cc0000', strike: false});
      } else if (status === '遅刻') {
        colorInfo.push({row: overviewRow, bg: '#fff8e0', fg: '#996600', strike: false});
      } else if (status === '早退') {
        colorInfo.push({row: overviewRow, bg: '#e0f8ff', fg: '#006699', strike: false});
      } else if (status === '遅刻早退') {
        colorInfo.push({row: overviewRow, bg: '#e8ffe0', fg: '#339900', strike: false});
      } else if (status === '保健室') {
        colorInfo.push({row: overviewRow, bg: '#fff0e0', fg: '#cc6600', strike: false});
      } else if (status === '出停') {
        colorInfo.push({row: overviewRow, bg: '#f0e0ff', fg: '#6600cc', strike: false});
      }
    }

    // 一括書き込み（高速化）
    if (writeData.length > 0) {
      overviewSheet.getRange(4, baseCol, writeData.length, OVERVIEW_COLS_PER_CLASS)
        .setValues(writeData);
    }

    // 色適用
    for (var c = 0; c < colorInfo.length; c++) {
      var info = colorInfo[c];
      var rowRange = overviewSheet.getRange(info.row, baseCol, 1, OVERVIEW_COLS_PER_CLASS);
      rowRange.setBackground(info.bg).setFontColor(info.fg);
      if (info.strike) {
        rowRange.setFontLine('line-through');
      }
    }
  }
}


// ═══════════════════════════════════════════════════════════════════════
//  一覧シート自動更新（5分間隔）
// ═══════════════════════════════════════════════════════════════════════

/**
 * 5分間隔の自動更新トリガーを設定する
 * 既存のautoRefreshOverviewトリガーがあれば削除して再登録
 */
function setupAutoRefresh() {
  // 既存のトリガーを削除（重複防止）
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoRefreshOverview') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 5分間隔のトリガーを新規作成
  ScriptApp.newTrigger('autoRefreshOverview')
    .timeDriven()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('⏱ 一覧の自動更新をONにしました（5分ごと）', 'AUTO REFRESH', 3);
}

/**
 * 自動更新トリガーを解除する
 */
function removeAutoRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoRefreshOverview') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    removed > 0 ? '⏹ 自動更新を停止しました' : '⚠ 自動更新トリガーが見つかりません',
    'AUTO REFRESH', 3
  );
}

/**
 * 自動更新で呼ばれる関数
 * メインシートがアクティブ（先生が入力中）の場合はスキップ
 */
function autoRefreshOverview() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var overviewSheet = ss.getSheetByName(SHEET_OVERVIEW);
  if (!overviewSheet) return;

  // アクティブシートがメインシートの場合は入力中とみなしてスキップ
  try {
    var activeSheet = ss.getActiveSheet();
    if (activeSheet && activeSheet.getName() === SHEET_MAIN) {
      return; // 先生が入力中なのでスキップ
    }
  } catch (e) {
    // タイムベーストリガーではgetActiveSheetが使えない場合がある
    // その場合は更新を実行する
  }

  // 更新中表示
  var originalTitle = overviewSheet.getRange(1, 1).getValue();
  overviewSheet.getRange(1, 1).setValue('🔄 一覧の自動更新中...')
    .setFontColor('#ff6600');
  SpreadsheetApp.flush();

  // タイトルの日付を更新
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  var mainDateStr = mainSheet ? String(mainSheet.getRange('A3').getValue()).replace('📅 ', '') : '';
  var newTitle = mainDateStr ? '出欠表　　' + mainDateStr : originalTitle;

  // データ更新
  _refreshOverviewData(overviewSheet);

  // 更新完了→タイトルを元に戻す
  overviewSheet.getRange(1, 1).setValue(newTitle)
    .setFontColor('#000000');
  SpreadsheetApp.flush();
}


// ═══════════════════════════════════════════════════════════════════════
//  運用フローシート構築
// ═══════════════════════════════════════════════════════════════════════

/**
 * 運用フローシートを構築する（内部関数）
 * 先生の毎日の運用手順をサイバーデザインで表示
 */
function _setupFlowSheet(sheet) {
  // 全体スタイル
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono');

  // 列幅設定
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 250); // メニュー名が見切れないように拡張
  sheet.setColumnWidth(3, 500);

  var r = 1;

  // ── タイトル ──
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('📖 コマンドセンター 運用マニュアル')
    .setFontSize(16).setFontWeight('bold').setFontColor(COLOR.NEON_BLUE)
    .setBackground(COLOR.HEADER_BG).setHorizontalAlignment('center');
  r++;

  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('※ すべての操作はメニュー「🛸 コマンドセンター」から実行できます')
    .setFontSize(9).setFontColor(COLOR.CYBER_GREEN)
    .setBackground(COLOR.HEADER_BG);
  r += 2;

  // ── セクション1: 毎日の運用フロー ──
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('━━ ⭐ 毎日の運用フロー ⭐ ━━')
    .setFontSize(14).setFontWeight('bold').setFontColor(COLOR.CYBER_GREEN)
    .setBackground(COLOR.SEPARATOR_BG).setHorizontalAlignment('center');
  r += 2;

  var flowData = [
    ['STEP', 'タイミング', '操作内容'],
    ['1', '朝（出勤時）', 'メニュー → 「📦+🔄 アーカイブ＆リセット」を実行（前日データを保存し、今日の日付に更新）'],
    ['2', '朝のHR', '各クラスの出欠を入力： D列で「欠席」「遅刻」等を選択。E列はプルダウンから理由を選択、F列に備考を記入'],
    ['3', '日中', '遅刻・早退・保健室等のステータス変更があればその都度更新'],
    ['4', 'ー', '必要に応じて「📋 本日の欠席者リスト」や「🚨 連続欠席チェック」機能を活用'],
    ['5', '帰りのHR後', '確認のみ（明日の朝にアーカイブ＆リセットするので、そのままでOK）'],
  ];

  // ヘッダー
  sheet.getRange(r, 1, 1, 3).setValues([flowData[0]])
    .setBackground(COLOR.DARK_NAVY).setFontColor(COLOR.CYBER_GREEN)
    .setFontWeight('bold').setHorizontalAlignment('center');
  r++;

  // データ行
  for (var i = 1; i < flowData.length; i++) {
    sheet.getRange(r, 1, 1, 3).setValues([flowData[i]]);
    sheet.getRange(r, 1).setHorizontalAlignment('center').setFontColor(COLOR.MAGENTA).setFontWeight('bold');
    sheet.getRange(r, 2).setHorizontalAlignment('center').setFontColor(COLOR.ORANGE);
    if (i % 2 === 0) {
      sheet.getRange(r, 1, 1, 3).setBackground('#050510');
    }
    r++;
  }

  // 罫線
  sheet.getRange(r - flowData.length, 1, flowData.length, 3)
    .setBorder(true, true, true, true, true, true, COLOR.CYBER_GREEN, SpreadsheetApp.BorderStyle.SOLID);

  r += 2;

  // ── セクション2: コマンドセンターの機能一覧 ──
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('━━ 🛸 コマンドセンター機能一覧 ━━')
    .setFontSize(14).setFontWeight('bold').setFontColor(COLOR.CYBER_GREEN)
    .setBackground(COLOR.SEPARATOR_BG).setHorizontalAlignment('center');
  r += 2;

  var commandData = [
    ['メニュー名', '説明'],
    ['🚀 クラスにジャンプ', 'メインシートの1組〜6組にワンクリックで移動'],
    ['📊 ダッシュボード更新', '分析シートにクラス別出席率、ステータス分布等のグラフを描画'],
    ['📋 本日の欠席者リスト', '欠席・遅刻・早退者をまとめてダイアログ表示。職員会議での報告に便利'],
    ['🔍 生徒の出欠履歴', '氏名で検索して過去の出欠記録（日付・理由・備考）を一覧表示'],
    ['📊 月間出欠集計', '指定月の出欠日数・遅刻回数等をクラス別に集計し、新シートに出力'],
    ['📅 学期末集計', '指定した学期・期間で出欠を集計し、新シートに出力'],
    ['🚨 連続欠席チェック', '3日連続欠席（理由が「体調不良」「家事都合」のみ）の生徒に🚨を表示'],
    ['📦 本日分をアーカイブ', '今日の出欠データをログシートに保存。二重実行防止付き'],
    ['🔄 翌日用にリセット', 'ステータス・理由・備考をクリアし、日付を「今日」に更新'],
    ['📦+🔄 アーカイブ＆リセット', '前日ログ保存＋当日リセットをまとめて実行（朝の出勤時に推奨）'],
    ['↩ アーカイブから復元', '誤ってリセットした場合などに、ログの最新データから現状を復元'],
    ['👤 転入生追加', 'クラス番号と氏名を入力→末尾に新番号で追加。設定シートの人数も自動+1'],
    ['🚪 転校処理（選択行）', '生徒の行を選択して実行→グレーアウト＋取り消し線。出欠集計から除外'],
    ['🔄 初期セットアップ', 'シート構造を初期化。クラス人数変更時やプルダウン更新時に再実行'],
    ['⚙ 人数設定シートを開く', '各クラスの人数を変更するシートへジャンプ'],
  ];

  // ヘッダー
  sheet.getRange(r, 2, 1, 2).setValues([commandData[0]])
    .setBackground(COLOR.DARK_NAVY).setFontColor(COLOR.CYBER_GREEN)
    .setFontWeight('bold').setHorizontalAlignment('center');
  r++;

  // データ行
  for (var i = 1; i < commandData.length; i++) {
    sheet.getRange(r, 2, 1, 2).setValues([commandData[i]]);
    
    // レイアウト調整（折り返し、縦中央）
    sheet.getRange(r, 2, 1, 2)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('middle');

    sheet.getRange(r, 2).setFontColor(COLOR.NEON_BLUE).setFontWeight('bold');
    if (i % 2 === 0) {
      sheet.getRange(r, 2, 1, 2).setBackground('#050510');
    }
    r++;
  }

  // 罫線
  sheet.getRange(r - commandData.length, 2, commandData.length, 2)
    .setBorder(true, true, true, true, true, true, COLOR.CYBER_GREEN, SpreadsheetApp.BorderStyle.SOLID);

  r += 2;

  // ── セクション3: ステータス一覧 ──
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('━━ 🎯 ステータス一覧と色 ━━')
    .setFontSize(14).setFontWeight('bold').setFontColor(COLOR.CYBER_GREEN)
    .setBackground(COLOR.SEPARATOR_BG).setHorizontalAlignment('center');
  r += 2;

  var statusData = [
    ['ステータス', '背景色', '説明'],
    ['出席', '黒（デフォルト）', '通常の出席状態'],
    ['欠席', 'ダークレッド', '欠席。E列に理由を記入'],
    ['遅刻', 'ダークイエロー', 'G列に到着時刻を記入'],
    ['早退', 'ダークシアン', 'F列に備考を記入'],
    ['遅刻早退', 'ダークライム', '遅刻して早退した生徒'],
    ['保健室', 'ダークオレンジ', '保健室へ移動した生徒'],
    ['出停', 'ダークマゼンタ', '出席停止（インフル等）'],
    ['転校', 'グレー', '行がグレーアウト＋取り消し線。集計から除外'],
  ];

  // ヘッダー
  sheet.getRange(r, 1, 1, 3).setValues([statusData[0]])
    .setBackground(COLOR.DARK_NAVY).setFontColor(COLOR.CYBER_GREEN)
    .setFontWeight('bold').setHorizontalAlignment('center');
  r++;

  // ステータスごとの行（実際の背景色に彩色）
  var statusColors = [
    [COLOR.BG_BLACK, COLOR.NEON_BLUE],        // 出席
    [COLOR.DARK_RED, '#ff4444'],               // 欠席
    [COLOR.DARK_YELLOW, '#ffff44'],            // 遅刻
    [COLOR.DARK_CYAN, '#44ffff'],              // 早退
    [COLOR.DARK_LIME, '#aaff44'],              // 遅刻早退
    [COLOR.DARK_ORANGE, COLOR.ORANGE],         // 保健室
    [COLOR.DARK_MAGENTA, COLOR.MAGENTA],       // 出停
    [COLOR.GRAY, COLOR.GRAY_TEXT],             // 転校
  ];

  for (var i = 1; i < statusData.length; i++) {
    sheet.getRange(r, 1, 1, 3).setValues([statusData[i]]);
    sheet.getRange(r, 1, 1, 3)
      .setBackground(statusColors[i - 1][0])
      .setFontColor(statusColors[i - 1][1]);
    if (statusData[i][0] === '転校') {
      sheet.getRange(r, 1, 1, 3).setFontLine('line-through');
    }
    r++;
  }

  // 罫線
  sheet.getRange(r - statusData.length, 1, statusData.length, 3)
    .setBorder(true, true, true, true, true, true, COLOR.CYBER_GREEN, SpreadsheetApp.BorderStyle.SOLID);

  r += 2;

  // ── フッター ──
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('🛸 CYBER COMMAND CENTER v4.0 — Powered by Google Apps Script')
    .setFontSize(9).setFontColor('#333333').setHorizontalAlignment('center');
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能I】月間出欠集計
// ═══════════════════════════════════════════════════════════════════════

/**
 * 指定月の出欠集計を新シート「月間集計」に出力する
 */
function monthlyAttendanceSummary() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_LOG);

  if (!logSheet || logSheet.getLastRow() <= 1) {
    ui.alert('⚠ ログにデータがありません。先にアーカイブを実行してください。');
    return;
  }

  // 月の入力
  var monthResponse = ui.prompt(
    '📊 月間出欠集計',
    '集計する年月を入力してください（例: 2026/02）:',
    ui.ButtonSet.OK_CANCEL
  );
  if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

  var inputMonth = monthResponse.getResponseText().trim().replace('-', '/');
  if (!/^\d{4}\/\d{1,2}$/.test(inputMonth)) {
    ui.alert('⚠ 形式が正しくありません。「2026/02」のように入力してください。');
    return;
  }

  // 月の文字列（ログの日付と比較用）
  var parts = inputMonth.split('/');
  var targetYear = parts[0];
  var targetMonth = ('0' + parts[1]).slice(-2);
  var monthPrefix = targetYear + '/' + targetMonth;
  var monthPrefixDash = targetYear + '-' + targetMonth;

  // ログデータを読み取り
  var allLogData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 8).getValues();

  // 月のデータだけ抽出
  var monthData = [];
  for (var i = 0; i < allLogData.length; i++) {
    var dateStr = String(allLogData[i][0]);
    if (dateStr.indexOf(monthPrefix) === 0 || dateStr.indexOf(monthPrefixDash) === 0) {
      monthData.push(allLogData[i]);
    }
  }

  if (monthData.length === 0) {
    ui.alert('⚠ ' + inputMonth + ' のデータがログにありません。');
    return;
  }

  // クラス×生徒ごとに集計
  var summary = {}; // key: "cls-no" → {name, 出席, 欠席, 遅刻, 早退, 保健室, 出停, days}
  var allDates = {};
  for (var i = 0; i < monthData.length; i++) {
    var row = monthData[i];
    var cls = Number(row[1]);
    var no = Number(row[2]);
    var name = String(row[3]);
    var status = String(row[4]).trim();
    var dateStr = String(row[0]);

    allDates[dateStr] = true;
    var key = cls + '-' + no;
    if (!summary[key]) {
      summary[key] = {cls: cls, no: no, name: name, '出席': 0, '欠席': 0, '遅刻': 0, '早退': 0, '遅刻早退': 0, '保健室': 0, '出停': 0};
    }
    if (summary[key].hasOwnProperty(status)) {
      summary[key][status]++;
    } else {
      summary[key]['出席']++;
    }
  }

  var totalDays = Object.keys(allDates).length;

  // 集計シートを作成
  var sheetName = '月間集計_' + targetYear + targetMonth;
  var summarySheet = ss.getSheetByName(sheetName);
  if (summarySheet) {
    ss.deleteSheet(summarySheet);
  }
  summarySheet = ss.insertSheet(sheetName);

  // ヘッダー
  var headers = ['クラス', 'No', '氏名', '出席', '欠席', '遅刻', '早退', '遅刻早退', '保健室', '出停', '授業日数'];
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(COLOR.DARK_NAVY)
    .setFontColor(COLOR.CYBER_GREEN)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // データ行
  var keys = Object.keys(summary).sort(function(a, b) {
    var aParts = a.split('-');
    var bParts = b.split('-');
    if (Number(aParts[0]) !== Number(bParts[0])) return Number(aParts[0]) - Number(bParts[0]);
    return Number(aParts[1]) - Number(bParts[1]);
  });

  var writeData = [];
  for (var k = 0; k < keys.length; k++) {
    var s = summary[keys[k]];
    writeData.push([s.cls + '組', s.no, s.name, s['出席'], s['欠席'], s['遅刻'], s['早退'], s['遅刻早退'], s['保健室'], s['出停'], totalDays]);
  }

  if (writeData.length > 0) {
    summarySheet.getRange(2, 1, writeData.length, headers.length).setValues(writeData);
  }

  // スタイル
  summarySheet.getRange(1, 1, writeData.length + 1, headers.length)
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono')
    .setBorder(true, true, true, true, true, true, '#333333', SpreadsheetApp.BorderStyle.SOLID);

  // 欠席列を赤く
  for (var r = 0; r < writeData.length; r++) {
    if (writeData[r][4] > 0) { // 欠席 > 0
      summarySheet.getRange(r + 2, 5).setFontColor('#ff4444').setFontWeight('bold');
    }
    if (writeData[r][5] > 0) { // 遅刻 > 0
      summarySheet.getRange(r + 2, 6).setFontColor('#ffaa00').setFontWeight('bold');
    }
  }

  // 列幅
  summarySheet.setColumnWidth(1, 60);
  summarySheet.setColumnWidth(2, 40);
  summarySheet.setColumnWidth(3, 100);
  for (var c = 4; c <= 11; c++) {
    summarySheet.setColumnWidth(c, 65);
  }

  // タイトル行は1行目のヘッダーのまま
  summarySheet.setFrozenRows(1);

  ss.setActiveSheet(summarySheet);
  ss.toast('📊 ' + inputMonth + ' の集計完了（' + totalDays + '日間、' + keys.length + '名）', '月間集計', 5);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能J】本日の欠席者リスト
// ═══════════════════════════════════════════════════════════════════════

/**
 * メインシートから本日の欠席・遅刻・早退・保健室・出停の生徒をダイアログ表示
 */
function showTodayAbsentList() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  var sizes = getAllClassSizes();
  var absentList = [];

  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];

    for (var s = 0; s < numStudents; s++) {
      var rowNum = startRow + s;
      var rowData = mainSheet.getRange(rowNum, 1, 1, NUM_COLS).getValues()[0];
      var no = rowData[0];
      var name = String(rowData[1]).trim();
      var status = String(rowData[3]).trim();
      var reason = String(rowData[4] || '');
      var note = String(rowData[5] || '');
      var time = rowData[6];

      if (!name || status === '出席' || status === '' || status === STATUS_TRANSFERRED) continue;

      var detail = cls + '組 No.' + no + ' ' + name + ' 【' + status + '】';
      if (status === '欠席' && reason) detail += ' 理由: ' + reason;
      if (status === '遅刻' && time) detail += ' ' + _formatTime(time);
      if (status === '早退' && time) detail += ' ' + _formatTime(time);
      if (note) detail += ' 備考: ' + note;

      absentList.push(detail);
    }
  }

  // 日付取得
  var dateStr = mainSheet.getRange('A3').getValue();

  if (absentList.length === 0) {
    ui.alert('✅ 本日の欠席者はいません\n\n' + dateStr);
    return;
  }

  var msg = '📋 本日の欠席者・遅刻者リスト\n' + dateStr + '\n\n';
  msg += absentList.join('\n');
  msg += '\n\n合計: ' + absentList.length + '名';

  ui.alert('📋 欠席者リスト（' + absentList.length + '名）', msg, ui.ButtonSet.OK);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能J-2】日付指定の欠席者リスト（ログから取得）
// ═══════════════════════════════════════════════════════════════════════

/**
 * ログシートの日付一覧をカレンダーHTMLダイアログで表示し、クリックで欠席者リストを取得
 */
function showDateAbsentList() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_LOG);

  if (!logSheet) {
    ui.alert('⚠ ログシートが見つかりません。');
    return;
  }

  var lastRow = logSheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('⚠ ログにデータがありません。\n先にアーカイブを実行してください。');
    return;
  }

  // ログから日付の一覧を取得（重複排除）
  var dateCol = logSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var dateSet = {};
  for (var i = 0; i < dateCol.length; i++) {
    var raw = dateCol[i][0];
    var d;
    if (raw instanceof Date) {
      d = Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy/MM/dd');
    } else {
      d = String(raw).replace(/-/g, '/').trim();
    }
    if (d) dateSet[d] = true;
  }
  var dates = Object.keys(dateSet);

  if (dates.length === 0) {
    ui.alert('⚠ ログに日付データがありません。');
    return;
  }

  // 日付一覧をJSON文字列で渡す
  var datesJson = JSON.stringify(dates);

  var html = ''
    + '<!DOCTYPE html><html><head><style>'
    + '* { box-sizing: border-box; margin: 0; padding: 0; }'
    + 'body { font-family: "Meiryo", sans-serif; background: #1a1a2e; color: #e0e0e0; padding: 16px; }'
    + '.cal-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }'
    + '.cal-header h3 { color: #00e5ff; font-size: 16px; }'
    + '.nav-btn { background: none; border: 1px solid #444; color: #00e5ff; width: 32px; height: 32px;'
    + '  border-radius: 50%; cursor: pointer; font-size: 16px; transition: all 0.2s; }'
    + '.nav-btn:hover { background: #00e5ff; color: #0d1b2a; }'
    + '.weekday-row { display: grid; grid-template-columns: repeat(7, 1fr); text-align: center;'
    + '  font-size: 12px; font-weight: bold; color: #888; padding: 4px 0; border-bottom: 1px solid #333; margin-bottom: 4px; }'
    + '.weekday-row .sun { color: #ff5252; }'
    + '.weekday-row .sat { color: #448aff; }'
    + '.cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; }'
    + '.day-cell { text-align: center; padding: 8px 0; font-size: 13px; border-radius: 50%;'
    + '  width: 36px; height: 36px; line-height: 36px; padding: 0; margin: 2px auto;'
    + '  transition: all 0.2s; position: relative; }'
    + '.day-cell.empty { }'
    + '.day-cell.no-data { color: #555; }'
    + '.day-cell.has-data { color: #00e5ff; cursor: pointer; font-weight: bold; background: rgba(0,229,255,0.1); }'
    + '.day-cell.has-data:hover { background: #00e5ff; color: #0d1b2a; }'
    + '.day-cell.selected { background: #00e5ff !important; color: #0d1b2a !important; }'
    + '.day-cell.today { border: 2px solid #ffab00; }'
    + '.day-cell.sun { color: #ff5252; }'
    + '.day-cell.sat { color: #448aff; }'
    + '.day-cell.has-data.sun { color: #ff8a80; background: rgba(255,82,82,0.1); }'
    + '.day-cell.has-data.sun:hover, .day-cell.selected.sun { background: #ff5252 !important; color: #fff !important; }'
    + '.day-cell.has-data.sat { color: #82b1ff; background: rgba(68,138,255,0.1); }'
    + '.day-cell.has-data.sat:hover, .day-cell.selected.sat { background: #448aff !important; color: #fff !important; }'
    + '#result { margin-top: 14px; white-space: pre-wrap; font-size: 12px; line-height: 1.7;'
    + '  background: #0d1b2a; padding: 12px; border-radius: 8px; border: 1px solid #333; display: none;'
    + '  max-height: 200px; overflow-y: auto; }'
    + '#loading { display: none; color: #ffab00; font-size: 13px; margin-top: 12px; }'
    + '.summary { color: #00e5ff; font-weight: bold; margin-top: 8px; }'
    + '</style></head><body>'
    + '<div class="cal-header">'
    + '  <button class="nav-btn" onclick="changeMonth(-1)">◀</button>'
    + '  <h3 id="monthLabel"></h3>'
    + '  <button class="nav-btn" onclick="changeMonth(1)">▶</button>'
    + '</div>'
    + '<div class="weekday-row">'
    + '  <div class="sun">日</div><div>月</div><div>火</div><div>水</div><div>木</div><div>金</div><div class="sat">土</div>'
    + '</div>'
    + '<div class="cal-grid" id="calGrid"></div>'
    + '<div id="loading">⏳ 読み込み中...</div>'
    + '<div id="result"></div>'
    + '<script>'
    + 'var logDates = ' + datesJson + ';'
    + 'var currentYear, currentMonth;'
    + ''
    + 'function init() {'
    + '  var now = new Date();'
    + '  currentYear = now.getFullYear();'
    + '  currentMonth = now.getMonth();'
    + '  renderCalendar();'
    + '}'
    + ''
    + 'function changeMonth(delta) {'
    + '  currentMonth += delta;'
    + '  if (currentMonth < 0) { currentMonth = 11; currentYear--; }'
    + '  if (currentMonth > 11) { currentMonth = 0; currentYear++; }'
    + '  renderCalendar();'
    + '}'
    + ''
    + 'function renderCalendar() {'
    + '  var label = currentYear + "年" + (currentMonth + 1) + "月";'
    + '  document.getElementById("monthLabel").textContent = label;'
    + '  var grid = document.getElementById("calGrid");'
    + '  grid.innerHTML = "";'
    + '  var firstDay = new Date(currentYear, currentMonth, 1).getDay();'
    + '  var daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();'
    + '  var today = new Date();'
    + '  var todayStr = today.getFullYear() + "/" + String(today.getMonth()+1).padStart(2,"0") + "/" + String(today.getDate()).padStart(2,"0");'
    + '  for (var i = 0; i < firstDay; i++) {'
    + '    var empty = document.createElement("div");'
    + '    empty.className = "day-cell empty";'
    + '    grid.appendChild(empty);'
    + '  }'
    + '  for (var d = 1; d <= daysInMonth; d++) {'
    + '    var cell = document.createElement("div");'
    + '    cell.textContent = d;'
    + '    var mm = String(currentMonth + 1).padStart(2, "0");'
    + '    var dd = String(d).padStart(2, "0");'
    + '    var dateStr = currentYear + "/" + mm + "/" + dd;'
    + '    var dow = new Date(currentYear, currentMonth, d).getDay();'
    + '    var cls = "day-cell";'
    + '    if (dow === 0) cls += " sun";'
    + '    if (dow === 6) cls += " sat";'
    + '    if (dateStr === todayStr) cls += " today";'
    + '    var hasData = logDates.indexOf(dateStr) >= 0;'
    + '    if (hasData) {'
    + '      cls += " has-data";'
    + '      cell.onclick = (function(ds){ return function(){ selectDate(ds, this); }; })(dateStr);'
    + '    } else {'
    + '      cls += " no-data";'
    + '    }'
    + '    cell.className = cls;'
    + '    grid.appendChild(cell);'
    + '  }'
    + '}'
    + ''
    + 'function selectDate(date, el) {'
    + '  document.getElementById("loading").style.display = "block";'
    + '  document.getElementById("result").style.display = "none";'
    + '  document.querySelectorAll(".day-cell.selected").forEach(function(c){ c.classList.remove("selected"); });'
    + '  el.classList.add("selected");'
    + '  google.script.run.withSuccessHandler(showResult).withFailureHandler(showError)._getAbsentListForDate(date);'
    + '}'
    + ''
    + 'function showResult(data) {'
    + '  document.getElementById("loading").style.display = "none";'
    + '  var el = document.getElementById("result");'
    + '  el.style.display = "block";'
    + '  el.innerHTML = data;'
    + '}'
    + 'function showError(e) {'
    + '  document.getElementById("loading").style.display = "none";'
    + '  var el = document.getElementById("result");'
    + '  el.style.display = "block";'
    + '  el.textContent = "⚠ エラー: " + e.message;'
    + '}'
    + 'init();'
    + '</script></body></html>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(380)
    .setHeight(520);
  ui.showModalDialog(htmlOutput, '📋 日付指定 欠席者リスト');
}

/**
 * 指定日付の欠席者データをHTML文字列で返す（HTMLダイアログから呼び出し用）
 */
function _getAbsentListForDate(targetDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) return '<span style="color:#ff5252">⚠ ログシートが見つかりません。</span>';

  var lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return '<span style="color:#ff5252">⚠ ログにデータがありません。</span>';

  var allLogData = logSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var normalizedTarget = targetDate.replace(/-/g, '/');

  var absentList = [];
  for (var i = 0; i < allLogData.length; i++) {
    var rawDate = allLogData[i][0];
    var logDate;
    if (rawDate instanceof Date) {
      logDate = Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    } else {
      logDate = String(rawDate).replace(/-/g, '/');
    }
    if (logDate !== normalizedTarget) continue;

    var cls    = allLogData[i][1];
    var no     = allLogData[i][2];
    var name   = String(allLogData[i][3]).trim();
    var status = String(allLogData[i][4]).trim();
    var reason = String(allLogData[i][5] || '');
    var note   = String(allLogData[i][6] || '');
    var time   = allLogData[i][7];

    if (status === '出席' || status === '') continue;

    var detail = cls + '組 No.' + no + ' ' + name + ' 【' + status + '】';
    if (status === '欠席' && reason) detail += '  理由: ' + reason;
    if (status === '遅刻' && time) detail += '  ' + _formatTime(time);
    if (status === '早退' && time) detail += '  ' + _formatTime(time);
    if (status === '保健室' && time) detail += '  ' + _formatTime(time);
    if (note) detail += '  備考: ' + note;

    absentList.push(detail);
  }

  if (absentList.length === 0) {
    return '<span style="color:#69f0ae">✅ ' + normalizedTarget + ' の欠席者はいません（全員出席）</span>';
  }

  var html = '<div style="color:#ffab00; font-weight:bold; margin-bottom:8px;">📋 ' + normalizedTarget + ' の欠席者・遅刻者</div>';
  for (var j = 0; j < absentList.length; j++) {
    html += '<div>' + absentList[j] + '</div>';
  }
  html += '<div class="summary">合計: ' + absentList.length + '名</div>';
  return html;
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能K】特定生徒の出欠履歴
// ═══════════════════════════════════════════════════════════════════════

/**
 * 生徒名を入力して、過去の出欠記録をダイアログで表示
 */
function showStudentHistory() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_LOG);

  if (!logSheet || logSheet.getLastRow() <= 1) {
    ui.alert('⚠ ログにデータがありません。');
    return;
  }

  // 生徒名の入力
  var nameResponse = ui.prompt(
    '🔍 生徒の出欠履歴',
    '検索する生徒の氏名（部分一致可）:',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;

  var searchName = nameResponse.getResponseText().trim();
  if (!searchName) {
    ui.alert('⚠ 氏名が入力されていません。');
    return;
  }

  // ログから検索
  var allLogData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 8).getValues();
  var results = [];
  var countByStatus = {};

  for (var i = 0; i < allLogData.length; i++) {
    var row = allLogData[i];
    var name = String(row[3]);

    if (name.indexOf(searchName) === -1) continue;

    var dateStr = String(row[0]);
    var cls = row[1];
    var no = row[2];
    var status = String(row[4]).trim();
    var reason = String(row[5] || '');
    var note = String(row[6] || '');

    // ステータス集計
    if (!countByStatus[status]) countByStatus[status] = 0;
    countByStatus[status]++;

    // 出席以外のみ詳細に表示
    if (status !== '出席') {
      var line = dateStr + ' ' + status;
      if (reason) line += ' (' + reason + ')';
      if (note) line += ' ' + note;
      results.push(line);
    }
  }

  if (Object.keys(countByStatus).length === 0) {
    ui.alert('⚠ 「' + searchName + '」に一致する生徒が見つかりません。');
    return;
  }

  // 集計サマリー
  var summaryParts = [];
  var statusOrder = ['出席', '欠席', '遅刻', '早退', '遅刻早退', '保健室', '出停'];
  for (var s = 0; s < statusOrder.length; s++) {
    if (countByStatus[statusOrder[s]]) {
      summaryParts.push(statusOrder[s] + ': ' + countByStatus[statusOrder[s]] + '回');
    }
  }

  var msg = '🔍 「' + searchName + '」の出欠履歴\n\n';
  msg += '【集計】\n' + summaryParts.join('  /  ') + '\n\n';

  if (results.length > 0) {
    msg += '【出席以外の記録】\n';
    // 最新50件に制限
    var showResults = results.slice(-50);
    msg += showResults.join('\n');
    if (results.length > 50) {
      msg += '\n\n（最新50件を表示。全' + results.length + '件）';
    }
  } else {
    msg += '※ 全日出席です！';
  }

  ui.alert('🔍 出欠履歴: ' + searchName, msg, ui.ButtonSet.OK);
}


// ═══════════════════════════════════════════════════════════════════════
//  【機能L】学期末集計
// ═══════════════════════════════════════════════════════════════════════

/**
 * 学期ごとの期間を指定して出欠を集計する
 */
function semesterSummary() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(SHEET_LOG);

  if (!logSheet || logSheet.getLastRow() <= 1) {
    ui.alert('⚠ ログにデータがありません。');
    return;
  }

  // 学期選択
  var semesterResponse = ui.prompt(
    '📅 学期末集計',
    '学期を選択してください:\n1 = 1学期\n2 = 2学期\n3 = 3学期\n\n番号を入力:',
    ui.ButtonSet.OK_CANCEL
  );
  if (semesterResponse.getSelectedButton() !== ui.Button.OK) return;

  var semesterNum = parseInt(semesterResponse.getResponseText().trim(), 10);
  if (isNaN(semesterNum) || semesterNum < 1 || semesterNum > 3) {
    ui.alert('⚠ 1〜3の番号を入力してください。');
    return;
  }

  // 期間入力（開始日）
  var startResponse = ui.prompt(
    '📅 ' + semesterNum + '学期 開始日',
    '開始日を入力してください（例: 2026/04/07）:',
    ui.ButtonSet.OK_CANCEL
  );
  if (startResponse.getSelectedButton() !== ui.Button.OK) return;

  var startDateStr = startResponse.getResponseText().trim().replace(/-/g, '/');

  // 期間入力（終了日）
  var endResponse = ui.prompt(
    '📅 ' + semesterNum + '学期 終了日',
    '終了日を入力してください（例: 2026/07/20）:',
    ui.ButtonSet.OK_CANCEL
  );
  if (endResponse.getSelectedButton() !== ui.Button.OK) return;

  var endDateStr = endResponse.getResponseText().trim().replace(/-/g, '/');

  // 日付をDateオブジェクトに変換
  var startDate = new Date(startDateStr);
  var endDate = new Date(endDateStr);

  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    ui.alert('⚠ 日付の形式が正しくありません。');
    return;
  }

  if (startDate > endDate) {
    ui.alert('⚠ 開始日が終了日より後になっています。');
    return;
  }

  // ログデータを読み取り
  var allLogData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 8).getValues();

  // 期間内のデータだけ抽出
  var periodData = [];
  var allDates = {};
  for (var i = 0; i < allLogData.length; i++) {
    var logDate = new Date(String(allLogData[i][0]).replace(/-/g, '/'));
    if (logDate >= startDate && logDate <= endDate) {
      periodData.push(allLogData[i]);
      allDates[String(allLogData[i][0])] = true;
    }
  }

  if (periodData.length === 0) {
    ui.alert('⚠ 指定期間のデータがありません。');
    return;
  }

  var totalDays = Object.keys(allDates).length;

  // 集計
  var summary = {};
  for (var i = 0; i < periodData.length; i++) {
    var row = periodData[i];
    var cls = Number(row[1]);
    var no = Number(row[2]);
    var name = String(row[3]);
    var status = String(row[4]).trim();

    var key = cls + '-' + no;
    if (!summary[key]) {
      summary[key] = {cls: cls, no: no, name: name, '出席': 0, '欠席': 0, '遅刻': 0, '早退': 0, '遅刻早退': 0, '保健室': 0, '出停': 0};
    }
    if (summary[key].hasOwnProperty(status)) {
      summary[key][status]++;
    } else {
      summary[key]['出席']++;
    }
  }

  // 集計シートを作成
  var semesterNames = {1: '1学期', 2: '2学期', 3: '3学期'};
  var sheetName = semesterNames[semesterNum] + '集計';
  var summarySheet = ss.getSheetByName(sheetName);
  if (summarySheet) {
    ss.deleteSheet(summarySheet);
  }
  summarySheet = ss.insertSheet(sheetName);

  // タイトル
  var headers = ['クラス', 'No', '氏名', '出席', '欠席', '遅刻', '早退', '遅刻早退', '保健室', '出停', '授業日数'];
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(COLOR.DARK_NAVY)
    .setFontColor(COLOR.CYBER_GREEN)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // データ
  var keys = Object.keys(summary).sort(function(a, b) {
    var aParts = a.split('-');
    var bParts = b.split('-');
    if (Number(aParts[0]) !== Number(bParts[0])) return Number(aParts[0]) - Number(bParts[0]);
    return Number(aParts[1]) - Number(bParts[1]);
  });

  var writeData = [];
  for (var k = 0; k < keys.length; k++) {
    var s = summary[keys[k]];
    writeData.push([s.cls + '組', s.no, s.name, s['出席'], s['欠席'], s['遅刻'], s['早退'], s['遅刻早退'], s['保健室'], s['出停'], totalDays]);
  }

  if (writeData.length > 0) {
    summarySheet.getRange(2, 1, writeData.length, headers.length).setValues(writeData);
  }

  // スタイル
  summarySheet.getRange(1, 1, writeData.length + 1, headers.length)
    .setBackground(COLOR.BG_BLACK)
    .setFontColor(COLOR.NEON_BLUE)
    .setFontFamily('Roboto Mono')
    .setBorder(true, true, true, true, true, true, '#333333', SpreadsheetApp.BorderStyle.SOLID);

  for (var r = 0; r < writeData.length; r++) {
    if (writeData[r][4] > 0) {
      summarySheet.getRange(r + 2, 5).setFontColor('#ff4444').setFontWeight('bold');
    }
    if (writeData[r][5] > 0) {
      summarySheet.getRange(r + 2, 6).setFontColor('#ffaa00').setFontWeight('bold');
    }
  }

  // 列幅
  summarySheet.setColumnWidth(1, 60);
  summarySheet.setColumnWidth(2, 40);
  summarySheet.setColumnWidth(3, 100);
  for (var c = 4; c <= 11; c++) {
    summarySheet.setColumnWidth(c, 65);
  }

  summarySheet.setFrozenRows(1);
  ss.setActiveSheet(summarySheet);
  ss.toast('📅 ' + semesterNames[semesterNum] + '（' + startDateStr + '〜' + endDateStr + '）集計完了！ ' + totalDays + '日間', '学期末集計', 5);
}
