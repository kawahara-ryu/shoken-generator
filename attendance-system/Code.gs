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
  '保健室': '保',
  '出停': '停',
  '転校': '転'
};

/** 全ステータスリスト（プルダウン用、転校含む） */
var ALL_STATUSES = ['出席', '欠席', '遅刻', '早退', '保健室', '出停', '転校'];

/** 集計対象ステータス（転校を除く） */
var ACTIVE_STATUSES = ['出席', '欠席', '遅刻', '早退', '保健室', '出停'];

// ─── ユーティリティ ──────────────────────────────────────────────────────

/**
 * 設定シートから各クラスの生徒数を読み取る
 * @param {number} classNum - クラス番号 (1〜6)
 * @return {number} そのクラスの生徒数
 */
function getStudentsPerClass(classNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) return DEFAULT_STUDENTS;

  // 設定シート: 行1=タイトル, 行2=ヘッダー, 行3〜8=クラス1〜6のデータ
  var value = configSheet.getRange(classNum + 2, 2).getValue();
  if (value && !isNaN(value) && value > 0) {
    return Number(value);
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
  var row = 5;
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
 * 指定クラスの区切り行を返す（classNum >= 2）
 */
function getSeparatorRow(classNum) {
  return getClassHeaderRow(classNum) - 1;
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

    if (status === '欠席') {
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
      if (status === STATUS_TRANSFERRED) continue; // 転校生は除外
      if (status === '欠席') {
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
  var statusCounts = { '出席': 0, '欠席': 0, '遅刻': 0, '早退': 0, '保健室': 0, '出停': 0 };
  var transferCounts = []; // クラスごとの転校者数

  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var startRow = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];
    var classCount = { '出席': 0, '欠席': 0, '遅刻': 0, '早退': 0, '保健室': 0, '出停': 0, '合計': 0, '転校': 0 };

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
                       classCount['遅刻'], classCount['早退'], classCount['保健室'], classCount['出停'],
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
  var tableHeaders = ['クラス', '出席率(%)', '出席', '欠席', '遅刻', '早退', '保健室', '出停', '在籍数', '転校'];
  for (var h = 0; h < tableHeaders.length; h++) {
    analysisSheet.getRange(4, h + 1).setValue(tableHeaders[h]).setFontWeight('bold');
  }
  analysisSheet.getRange(4, 1, 1, 10)
    .setBackground(COLOR.DARK_NAVY)
    .setFontColor(COLOR.CYBER_GREEN);

  for (var r = 0; r < summaryData.length; r++) {
    analysisSheet.getRange(5 + r, 1, 1, 10).setValues([summaryData[r]]);
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
    .addRange(analysisSheet.getRange('A13:B19'))
    .setPosition(4, 19, 0, 0)
    .setOption('title', '🔮 ステータス分布')
    .setOption('titleTextStyle', { color: COLOR.MAGENTA, fontSize: 14, bold: true })
    .setOption('backgroundColor', { fill: COLOR.BG_BLACK })
    .setOption('legend', { textStyle: { color: COLOR.NEON_BLUE }, position: 'right' })
    .setOption('pieSliceBorderColor', COLOR.BG_BLACK)
    .setOption('colors', [COLOR.CYBER_GREEN, '#ff0040', '#ffff00', COLOR.NEON_BLUE, COLOR.ORANGE, COLOR.MAGENTA])
    .setOption('chartArea', { backgroundColor: { fill: COLOR.BG_BLACK } })
    .setOption('width', 500).setOption('height', 350)
    .build();
  analysisSheet.insertChart(pieChart);

  // グラフ3: クラス別詳細（積み上げ棒グラフ）
  var stackedChart = analysisSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(analysisSheet.getRange('A4:H10'))
    .setPosition(22, 12, 0, 0)
    .setOption('title', '📊 クラス別ステータス詳細')
    .setOption('titleTextStyle', { color: COLOR.CYBER_GREEN, fontSize: 14, bold: true })
    .setOption('backgroundColor', { fill: COLOR.BG_BLACK })
    .setOption('isStacked', true)
    .setOption('legend', { textStyle: { color: COLOR.NEON_BLUE }, position: 'top' })
    .setOption('hAxis', { textStyle: { color: COLOR.NEON_BLUE }, gridlines: { color: '#1a1a2e' } })
    .setOption('vAxis', { textStyle: { color: COLOR.NEON_BLUE } })
    .setOption('colors', [COLOR.NEON_BLUE, COLOR.CYBER_GREEN, '#ff0040', '#ffff00', COLOR.ORANGE, COLOR.MAGENTA])
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

  var sizes = getAllClassSizes();
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
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  var confirm = ui.alert(
    '🔄 デイリーリセット確認',
    '全生徒のステータスを「出席」に戻し、理由・備考欄をクリアします。\n日付を「' + todayStr + '」に更新します。\n\n⚠ アーカイブは済んでいますか？\n（未保存のデータは失われます）\n\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

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

  mainSheet.getRange('A3').setValue('📅 ' + todayStr);

  SpreadsheetApp.flush();
  ss.toast('🔄 デイリーリセット完了（' + todayStr + '）', 'SYSTEM RESET', 3);
}

/**
 * アーカイブ + リセットを一括実行
 */
function dailyArchiveAndReset() {
  dailyArchive();
  dailyReset();
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
  var row = 5;
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

  // 設定シートの構築
  _setupConfigSheet(configSheet);

  // クラス人数を取得
  var sizes = getAllClassSizes();

  // メインシートの行数を確保
  var totalRows = 4;
  for (var i = 0; i < NUM_CLASSES; i++) {
    totalRows += sizes[i];
    if (i < NUM_CLASSES - 1) totalRows += 2;
  }
  totalRows += 10;
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

  // 列幅設定
  mainSheet.setColumnWidth(1, 45);
  mainSheet.setColumnWidth(2, 130);
  mainSheet.setColumnWidth(3, 45);
  mainSheet.setColumnWidth(4, 110);
  mainSheet.setColumnWidth(5, 160);
  mainSheet.setColumnWidth(6, 160);
  mainSheet.setColumnWidth(7, 90);

  // 行1: メインタイトル
  mainSheet.getRange('A1').setValue('🛸 CYBER COMMAND CENTER — 出欠管理システム v4.0');
  mainSheet.getRange(1, 1, 1, NUM_COLS)
    .merge()
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor(COLOR.NEON_BLUE)
    .setBackground(COLOR.HEADER_BG)
    .setHorizontalAlignment('center');

  // 行2: ナビゲーションボタン配置エリア
  mainSheet.getRange('A2').setValue('◄ ここに図形ボタンを配置 ► [1組] [2組] [3組] [4組] [5組] [6組] [分析]');
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

  // 各クラスの枠を生成
  for (var cls = 1; cls <= NUM_CLASSES; cls++) {
    var headerRow = getClassHeaderRow(cls);
    var startRow  = getClassStartRow(cls);
    var numStudents = sizes[cls - 1];

    // 区切りライン（2組以降）
    if (cls >= 2) {
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
    }

    // 列ヘッダー
    var headers = ['No', '氏名', '🚨', 'ステータス', '理由', '備考', '時刻'];
    mainSheet.getRange(headerRow, 1, 1, NUM_COLS).setValues([headers]);
    if (cls === 1) {
      mainSheet.getRange('A3').setValue('📅 ' + todayStr + '　　　　1組 (' + numStudents + '名)');
    }
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
  analysisSheet.getRange('A2').setValue('📊 「ダッシュボード更新」ボタンを押してグラフを生成')
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
  var jumpMenu = ui.createMenu('� クラスにジャンプ')
    .addItem('1組へ', 'jumpToClass1')
    .addItem('2組へ', 'jumpToClass2')
    .addItem('3組へ', 'jumpToClass3')
    .addItem('4組へ', 'jumpToClass4')
    .addItem('5組へ', 'jumpToClass5')
    .addItem('6組へ', 'jumpToClass6');

  ui.createMenu('�🛸 コマンドセンター')
    .addSubMenu(jumpMenu)
    .addItem('📊 分析シートへ', 'jumpToAnalysis')
    .addItem('📖 運用フローへ', 'jumpToFlow')
    .addItem('🗂 一覧シートへ', 'jumpToOverview')
    .addSeparator()
    .addItem('🔄 一覧を更新', 'refreshOverview')
    .addItem('📦 本日分をアーカイブ', 'dailyArchive')
    .addItem('🔄 翌日用にリセット', 'dailyReset')
    .addItem('📦+🔄 アーカイブ＆リセット', 'dailyArchiveAndReset')
    .addSeparator()
    .addItem('🚨 連続欠席チェック', 'checkConsecutiveAbsence')
    .addItem('📊 ダッシュボード更新', 'refreshHologramCharts')
    .addSeparator()
    .addItem('👤 転入生追加', 'addTransferStudent')
    .addItem('🚪 転校処理（選択行）', 'markAsTransferred')
    .addItem('↩ 転校取り消し（選択行）', 'undoTransfer')
    .addSeparator()
    .addItem('🔄 初期セットアップ', 'setupSpreadsheet')
    .addItem('⚙ 人数設定シートを開く', 'jumpToConfig')
    .addToUi();

  // 日付を自動更新（スプレッドシートを開いた時点の日付に）
  _updateDateDisplay();
}

/**
 * メインシートの日付表示を今日の日付に更新する（内部関数）
 */
function _updateDateDisplay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName(SHEET_MAIN);
  if (!mainSheet) return;

  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  var currentA3 = String(mainSheet.getRange('A3').getValue());
  // 日付部分だけ更新（クラス人数情報は保持）
  if (currentA3.indexOf('　　') !== -1) {
    var suffix = currentA3.substring(currentA3.indexOf('　　'));
    mainSheet.getRange('A3').setValue('📅 ' + todayStr + suffix);
  } else {
    mainSheet.getRange('A3').setValue('📅 ' + todayStr);
  }
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
  // 全体スタイル
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setFontFamily('Meiryo');

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
  var totalRows = maxStudents + 4; // タイトル(1) + 記入例(1) + ヘッダー(1) + データ + フッター
  if (sheet.getMaxRows() < totalRows + 5) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRows + 5 - sheet.getMaxRows());
  }

  // 列幅設定（各クラスの列）
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    sheet.setColumnWidth(baseCol, 35);      // No
    sheet.setColumnWidth(baseCol + 1, 85);  // 氏名
    sheet.setColumnWidth(baseCol + 2, 30);  // 状態
    sheet.setColumnWidth(baseCol + 3, 100); // 備考
  }

  // 行の高さを詰める
  for (var row = 1; row <= totalRows + 3; row++) {
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

  // ── データ行の罫線を設定 ──
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    var numStudents = sizes[cls];
    if (numStudents > 0) {
      // データ範囲に罫線
      sheet.getRange(4, baseCol, numStudents, OVERVIEW_COLS_PER_CLASS)
        .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID)
        .setFontSize(9);

      // No列とステータス列を中央揃え
      sheet.getRange(4, baseCol, numStudents, 1).setHorizontalAlignment('center');
      sheet.getRange(4, baseCol + 2, numStudents, 1).setHorizontalAlignment('center');
    }
  }

  // ── フッター: 記入凡例 ──
  var footerRow = maxStudents + 5;
  sheet.getRange(footerRow, 1, 1, totalCols).merge();
  sheet.getRange(footerRow, 1)
    .setValue('【凡例】 欠=欠席（理由）　チ=遅刻（時間）　ソ=早退　保=保健室　停=出停　転=転校（グレー）　空白=出席')
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

  // タイトルの日付を更新
  var totalCols = NUM_CLASSES * OVERVIEW_COLS_PER_CLASS;
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd (E)');
  overviewSheet.getRange(1, 1).setValue('出欠表　　' + todayStr);

  // データ更新
  _refreshOverviewData(overviewSheet);

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

  // 各クラスのデータを読み取って一覧に書き込む
  for (var cls = 0; cls < NUM_CLASSES; cls++) {
    var classNum = cls + 1;
    var baseCol = cls * OVERVIEW_COLS_PER_CLASS + 1;
    var startRow = getClassStartRow(classNum);
    var numStudents = sizes[cls];

    // まずデータ範囲をクリア（最大人数分）
    if (maxStudents > 0) {
      overviewSheet.getRange(4, baseCol, maxStudents, OVERVIEW_COLS_PER_CLASS).clearContent();
      overviewSheet.getRange(4, baseCol, maxStudents, OVERVIEW_COLS_PER_CLASS)
        .setBackground('#ffffff')
        .setFontColor('#000000')
        .setFontLine('none');
    }

    // メインシートからデータを読み取る
    if (numStudents > 0) {
      var mainData = mainSheet.getRange(startRow, 1, numStudents, NUM_COLS).getValues();

      for (var s = 0; s < numStudents; s++) {
        var dataRow = mainData[s];
        var no = dataRow[0];         // A: No
        var name = dataRow[1];       // B: 氏名
        var status = String(dataRow[3]).trim(); // D: ステータス
        var reason = dataRow[4];     // E: 理由
        var note = dataRow[5];       // F: 備考
        var time = dataRow[6];       // G: 時刻

        var overviewRow = 4 + s;

        // No
        overviewSheet.getRange(overviewRow, baseCol).setValue(no);

        // 氏名
        overviewSheet.getRange(overviewRow, baseCol + 1).setValue(name);

        // ステータス省略記号
        var shortStatus = STATUS_SHORT.hasOwnProperty(status) ? STATUS_SHORT[status] : status;
        overviewSheet.getRange(overviewRow, baseCol + 2).setValue(shortStatus);

        // 備考欄: ステータスに応じた情報を省略表示（最大6文字）
        var noteText = '';
        if (status === '欠席' && reason) {
          noteText = _truncate(String(reason), 6);
        } else if (status === '遅刻' && time) {
          noteText = String(time);
        } else if (status === '早退') {
          noteText = note ? _truncate(String(note), 6) : '';
        } else if (note) {
          noteText = _truncate(String(note), 6);
        }
        overviewSheet.getRange(overviewRow, baseCol + 3).setValue(noteText);

        // ステータスに応じた色分け
        var rowRange = overviewSheet.getRange(overviewRow, baseCol, 1, OVERVIEW_COLS_PER_CLASS);
        if (status === '転校') {
          rowRange.setBackground('#d0d0d0').setFontColor('#888888').setFontLine('line-through');
        } else if (status === '欠席') {
          rowRange.setBackground('#ffe0e0').setFontColor('#cc0000');
        } else if (status === '遅刻') {
          rowRange.setBackground('#fff8e0').setFontColor('#996600');
        } else if (status === '早退') {
          rowRange.setBackground('#e0f8ff').setFontColor('#006699');
        } else if (status === '保健室') {
          rowRange.setBackground('#fff0e0').setFontColor('#cc6600');
        } else if (status === '出停') {
          rowRange.setBackground('#f0e0ff').setFontColor('#6600cc');
        }
      }
    }
  }
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
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 600);

  var r = 1;

  // ── タイトル ──
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('📖 サイバー・コマンドセンター 運用マニュアル')
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
    ['1', '朝（出勤時）', 'スプレッドシートを開く → 日付が自動で今日に更新されます'],
    ['2', '朝のHR', '各クラスの出欠を入力： D列のプルダウンで「欠席」「遅刻」等を選択。E列に理由、F列に備考を記入'],
    ['3', '日中', '遅刻・早退・保健室等のステータス変更があればその都度更新'],
    ['4', '帰りのHR後', 'メニュー → 「📦 本日分をアーカイブ」でログに保存'],
    ['5', '翌日の朝', 'メニュー → 「🔄 翌日用にリセット」で全員を「出席」に戻す'],
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

  // ── 注意事項 ──
  sheet.getRange(r, 1, 1, 3).merge()
    .setValue('⚠ 休日は何もしなくてOK！ スプレッドシートを開かなければデータは動きません。')
    .setFontSize(11).setFontColor('#ffff00').setFontWeight('bold');
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
    ['📦 本日分をアーカイブ', '今日の出欠データをログシートに保存。二重実行防止付き'],
    ['🔄 翌日用にリセット', 'ステータスを全員「出席」に戻し、理由・備考をクリア'],
    ['📦+🔄 アーカイブ＆リセット', '上記2つをまとめて実行（帰りのHR後に便利）'],
    ['🚨 連続欠席チェック', '3日連続欠席の生徒にC列に🚨マークを表示'],
    ['📊 ダッシュボード更新', '分析シートにクラス別出席率、ステータス分布等のグラフを描画'],
    ['👤 転入生追加', 'クラス番号と氏名を入力→末尾に新番号で追加。設定シートの人数も自動+1'],
    ['🚪 転校処理（選択行）', '生徒の行を選択して実行→グレーアウト＋取り消し線。出欠集計から除外'],
    ['↩ 転校取り消し（選択行）', '誤って転校にした場合に元に戻す'],
    ['🔄 初期セットアップ', 'シート構造を初期化。クラス人数変更時に再実行'],
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
