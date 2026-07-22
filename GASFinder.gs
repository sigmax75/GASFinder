// ============================================================
// CONFIG
// ============================================================
var CONFIG = {
  INPUT_SHEET_NAME: 'FileID一覧',
  OUTPUT_SHEET_NAME: '結果',
  INPUT_SOURCE: 'sheet',
  HARDCODED_FILE_IDS: [],
  USE_APPS_SCRIPT_API: false,
  TIME_LIMIT_SEC: 240,
  TRIGGER_INTERVAL_MIN: 5,
  HIGHLIGHT_COLOR: '#FFFF00',
  HEADER_COLOR: '#4A86C8',
  HEADER_FONT_COLOR: '#FFFFFF'
};

// ============================================================
// PROP_KEYS
// ============================================================
var PROP_KEYS = {
  RESUME_INDEX: 'GASFINDER_RESUME_INDEX',
  TOTAL_COUNT: 'GASFINDER_TOTAL_COUNT'
};

// ============================================================
// Menu
// ============================================================

// onOpen - add custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GASFinder')
    .addItem('GAS調査を実行(手動)', 'runGASFinder')
    .addSeparator()
    .addItem('自動調査を開始', 'startAutoScan')
    .addItem('自動調査を停止', 'stopAutoScan')
    .addToUi();
}

// ============================================================
// Manual execution
// ============================================================

// runGASFinder - manual scan
function runGASFinder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileIds = getFileIds_(ss, true);

  if (fileIds.length === 0) {
    SpreadsheetApp.getUi().alert('チェック対象のFileIDが見つかりませんでした。');
    return;
  }

  var resultSheet = prepareResultSheet_(ss, true);

  var results = [];
  for (var i = 0; i < fileIds.length; i++) {
    var fileId = fileIds[i].trim();
    if (fileId === '') continue;

    Logger.log('Processing: ' + (i + 1) + '/' + fileIds.length + ' - ' + fileId);
    var result = checkFile_(fileId);
    results.push(result);
  }

  writeResults_(resultSheet, results);

  SpreadsheetApp.getUi().alert(
    '調査完了' + '\n' +
    '処理件数: ' + results.length + '件' + '\n' +
    '結果は' + CONFIG.OUTPUT_SHEET_NAME + 'シートをご確認ください。'
  );
}

// ============================================================
// Auto scan
// ============================================================

// startAutoScan - start auto scan with trigger
function startAutoScan() {
  var ui = SpreadsheetApp.getUi();

  deleteGASFinderTriggers_();
  clearProgress_();

  ScriptApp.newTrigger('runAutoScan')
    .timeBased()
    .everyMinutes(CONFIG.TRIGGER_INTERVAL_MIN)
    .create();

  ui.alert(
    '自動調査を開始します。' + '\n' +
    CONFIG.TRIGGER_INTERVAL_MIN + '分毎にトリガーが実行されます。' + '\n' +
    '初回実行を開始します。'
  );

  runAutoScan();
}

// stopAutoScan - stop auto scan
function stopAutoScan() {
  deleteGASFinderTriggers_();
  clearProgress_();

  SpreadsheetApp.getUi().alert(
    '自動調査を停止しました。' + '\n' + 'トリガーと進捗データを削除しました。'
  );
}

// runAutoScan - triggered auto scan with resume support
function runAutoScan() {
  var startTime = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();

  var fileIds = getFileIds_(ss, false);
  if (fileIds.length === 0) {
    Logger.log('GASFinder: No FileIDs found.');
    deleteGASFinderTriggers_();
    clearProgress_();
    return;
  }

  var resumeIndex = getResumeIndex_();
  var totalCount = fileIds.length;
  var isFirstRun = (resumeIndex === 0);

  var resultSheet;
  if (isFirstRun) {
    resultSheet = prepareResultSheet_(ss, true);
    props.setProperty(PROP_KEYS.TOTAL_COUNT, String(totalCount));
    Logger.log('GASFinder: Starting auto scan - total ' + totalCount + ' files');
  } else {
    resultSheet = prepareResultSheet_(ss, false);
    Logger.log('GASFinder: Resuming from ' + resumeIndex + '/' + totalCount);
  }

  writeStatus_(resultSheet, resumeIndex, totalCount, 'Processing...');

  var results = [];

  for (var i = resumeIndex; i < fileIds.length; i++) {
    var elapsed = (new Date().getTime() - startTime) / 1000;
    if (elapsed >= CONFIG.TIME_LIMIT_SEC) {
      saveResumeIndex_(i);
      Logger.log('GASFinder: Time limit - ' + i + '/' + totalCount + ' processed');

      if (results.length > 0) {
        appendResults_(resultSheet, results);
      }
      writeStatus_(resultSheet, i, totalCount, 'Suspended - next trigger');
      return;
    }

    var fileId = fileIds[i].trim();
    if (fileId === '') continue;

    Logger.log('Processing: ' + (i + 1) + '/' + totalCount + ' - ' + fileId);
    var result = checkFile_(fileId);
    results.push(result);
  }

  if (results.length > 0) {
    appendResults_(resultSheet, results);
  }
  writeStatus_(resultSheet, totalCount, totalCount, 'Completed');
  Logger.log('GASFinder: All done - ' + totalCount + ' files');

  deleteGASFinderTriggers_();
  clearProgress_();
}

// ============================================================
// Progress management
// ============================================================

// getResumeIndex_ - get resume index from properties
function getResumeIndex_() {
  var val = PropertiesService.getScriptProperties()
    .getProperty(PROP_KEYS.RESUME_INDEX);
  return val ? parseInt(val, 10) : 0;
}

// saveResumeIndex_ - save resume index
function saveResumeIndex_(index) {
  PropertiesService.getScriptProperties()
    .setProperty(PROP_KEYS.RESUME_INDEX, String(index));
}

// clearProgress_ - clear progress properties
function clearProgress_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_KEYS.RESUME_INDEX);
  props.deleteProperty(PROP_KEYS.TOTAL_COUNT);
}

// ============================================================
// Trigger management
// ============================================================

// deleteGASFinderTriggers_ - delete all runAutoScan triggers
function deleteGASFinderTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runAutoScan') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ============================================================
// Status display
// ============================================================

// writeStatus_ - show progress in G1
function writeStatus_(sheet, processed, total, status) {
  var statusText = '[Progress] ' + processed + '/' + total + ' - ' + status;
  sheet.getRange(1, 7).setValue(statusText);
  sheet.setColumnWidth(7, 350);
  Logger.log('GASFinder: ' + statusText);
}

// ============================================================
// FileID retrieval
// ============================================================

// getFileIds_ - get FileID list from sheet or hardcoded array
function getFileIds_(ss, useUi) {
  if (CONFIG.INPUT_SOURCE === 'hardcode') {
    return CONFIG.HARDCODED_FILE_IDS;
  }

  var sheet = ss.getSheetByName(CONFIG.INPUT_SHEET_NAME);
  if (!sheet) {
    if (useUi) {
      SpreadsheetApp.getUi().alert(
        CONFIG.INPUT_SHEET_NAME + 'シートが見つかりません。' + '\n' +
        'シートを作成してA列にFileIDを記載してください。'
      );
    } else {
      Logger.log('GASFinder: Input sheet not found - ' + CONFIG.INPUT_SHEET_NAME);
    }
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var ids = [];
  for (var i = 0; i < data.length; i++) {
    var val = String(data[i][0]).trim();
    if (val !== '' && val !== 'undefined' && val !== 'null') {
      ids.push(val);
    }
  }
  return ids;
}

// ============================================================
// File check
// ============================================================

// checkFile_ - check single file for bound GAS
function checkFile_(fileId) {
  var result = {
    fileId: fileId,
    fileName: '',
    owner: '',
    hasGas: false,
    scriptName: '',
    checkDate: new Date(),
    error: ''
  };

  try {
    var fileMeta = Drive.Files.get(fileId, { fields: 'name,owners' });
    result.fileName = fileMeta.name || '';
    if (fileMeta.owners && fileMeta.owners.length > 0) {
      result.owner = fileMeta.owners[0].emailAddress || '';
    }
  } catch (e) {
    Logger.log('DEBUG Drive.Files.get error: ' + String(e) + ' | message: ' + (e.message || 'none'));
    var errMsg = String(e.message || e);
    if (errMsg.indexOf('not found') !== -1 || errMsg.indexOf('404') !== -1) {
      result.error = 'ファイル不明';
    } else if (errMsg.indexOf('403') !== -1 || errMsg.indexOf('permission') !== -1 ||
               errMsg.indexOf('access') !== -1 || errMsg.indexOf('forbidden') !== -1) {
      result.error = 'アクセス不可';
    } else {
      result.error = 'Error: ' + errMsg;
    }
    return result;
  }

  var driveResult = checkGasByDriveApi_(fileId);
  if (driveResult.found) {
    result.hasGas = true;
    result.scriptName = driveResult.scriptName;
    return result;
  }

  if (CONFIG.USE_APPS_SCRIPT_API) {
    var apiResult = checkGasByAppsScriptApi_(fileId);
    if (apiResult.found) {
      result.hasGas = true;
      result.scriptName = apiResult.scriptName;
      return result;
    }
    if (apiResult.error) {
      result.scriptName = '(API check failed: ' + apiResult.error + ')';
    }
  }

  return result;
}

// ============================================================
// Drive API detection
// ============================================================

// checkGasByDriveApi_ - search for container-bound scripts via Drive API
function checkGasByDriveApi_(fileId) {
  var result = { found: false, scriptName: '' };

  try {
    var query = "'" + fileId + "' in parents" +
      " and mimeType = 'application/vnd.google-apps.script'" +
      " and trashed = false";
    var response = Drive.Files.list({
      q: query,
      fields: 'files(name)',
      pageSize: 10
    });

    if (response.files && response.files.length > 0) {
      result.found = true;
      var names = [];
      for (var i = 0; i < response.files.length; i++) {
        names.push(response.files[i].name);
      }
      result.scriptName = names.join(', ');
    }
  } catch (e) {
    Logger.log('Drive API search error (fileId: ' + fileId + '): ' + e.message);
  }

  return result;
}

// ============================================================
// Apps Script API detection (fallback)
// ============================================================

// checkGasByAppsScriptApi_ - check via Apps Script API
function checkGasByAppsScriptApi_(fileId) {
  var result = { found: false, scriptName: '', error: '' };

  try {
    var token = ScriptApp.getOAuthToken();
    var url = 'https://script.googleapis.com/v1/processes?' +
      'userProcessFilter.scriptId=' + fileId;

    var options = {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code === 200) {
      var json = JSON.parse(response.getContentText());
      if (json.processes && json.processes.length > 0) {
        result.found = true;
        result.scriptName = '(Detected via Apps Script API)';
      }
    } else if (code === 404) {
      // No script found - normal
    } else {
      result.error = 'HTTP ' + code;
    }
  } catch (e) {
    result.error = String(e.message || e);
    Logger.log('Apps Script API error (fileId: ' + fileId + '): ' + e.message);
  }

  return result;
}

// ============================================================
// Result sheet setup
// ============================================================

// prepareResultSheet_ - prepare or create the result sheet
function prepareResultSheet_(ss, clearSheet) {
  var sheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
    clearSheet = true;
  }

  if (clearSheet) {
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setBackground(null);

    var headers = ['FileID', 'ファイル名', 'オーナー', 'GAS有無', 'スクリプト名', 'チェック日時'];
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground(CONFIG.HEADER_COLOR);
    headerRange.setFontColor(CONFIG.HEADER_FONT_COLOR);
    headerRange.setFontWeight('bold');

    sheet.setColumnWidth(1, 320);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 220);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 160);

    sheet.setFrozenRows(1);
  }

  return sheet;
}

// ============================================================
// Result output
// ============================================================

// writeResults_ - write results from row 2 (manual mode)
function writeResults_(sheet, results) {
  if (results.length === 0) return;

  var rows = buildResultRows_(results);

  var dataRange = sheet.getRange(2, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);

  highlightGasRows_(sheet, results, 2);
}

// appendResults_ - append results below existing data (auto mode)
function appendResults_(sheet, results) {
  if (results.length === 0) return;

  var rows = buildResultRows_(results);

  var lastRow = sheet.getLastRow();
  var startRow = (lastRow < 1) ? 2 : lastRow + 1;

  var dataRange = sheet.getRange(startRow, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);

  highlightGasRows_(sheet, results, startRow);
}

// buildResultRows_ - build 2D array from results
function buildResultRows_(results) {
  var rows = [];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    var gasStatus;
    if (r.error) {
      gasStatus = r.error;
    } else {
      gasStatus = r.hasGas ? 'あり' : 'なし';
    }

    rows.push([
      r.fileId,
      r.fileName,
      r.owner,
      gasStatus,
      r.scriptName,
      formatDate_(r.checkDate)
    ]);
  }
  return rows;
}

// highlightGasRows_ - highlight rows with GAS in yellow
function highlightGasRows_(sheet, results, startRow) {
  for (var i = 0; i < results.length; i++) {
    if (results[i].hasGas) {
      sheet.getRange(startRow + i, 1, 1, 6)
        .setBackground(CONFIG.HIGHLIGHT_COLOR);
    }
  }
}

// ============================================================
// Utility
// ============================================================

// formatDate_ - format date for display
function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}
