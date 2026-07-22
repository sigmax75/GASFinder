// ============================================================
// CONFIG
// ============================================================
// No need to add Drive Advanced Service - uses DriveApp and SpreadsheetApp only
var CONFIG = {
  INPUT_SHEET_NAME: 'FileID一覧',
  OUTPUT_SHEET_NAME: '結果',
  INPUT_SOURCE: 'sheet',
  HARDCODED_FILE_IDS: [],
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
    .addSeparator()
    .addItem('再開位置を設定', 'setResumeIndex')
    .addToUi();
}

// setResumeIndex - set resume index manually
function setResumeIndex() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Resume Index',
    'Enter the row number to resume from (0-based index):',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    var index = parseInt(response.getResponseText(), 10);
    if (isNaN(index) || index < 0) {
      ui.alert('Invalid number.');
      return;
    }
    saveResumeIndex_(index);
    ui.alert('Resume index set to ' + index + '.');
  }
}

// ============================================================
// Manual execution
// ============================================================

// runGASFinder - manual scan with lock
function runGASFinder() {
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(10000);
  if (!hasLock) {
    SpreadsheetApp.getUi().alert('別の処理が実行中です。しばらく待ってから再実行してください。');
    return;
  }

  try {
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
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// Auto scan
// ============================================================

// startAutoScan - start auto scan with trigger
function startAutoScan() {
  var ui = SpreadsheetApp.getUi();

  var existingTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === 'runAutoScan') {
      ui.alert('既に自動調査が実行中です。' + '\n' + '停止してから再度開始してください。');
      return;
    }
  }

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
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(10000);
  if (!hasLock) {
    Logger.log('GASFinder: Another instance is running. Skipping.');
    return;
  }

  try {
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
  } finally {
    lock.releaseLock();
  }
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

// checkFile_ - check single file for bound GAS via indirect detection
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
    var file = DriveApp.getFileById(fileId);
    result.fileName = file.getName();
    var owner = file.getOwner();
    result.owner = owner ? owner.getEmail() : '';
  } catch (e) {
    Logger.log('DEBUG DriveApp.getFileById error: ' + String(e) + ' | message: ' + (e.message || 'none'));
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

  var indirectResult = checkGasIndirect_(fileId);
  if (indirectResult.found) {
    result.hasGas = true;
    result.scriptName = indirectResult.scriptName;
  } else if (indirectResult.details && indirectResult.details.indexOf('check failed') === 0) {
    result.scriptName = indirectResult.details;
  }

  return result;
}

// ============================================================
// Indirect GAS detection via custom function scan
// ============================================================

// checkGasIndirect_ - detect GAS by finding custom functions in formulas
function checkGasIndirect_(fileId) {
  var result = { found: false, scriptName: '', details: '' };
  try {
    var ss = SpreadsheetApp.openById(fileId);
    var sheets = ss.getSheets();
    var customFuncs = [];

    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var range = sheet.getDataRange();
      var formulas = range.getFormulas();

      for (var r = 0; r < formulas.length; r++) {
        for (var c = 0; c < formulas[r].length; c++) {
          var formula = formulas[r][c];
          if (formula === '') continue;

          // Extract function names from formula
          var funcMatches = formula.match(/[A-Za-z_][A-Za-z0-9_]*\s*\(/g);
          if (!funcMatches) continue;

          for (var m = 0; m < funcMatches.length; m++) {
            var funcName = funcMatches[m].replace(/\s*\($/, '').toUpperCase();
            if (!isBuiltinFunction_(funcName)) {
              customFuncs.push(funcName.toLowerCase());
            }
          }
        }
      }
    }

    if (customFuncs.length > 0) {
      // Remove duplicates
      var unique = [];
      for (var i = 0; i < customFuncs.length; i++) {
        if (unique.indexOf(customFuncs[i]) === -1) {
          unique.push(customFuncs[i]);
        }
      }
      result.found = true;
      result.scriptName = unique.join(', ');
      result.details = 'custom functions detected';
    }
  } catch (e) {
    Logger.log('Indirect check error (fileId: ' + fileId + '): ' + e.message);
    result.details = 'check failed: ' + e.message;
  }
  return result;
}

// ============================================================
// Built-in function list
// ============================================================

// isBuiltinFunction_ - check if a function name is a standard Sheets function
function isBuiltinFunction_(name) {
  return BUILTIN_FUNCTIONS_.indexOf(name) !== -1;
}

// BUILTIN_FUNCTIONS_ - list of standard Google Sheets functions (uppercase)
var BUILTIN_FUNCTIONS_ = [
  'ABS', 'ACCRINT', 'ACCRINTM', 'ACOS', 'ACOSH', 'ADDRESS', 'AMORLINC', 'AND',
  'ARRAYFORMULA', 'ASC', 'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH', 'AVEDEV',
  'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'BASE', 'BETA.DIST', 'BETA.INV', 'BIN2DEC',
  'BIN2HEX', 'BIN2OCT', 'BINOM.DIST', 'BINOM.DIST.RANGE', 'BINOM.INV', 'BITAND', 'BITLSHIFT', 'BITOR',
  'BITRSHIFT', 'BITXOR', 'BYCOL', 'BYROW', 'CEILING', 'CELL', 'CHAR', 'CHISQ.DIST',
  'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'CHOOSE', 'CHOOSECOLS', 'CHOOSEROWS', 'CLEAN',
  'CODE', 'COLUMN', 'COLUMNS', 'COMBIN', 'COMBINA', 'COMPLEX', 'CONCAT', 'CONCATENATE',
  'CONFIDENCE', 'CONFIDENCE.NORM', 'CONFIDENCE.T', 'CONVERT', 'CORREL', 'COS', 'COSH', 'COT',
  'COTH', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'COUPDAYBS', 'COUPDAYS',
  'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD', 'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH',
  'CUMIPMT', 'CUMPRINC', 'DATE', 'DATEDIF', 'DATEVALUE', 'DAVERAGE', 'DAY', 'DAYS',
  'DAYS360', 'DB', 'DCOUNT', 'DCOUNTA', 'DDB', 'DEC2BIN', 'DEC2HEX', 'DEC2OCT',
  'DECIMAL', 'DEGREES', 'DELTA', 'DETECT.LANGUAGE', 'DETECTLANGUAGE', 'DEVSQ', 'DGET', 'DISC',
  'DMAX', 'DMIN', 'DOLLAR', 'DOLLARDE', 'DOLLARFR', 'DPRODUCT', 'DSTDEV', 'DSTDEVP',
  'DSUM', 'DURATION', 'DVAR', 'DVARP', 'EDATE', 'EFFECT', 'ENCODEURL', 'EOMONTH',
  'ERF', 'ERF.PRECISE', 'ERFC', 'ERFC.PRECISE', 'ERROR.TYPE', 'EVEN', 'EXACT', 'EXP',
  'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FACT', 'FACTDOUBLE',
  'FALSE', 'FIELDVALUE', 'FILTER', 'FIND', 'FINDB', 'FISHER', 'FISHERINV', 'FIXED',
  'FLATTEN', 'FLOOR', 'FORECAST', 'FORECAST.LINEAR', 'FORMULATEXT', 'FREQUENCY', 'FV', 'FVSCHEDULE',
  'GAMMA', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN', 'GAMMALN.PRECISE', 'GAUSS', 'GCD', 'GEOMEAN',
  'GESTEP', 'GETPIVOTDATA', 'GOOGLEFINANCE', 'GOOGLETRANSLATE', 'GROWTH', 'HARMEAN', 'HEX2BIN', 'HEX2DEC',
  'HEX2OCT', 'HLOOKUP', 'HOUR', 'HSTACK', 'HYPGEOM.DIST', 'HYPERLINK', 'IF', 'IFERROR',
  'IFNA', 'IFS', 'IMAGE', 'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS',
  'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMDIV', 'IMEXP', 'IMLN', 'IMLOG10',
  'IMLOG2', 'IMPOWER', 'IMPORTDATA', 'IMPORTFEED', 'IMPORTHTML', 'IMPORTRANGE', 'IMPORTXML', 'IMPRODUCT',
  'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSQRT', 'IMSUB', 'IMSUM',
  'IMTAN', 'INDEX', 'INDIRECT', 'INFO', 'INT', 'INTERCEPT', 'INTRATE', 'IPMT',
  'IRR', 'ISBLANK', 'ISERR', 'ISERROR', 'ISEVEN', 'ISFORMULA', 'ISLOGICAL', 'ISNA',
  'ISNONTEXT', 'ISNUMBER', 'ISODD', 'ISOWEEKNUM', 'ISPMT', 'ISREF', 'ISTEXT', 'ISURL',
  'JIS', 'JOIN', 'KURT', 'LAMBDA', 'LARGE', 'LCM', 'LEFT', 'LEFTB',
  'LEN', 'LENB', 'LET', 'LINEST', 'LN', 'LOG', 'LOG10', 'LOGEST',
  'LOGNORM.DIST', 'LOGNORM.INV', 'LOOKUP', 'LOWER', 'MAKEARRAY', 'MAP', 'MARGINOFERROR', 'MATCH',
  'MAX', 'MAXA', 'MAXIFS', 'MDETERM', 'MDURATION', 'MEDIAN', 'MID', 'MIDB',
  'MIN', 'MINA', 'MINIFS', 'MINVERSE', 'MINUTE', 'MIRR', 'MMULT', 'MOD',
  'MODE', 'MODE.MULT', 'MODE.SNGL', 'MONTH', 'MROUND', 'MULTINOMIAL', 'MUNIT', 'N',
  'NA', 'NEGBINOM.DIST', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'NOMINAL', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST',
  'NORM.S.INV', 'NOT', 'NOW', 'NPER', 'NPV', 'OCT2BIN', 'OCT2DEC', 'OCT2HEX',
  'ODD', 'OFFSET', 'OR', 'PDURATION', 'PEARSON', 'PERCENTILE', 'PERCENTILE.EXC', 'PERCENTILE.INC',
  'PERCENTRANK', 'PERCENTRANK.EXC', 'PERCENTRANK.INC', 'PERMUT', 'PERMUTATIONA', 'PHI', 'PI', 'PMT',
  'POISSON', 'POISSON.DIST', 'POWER', 'PPMT', 'PRICE', 'PRICEDISC', 'PRICEMAT', 'PROB',
  'PRODUCT', 'PROPER', 'PV', 'QUARTILE', 'QUARTILE.EXC', 'QUARTILE.INC', 'QUERY', 'QUOTIENT',
  'RADIANS', 'RAND', 'RANDBETWEEN', 'RANK', 'RANK.AVG', 'RANK.EQ', 'RATE', 'RECEIVED',
  'REDUCE', 'REGEXEXTRACT', 'REGEXMATCH', 'REGEXREPLACE', 'REPLACE', 'REPLACEB', 'REPT', 'RIGHT',
  'RIGHTB', 'ROMAN', 'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'ROW', 'ROWS', 'RRI',
  'RSQ', 'SCAN', 'SEARCH', 'SEARCHB', 'SEC', 'SECH', 'SECOND', 'SEQUENCE',
  'SERIESSUM', 'SHEET', 'SHEETS', 'SIGN', 'SIN', 'SINH', 'SKEW', 'SKEW.P',
  'SLN', 'SLOPE', 'SMALL', 'SORT', 'SORTBY', 'SPARKLINE', 'SPLIT', 'SQRT',
  'SQRTPI', 'STANDARDIZE', 'STDEV', 'STDEV.P', 'STDEV.S', 'STDEVA', 'STDEVP', 'STDEVPA',
  'STEYX', 'SUBSTITUTE', 'SUBTOTAL', 'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ',
  'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'SWITCH', 'SYD', 'T', 'T.DIST', 'T.DIST.2T',
  'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'TAN', 'TANH', 'TBILLEQ', 'TBILLPRICE',
  'TBILLYIELD', 'TEXT', 'TEXTJOIN', 'TIME', 'TIMEVALUE', 'TOCOL', 'TODAY', 'TO_DATE',
  'TO_DOLLARS', 'TO_PERCENT', 'TO_PURE_NUMBER', 'TO_TEXT', 'TOROW', 'TRANSPOSE', 'TREND', 'TRIM',
  'TRIMMEAN', 'TRUE', 'TRUNC', 'TYPE', 'UNICHAR', 'UNICODE', 'UNIQUE', 'UPPER',
  'VALUE', 'VAR', 'VAR.P', 'VAR.S', 'VARA', 'VARP', 'VARPA', 'VDB',
  'VLOOKUP', 'VSTACK', 'WEEKDAY', 'WEEKNUM', 'WEIBULL', 'WEIBULL.DIST', 'WORKDAY', 'WORKDAY.INTL',
  'WRAPCOLS', 'WRAPROWS', 'XIRR', 'XLOOKUP', 'XMATCH', 'XNPV', 'XOR', 'YEAR',
  'YEARFRAC', 'YIELD', 'YIELDDISC', 'YIELDMAT', 'Z.TEST', 'ZTEST'
];

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
    sheet.setColumnWidth(4, 130);
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
    } else if (r.hasGas) {
      gasStatus = 'あり (custom func)';
    } else if (r.scriptName && r.scriptName.indexOf('check failed') === 0) {
      gasStatus = 'check failed';
    } else {
      gasStatus = 'なし';
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
