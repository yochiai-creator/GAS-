const TEMPLATE_SHEET = {
  '溶接機': '書式_溶接機',
  'クレーン': '書式_クレーン',
  '局所排気装置': '書式_局所排気',
  'コンプレッサ': '書式_コンプレッサ',
  'リフト': '書式_リフト'
};
function monthDateRange_(year, month) {
  return { start: new Date(year, month - 1, 1), end: new Date(year, month, 1) };
}
function getItemsForSheet_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const hasKubun = headers.indexOf('点検区分') > -1;
  const excluded = FIXED_HEAD.concat(FIXED_TAIL).concat(hasKubun ? ['点検区分'] : []);
  return headers.filter(function(h) { return excluded.indexOf(h) === -1; });
}
function getPdfRowMap_() {
  const sh = getSpreadsheet_().getSheetByName('PDF行設定');
  if (!sh) return {};
  const values = sh.getDataRange().getValues();
  const map = {};
  values.slice(1).forEach(function(r) {
    if (!r[0] || !r[1] || !r[2]) return;
    map[r[0] + '||' + r[1]] = r[2];
  });
  return map;
}
function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}
function findDayHeader_(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      if (data[r][c] === 1 && data[r][c + 1] === 2) {
        return { row: r + 1, col: c + 1 };
      }
    }
  }
  return null;
}
function generateMonthlyPDF(eqId, year, month) {
  const ss = getSpreadsheet_();
  const info = getEquipmentInfo_(eqId);
  if (!info) throw new Error('設備が見つかりません: ' + eqId);
  const category = info['設備区分'];
  const sheetName = CATEGORY_SHEET[category];
  const templateName = TEMPLATE_SHEET[category];
  if (!sheetName) throw new Error('未対応の設備区分です: ' + category);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('データシートが見つかりません: ' + sheetName);
  const template = ss.getSheetByName(templateName);
  if (!template) throw new Error('書式テンプレートが見つかりません: ' + templateName + '（先にシートをコピーしてください）');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const items = getItemsForSheet_(sh);
  const idxEq = headers.indexOf('設備');
  const idxDate = headers.indexOf('点検日');
  const range = monthDateRange_(year, month);
  const values = sh.getDataRange().getValues();
  const records = values.slice(1).filter(function(r) {
    if (r[idxEq] !== eqId) return false;
    const d = new Date(r[idxDate]);
    return d >= range.start && d < range.end;
  });
  const dayMap = {};
  records.forEach(function(r) {
    const day = new Date(r[idxDate]).getDate();
    if (!dayMap[day]) dayMap[day] = {};
    items.forEach(function(item) {
      dayMap[day][item] = r[headers.indexOf(item)];
    });
  });
  const work = template.copyTo(ss);
  ss.setActiveSheet(work);
  ss.moveActiveSheet(ss.getNumSheets());
  const head = findDayHeader_(work);
  if (!head) throw new Error('テンプレート内に日付ヘッダー（1,2,3...）が見つかりません: ' + templateName);
  const pdfRowMap = getPdfRowMap_();
  const placedItems = [];
  const legacyItems = [];
  items.forEach(function(item) {
    const mappedRow = pdfRowMap[category + '||' + item];
    if (mappedRow) {
      placedItems.push({ item: item, row: head.row + Number(mappedRow) });
    } else {
      legacyItems.push(item);
    }
  });
  legacyItems.forEach(function(item, i) {
    placedItems.push({ item: item, row: head.row + 1 + i });
  });
  // 複数項目が同じテンプレート行を共有する場合（例: 溶接機の「作業環境」「溶接装置」）は、
  // 1つでも✕があれば✕、それ以外で1つでも○があれば○、全て該当なしなら該当なしを書く。
  const rowGroups = {};
  placedItems.forEach(function(entry) {
    if (!rowGroups[entry.row]) rowGroups[entry.row] = [];
    rowGroups[entry.row].push(entry.item);
  });
  Object.keys(rowGroups).forEach(function(rowKey) {
    const row = Number(rowKey);
    const groupItems = rowGroups[rowKey];
    for (let d = 1; d <= 31; d++) {
      const values = groupItems
        .map(function(item) { return dayMap[d] ? dayMap[d][item] : ''; })
        .filter(function(v) { return v; });
      if (values.length === 0) continue;
      let val;
      if (values.indexOf('✕') > -1) val = '✕';
      else if (values.indexOf('○') > -1) val = '○';
      else val = values[0];
      const cell = work.getRange(row, head.col + d - 1);
      cell.setValue(val);
      if (val === '✕') cell.setFontColor('#D93025').setFontWeight('bold');
    }
  });
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&size=A4&portrait=false&fitw=true&gridlines=true&gid=' + work.getSheetId();
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const blob = response.getBlob().setName(sheetName + '_' + info['設備ID'] + '_' + year + '-' + ('0' + month).slice(-2) + '.pdf');
  getOrCreateFolder_('点検表PDF').createFile(blob);
  ss.deleteSheet(work);
  return blob.getName();
}
function generateAllMonthlyPDFs(year, month) {
  const list = getEquipmentList();
  const results = [];
  list.forEach(function(e) {
    try {
      results.push(generateMonthlyPDF(e['設備ID'], year, month));
    } catch (err) {
      results.push('エラー(' + e['設備ID'] + '): ' + err.message);
    }
  });
  Logger.log(results.join('\n'));
  return results;
}
function monthlyTriggerHandler() {
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  generateAllMonthlyPDFs(lastMonth.getFullYear(), lastMonth.getMonth() + 1);
}
function createMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'monthlyTriggerHandler') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('monthlyTriggerHandler')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();
}
function testPDF() {
  generateAllMonthlyPDFs(2026, 8);
}
