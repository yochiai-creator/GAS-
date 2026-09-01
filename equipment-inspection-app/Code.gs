const SPREADSHEET_ID = '10NiNUNnb6ql4KV3U_G_7sFJW7gPi7uLB8yB8X5i8t-c';
const FIXED_HEAD = ['記録ID','点検日','設備','点検者'];
const FIXED_TAIL = ['総合判定','異常内容','処置内容','写真','登録日時'];
const CATEGORY_SHEET = {
  '溶接機': '溶接機点検記録',
  'クレーン': 'クレーン点検記録',
  '局所排気装置': '局所排気装置点検記録',
  'コンプレッサ': 'コンプレッサ点検記録',
  'リフト': 'リフト点検記録'
};
function doGet(e) {
  const eq = (e && e.parameter && e.parameter.eq) ? e.parameter.eq : '';
  const t = HtmlService.createTemplateFromFile('Index');
  t.eq = eq;
  return t.evaluate()
    .setTitle('設備点検')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function getEquipmentList() {
  const sh = getSpreadsheet_().getSheetByName('設備マスタ');
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1).filter(function(r) { return r[0]; });
  return rows.map(function(r) {
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = r[i]; });
    return obj;
  });
}
function getEquipmentInfo_(eqId) {
  const list = getEquipmentList();
  for (let i = 0; i < list.length; i++) {
    if (list[i]['設備ID'] === eqId) return list[i];
  }
  return null;
}
function getFormForEquipment(eqId) {
  const info = getEquipmentInfo_(eqId);
  if (!info) return { error: '設備が見つかりません: ' + eqId };
  const category = info['設備区分'];
  const sheetName = CATEGORY_SHEET[category];
  if (!sheetName) return { error: '未対応の設備区分: ' + category };
  const sh = getSpreadsheet_().getSheetByName(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const hasKubun = headers.indexOf('点検区分') > -1;
  const excluded = FIXED_HEAD.concat(FIXED_TAIL).concat(hasKubun ? ['点検区分'] : []);
  const items = headers.filter(function(h) { return excluded.indexOf(h) === -1; });
  return {
    equipment: info,
    sheetName: sheetName,
    hasKubun: hasKubun,
    items: items,
    email: Session.getActiveUser().getEmail()
  };
}
function submitInspection(payload) {
  const info = getEquipmentInfo_(payload.eqId);
  if (!info) throw new Error('設備が見つかりません');
  const category = info['設備区分'];
  const sheetName = CATEGORY_SHEET[category];
  const sh = getSpreadsheet_().getSheetByName(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(function(h) {
    if (h === '記録ID') return Utilities.getUuid().slice(0, 8);
    if (h === '点検日') return new Date();
    if (h === '設備') return payload.eqId;
    if (h === '点検者') return Session.getActiveUser().getEmail();
    if (h === '点検区分') return payload.kubun || '';
    if (h === '総合判定') return payload.hantei || '';
    if (h === '異常内容') return payload.ijou || '';
    if (h === '処置内容') return payload.shochi || '';
    if (h === '写真') return '';
    if (h === '登録日時') return new Date();
    return (payload.items && payload.items[h]) ? payload.items[h] : '';
  });
  sh.appendRow(row);
  return { ok: true };
}
function getHistory(eqId, limit) {
  const info = getEquipmentInfo_(eqId);
  if (!info) return [];
  const category = info['設備区分'];
  const sheetName = CATEGORY_SHEET[category];
  const sh = getSpreadsheet_().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idxEq = headers.indexOf('設備');
  const idxDate = headers.indexOf('点検日');
  const idxUser = headers.indexOf('点検者');
  const idxHantei = headers.indexOf('総合判定');
  const idxIjou = headers.indexOf('異常内容');
  const idxTouroku = headers.indexOf('登録日時');
  const rows = values.slice(1).filter(function(r) { return r[idxEq] === eqId; });
  rows.sort(function(a, b) { return new Date(b[idxTouroku]) - new Date(a[idxTouroku]); });
  return rows.slice(0, limit || 20).map(function(r) {
    return {
      date: Utilities.formatDate(new Date(r[idxTouroku]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
      user: r[idxUser],
      hantei: r[idxHantei],
      ijou: r[idxIjou]
    };
  });
}
function getAllLogs(limit) {
  const eqList = getEquipmentList();
  const eqMap = {};
  eqList.forEach(function(e) { eqMap[e['設備ID']] = e; });
  let all = [];
  Object.keys(CATEGORY_SHEET).forEach(function(cat) {
    const sheetName = CATEGORY_SHEET[cat];
    const sh = getSpreadsheet_().getSheetByName(sheetName);
    if (!sh) return;
    const values = sh.getDataRange().getValues();
    const headers = values[0];
    const idxEq = headers.indexOf('設備');
    const idxUser = headers.indexOf('点検者');
    const idxHantei = headers.indexOf('総合判定');
    const idxIjou = headers.indexOf('異常内容');
    const idxTouroku = headers.indexOf('登録日時');
    values.slice(1).forEach(function(r) {
      if (!r[idxEq]) return;
      const info = eqMap[r[idxEq]];
      all.push({
        eqName: info ? info['設備名'] : r[idxEq],
        user: r[idxUser],
        hantei: r[idxHantei],
        ijou: r[idxIjou],
        touroku: r[idxTouroku]
      });
    });
  });
  all.sort(function(a, b) { return new Date(b.touroku) - new Date(a.touroku); });
  return all.slice(0, limit || 50).map(function(r) {
    return {
      date: Utilities.formatDate(new Date(r.touroku), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
      eqName: r.eqName,
      user: r.user,
      hantei: r.hantei,
      ijou: r.ijou
    };
  });
}
function getCategoryList() {
  return Object.keys(CATEGORY_SHEET);
}
function getCategoryInfo() {
  return Object.keys(CATEGORY_SHEET).map(function(cat) {
    const sh = getSpreadsheet_().getSheetByName(CATEGORY_SHEET[cat]);
    const headers = sh ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    return { category: cat, hasKubun: headers.indexOf('点検区分') > -1 };
  });
}
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
function getEquipmentQrList() {
  const baseUrl = getWebAppUrl();
  return getEquipmentList().map(function(e) {
    return {
      eqId: e['設備ID'],
      name: e['設備名'],
      category: e['設備区分'],
      url: baseUrl + '?eq=' + encodeURIComponent(e['設備ID'])
    };
  });
}
function addInspectionItem(payload) {
  const category = (payload.category || '').trim();
  const itemName = (payload.itemName || '').trim();
  const kubun = (payload.kubun || '').trim();
  if (!category || !itemName) throw new Error('設備区分と項目名は必須です');
  const sheetName = CATEGORY_SHEET[category];
  if (!sheetName) throw new Error('未対応の設備区分です: ' + category);
  const sh = getSpreadsheet_().getSheetByName(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const hasKubun = headers.indexOf('点検区分') > -1;
  if (hasKubun && !kubun) throw new Error('この設備区分では点検区分の指定が必要です');
  const fullName = (hasKubun && kubun) ? (kubun + '_' + itemName) : itemName;
  if (headers.indexOf(fullName) > -1) {
    throw new Error('この項目は既に存在します: ' + fullName);
  }
  const tailIdx = headers.indexOf(FIXED_TAIL[0]);
  if (tailIdx === -1) throw new Error('シート構成が想定と異なります（' + FIXED_TAIL[0] + '列が見つかりません）');
  const insertCol = tailIdx + 1;
  sh.insertColumnBefore(insertCol);
  const headerCell = sh.getRange(1, insertCol);
  headerCell.setValue(fullName)
    .setBackground('#E2EFDA')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['○', '✕', '該当なし'], true)
    .setAllowInvalid(false)
    .build();
  const rowCount = Math.max(sh.getMaxRows() - 1, 299);
  sh.getRange(2, insertCol, rowCount, 1).setDataValidation(rule);
  return { ok: true, item: fullName };
}
function addEquipment(payload) {
  if (!payload.eqId || !payload.category || !payload.name) {
    throw new Error('設備ID・区分・設備名は必須です');
  }
  if (!CATEGORY_SHEET[payload.category]) {
    throw new Error('未対応の設備区分です: ' + payload.category);
  }
  const list = getEquipmentList();
  if (list.some(function(e) { return e['設備ID'] === payload.eqId; })) {
    throw new Error('この設備IDは既に登録されています: ' + payload.eqId);
  }
  const sh = getSpreadsheet_().getSheetByName('設備マスタ');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(function(h) {
    if (h === '設備ID') return payload.eqId;
    if (h === '設備区分') return payload.category;
    if (h === '設備名') return payload.name;
    if (h === '管理番号') return payload.manageNo || payload.eqId;
    if (h === '設置場所') return payload.location || '';
    if (h === '部署') return payload.dept || '';
    if (h === '責任者') return payload.owner || '';
    if (h === '点検頻度') return payload.freq || '';
    if (h === '備考') return payload.note || '';
    return '';
  });
  sh.appendRow(row);
  return { ok: true };
}
