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
function setPdfRowMapping_(category, itemName, row) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('PDF行設定');
  if (!sh) {
    sh = ss.insertSheet('PDF行設定');
    sh.getRange(1, 1, 1, 4).setValues([['設備区分', '項目名', 'テンプレート行番号（日付ヘッダー行から何行下か）', '登録日時']]);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  const values = sh.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (values[r][0] === category && values[r][1] === itemName) {
      sh.getRange(r + 1, 3).setValue(row);
      sh.getRange(r + 1, 4).setValue(new Date());
      return;
    }
  }
  sh.appendRow([category, itemName, row, new Date()]);
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
