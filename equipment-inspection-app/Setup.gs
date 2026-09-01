function addLiftInspection() {
  const ss = SpreadsheetApp.openById('10NiNUNnb6ql4KV3U_G_7sFJW7gPi7uLB8yB8X5i8t-c');
  const FONT = "Yu Gothic";
  const HEADER_BG = "#1F4E78";
  const KEY_BG = "#FCE4D6";
  const ITEM_BG = "#E2EFDA";
  const master = ss.getSheetByName('設備マスタ');
  const existing = master.getRange(2, 1, Math.max(master.getLastRow() - 1, 0), 1).getValues().flat();
  if (existing.indexOf('LF-01') === -1) {
    master.appendRow(['LF-01', 'リフト', 'フォークリフト', 'LF-01', '', '野田組', '', '作業前', '管理番号は仮。実際の番号に置換してください']);
  }
  if (ss.getSheetByName('リフト点検記録')) {
    Logger.log('リフト点検記録は既に存在します。スキップしました。');
    return;
  }
  const sheet = ss.insertSheet('リフト点検記録');
  const items = [
    'ハンドブレーキ_効き具合良好','クラッチペダル_遊びと切れ良好','フォークアタッチメント_損傷なし',
    'バックレストヘッドガード_損傷なし','チェーン_左右張り同じ','マスト_上下前後傾の作動良好',
    'シリンダーホース_油漏れなし','タイヤ取付けボルトナット_ゆるみなし','タイヤ_空気圧損傷なし',
    'ギヤーボックス_油漏れなし','ハンドル_遊びと切れ良好','灯火計器類_作動良好',
    'ホーン_鳴り良好','排気_色正常','異音_なし',
  ];
  const head = ['記録ID', '点検日', '設備', '点検者'].concat(items).concat(['総合判定', '異常内容', '処置内容', '写真', '登録日時']);
  sheet.getRange(1, 1, 1, head.length).setValues([head]);
  const range = sheet.getRange(1, 1, 1, head.length);
  range.setBackground(HEADER_BG).setFontColor('#FFFFFF').setFontWeight('bold')
       .setFontFamily(FONT).setFontSize(11).setHorizontalAlignment('center')
       .setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setBackground(KEY_BG).setFontColor('#000000');
  for (let c = 5; c < 5 + items.length; c++) {
    sheet.getRange(1, c).setBackground(ITEM_BG).setFontColor('#000000');
  }
  sheet.setColumnWidths(1, head.length, 150);
  sheet.setFrozenColumns(4);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['○', '✕', '該当なし'], true)
    .setAllowInvalid(false)
    .build();
  for (let c = 5; c < 5 + items.length; c++) {
    sheet.getRange(2, c, 299, 1).setDataValidation(rule);
  }
  Logger.log('リフト点検記録を追加しました。設備マスタにも LF-01 を登録しました。');
}
