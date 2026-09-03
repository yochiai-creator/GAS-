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

// 既存項目のPDFテンプレート行番号を一括登録する（1回限り実行）。
// 書式_溶接機・書式_クレーンの実物を1行ずつ確認して求めた値。
// 局所排気装置・コンプレッサは元々1項目=1行が連番で並んでおり、
// generateMonthlyPDFの従来ロジック（先頭から順番に埋める方式）で
// すでに正しく動くため、ここには含めていない。
//
// 溶接機の「作業環境」「溶接装置」は、テンプレート上は複数項目が1行に
// まとまっており、結果を書き込むマスが1つしか無い。この6項目は
// 「1つでも✕なら✕、それ以外で1つでも○があれば○、全て該当なしなら該当なし」
// というルールで1マスにまとめて書き込む（generateMonthlyPDF側で対応済み）。
// そのため、同じ行番号を複数の項目に割り当てている。
//
// 対象外（要修正、リフトのみ別関数で対応）:
// - リフト全項目
//   → 書式_リフトには「1,2,3…」という日付の数字が無く「日」という文字が
//     並んでいるだけなので、generateMonthlyPDFが日付ヘッダーを見つけられず
//     常にエラーになる。fixLiftTemplateDayHeader()を先に実行すること。
function backfillPdfRowMappings() {
  const mappings = [
    // 溶接機: 1行1項目として分離できる3項目
    ['溶接機', 'ホルダー_破損汚れなし', 3],
    ['溶接機', '溶接用ケーブル_被覆の損傷劣化なし', 4],
    ['溶接機', '取扱責任者_表示あり', 5],
    // 溶接機: 「作業環境」の2項目は同じ1行を共有（1つでも✕なら✕ルール）
    ['溶接機', '作業環境_引火物爆発物可燃物なし', 1],
    ['溶接機', '作業環境_整理整頓されている', 1],
    // 溶接機: 「溶接装置」の4項目は同じ1行を共有（1つでも✕なら✕ルール）
    ['溶接機', '溶接装置_本体スイッチ類の損傷なし', 2],
    ['溶接機', '溶接装置_カバー類の損傷なし', 2],
    ['溶接機', '溶接装置_接続箇所の緩みなし', 2],
    ['溶接機', '溶接装置_確実に作動する', 2],
    // クレーン: 10項目とも1行1項目だが、行番号は連番ではない
    ['クレーン', '巻上げワイヤーロープ_異常なし', 1],
    ['クレーン', '作動_上下東西南北円滑', 12],
    ['クレーン', 'ストッパー_脱落なし', 14],
    ['クレーン', 'リミットスイッチ_確実に作動', 15],
    ['クレーン', 'フックブロック_円滑に回転', 16],
    ['クレーン', 'フックナット_緩み変形なし', 18],
    ['クレーン', 'ドラム_正しく巻付け', 20],
    ['クレーン', 'ワイヤー外れ止め金具_効き良好', 23],
    ['クレーン', 'ブレーキ_効き良好', 25],
    ['クレーン', '異常音異常振動_なし', 27]
  ];
  mappings.forEach(function(m) {
    setPdfRowMapping_(m[0], m[1], m[2]);
  });
  Logger.log(mappings.length + '件のPDF行設定を登録しました。');
}

// 書式_リフトの日付欄（「日」という文字だけが並んでいる状態）に、
// 実際の数字 1〜31 を書き込む（1回限り実行）。
// 「点検日」行と「作業前点検項目」行の両方に同じ数字を入れる
// （generateMonthlyPDFが日付ヘッダーとして読むのは後者）。
// このバージョンは値を書き込むだけで、列の追加・削除は一切行わない。
// 31列分の空き（既存の日列＋手動で追加した列）が既にある前提。
// まだ列が31に足りない場合は、Sheets上で該当行の右クリック→
// 「列を挿入」で先に手動で列を追加してから実行すること。
function setLiftDayNumbers() {
  const sh = getSpreadsheet_().getSheetByName('書式_リフト');
  if (!sh) throw new Error('書式_リフトシートが見つかりません');
  const data = sh.getDataRange().getValues();
  let found = 0;
  ['点検日', '作業前点検項目'].forEach(function(label) {
    for (let r = 0; r < data.length; r++) {
      const c = data[r].indexOf(label);
      if (c > -1) {
        for (let d = 1; d <= 31; d++) {
          sh.getRange(r + 1, c + 1 + d).setValue(d);
        }
        found++;
        break;
      }
    }
  });
  if (found < 2) throw new Error('「点検日」または「作業前点検項目」の行が見つかりませんでした。シート構成を確認してください。');
  Logger.log('書式_リフトの日付欄を1〜31の数字に設定しました（列の追加・削除は行っていません）。');
}
