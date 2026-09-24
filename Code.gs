/**
 * スマート持ち物リスト Web API (Google Apps Script)
 * スプレッドシートと連携し、GitHub Pages (フロントエンド) と通信するためのバックエンドスクリプトです。
 */

// ① GETリクエストのハンドラー（データの取得・稼働確認）
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'get';
    let result;

    if (action === 'get') {
      result = getItemsData();
    } else {
      result = { status: 'ok', message: 'Smart Belongings API is active.' };
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ② POSTリクエストのハンドラー（追加・更新・削除・取得）
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('リクエストボディが空です');
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result;

    if (action === 'get') {
      result = getItemsData();
    } else if (action === 'add') {
      result = addNewItem(payload.item);
    } else if (action === 'update') {
      result = updateItem(payload.item);
    } else if (action === 'delete') {
      result = removeItem(payload.id);
    } else {
      throw new Error('未定義のアクションです: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --------------------------------------------------------
// 以下は、スプレッドシートのデータを読み書きする処理です
// --------------------------------------------------------

// スプレッドシートから全アイテムを取得
function getItemsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues(); // すべて文字列として取得
  const items = [];

  // 1行目はヘッダーなので2行目から読み込み
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // IDがない空行はスキップ

    items.push({
      id: data[i][0],           // A列: ID
      categoryMain: data[i][1], // B列: 項目（大項目）
      categorySub: data[i][2],  // C列: 分類（中項目）
      categoryMinor: data[i][3],// D列: 小項目
      name: data[i][4],         // E列: 商品名
      color: data[i][5],        // F列: 色
      brand: data[i][6],        // G列: ブランド
      size: data[i][7],         // H列: サイズ
      date: data[i][8],         // I列: 買った時期
      price: data[i][9],        // J列: 金額
      notes: data[i][10]        // K列: 備考
    });
  }
  return items;
}

// 新規アイテムの追加
function addNewItem(item) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rowData = [
    item.id,
    item.categoryMain,
    item.categorySub,
    item.categoryMinor,
    item.name,
    item.color,
    item.brand,
    item.size,
    item.date,
    item.price,
    item.notes
  ];

  sheet.appendRow(rowData);
  return getItemsData();
}

// アイテムの削除
function removeItem(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getItemsData();
}

// アイテムの更新
function updateItem(item) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == item.id) {
      const rowData = [
        item.id,
        item.categoryMain,
        item.categorySub,
        item.categoryMinor,
        item.name,
        item.color,
        item.brand,
        item.size,
        item.date,
        item.price,
        item.notes
      ];
      sheet.getRange(i + 1, 1, 1, 11).setValues([rowData]);
      break;
    }
  }
  return getItemsData();
}
