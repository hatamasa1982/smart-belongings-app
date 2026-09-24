function doGet(e) {
  // スプレッドシートのデータを読み込む
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const items = [];
  // 1行目は見出し（ID, 名前, カテゴリ）を想定するため、2行目（i=1）からループする
  for (let i = 1; i < data.length; i++) {
    items.push({
      id: data[i][0],
      name: data[i][1],
      category: data[i][2]
    });
  }
  
  // JSON形式でデータを返す（HTML側で受け取るため）
  return ContentService.createTextOutput(JSON.stringify(items))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // HTML側から送られてきたデータを受け取る
  const params = JSON.parse(e.postData.contents);
  
  if (params.action === 'add') {
    // 追加の処理
    sheet.appendRow([params.item.id, params.item.name, params.item.category]);
  } else if (params.action === 'delete') {
    // 削除の処理：該当するIDの行を探して削除する
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == params.id) {
        sheet.deleteRow(i + 1); // スプレッドシートは1行目から始まるため +1
        break;
      }
    }
  }
  
  // 成功したことを返す
  return ContentService.createTextOutput(JSON.stringify({status: 'success'}))
    .setMimeType(ContentService.MimeType.JSON);
}
