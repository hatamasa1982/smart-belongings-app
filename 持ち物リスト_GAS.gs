function doGet(e) {
  // ブラウザ（スマホやPC）からアクセスされたら、「持ち物リスト_index」の画面を表示します
  return HtmlService.createHtmlOutputFromFile('持ち物リスト_index')
    .setTitle('私の持ち物リスト') // ブラウザのタブの名前
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // スマホできれいに表示するおまじない
}

// --------------------------------------------------------
// 以下は、画面（HTML）の裏側で動いてデータを出し入れするプログラムです
// --------------------------------------------------------

// ① データをスプレッドシートから読み取って返す
function getItemsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues(); // getValues()から変更：すべて文字列として取得
  const items = [];
  
  // 1行目は見出しなので、2行目（i=1）からデータを取る
  for (let i = 1; i < data.length; i++) {
    // データが空の行（IDがない行）は飛ばす
    if (!data[i][0]) continue;

    // スプレッドシートの列の順番に合わせてデータを取得（A列が[0], J列が[9]）
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
  return items; // 画面側にデータを渡す
}

// ② データをスプレッドシートに追加する
function addNewItem(item) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // スプレッドシートの列の順番（A列〜J列）に合わせて、1行分のデータを作る
  const rowData = [
    item.id,           // A列: ID
    item.categoryMain, // B列: 項目
    item.categorySub,  // C列: 分類
    item.categoryMinor,// D列: 小項目
    item.name,         // E列: 商品名
    item.color,        // F列: 色
    item.brand,        // G列: ブランド
    item.size,         // H列: サイズ
    item.date,         // I列: 買った時期
    item.price,        // J列: 金額
    item.notes         // K列: 備考
  ];
  
  sheet.appendRow(rowData); // スプレッドシートの一番下に1行追加
  
  return getItemsData(); // 最新のデータをまとめて返す
}

// ③ データをスプレッドシートから削除する
function removeItem(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues(); // ここはID比較なので元のままでもOKですが、安全のため
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1); // 該当する行を削除
      break;
    }
  }
  
  return getItemsData(); // 最新のデータをまとめて返す
}

// ④ データをスプレッドシートで更新（上書き）する
function updateItem(item) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == item.id) {
      // 該当する行のデータをすべて上書きする
      const rowData = [
        item.id,           // A列: ID
        item.categoryMain, // B列: 項目
        item.categorySub,  // C列: 分類
        item.categoryMinor,// D列: 小項目
        item.name,         // E列: 商品名
        item.color,        // F列: 色
        item.brand,        // G列: ブランド
        item.size,         // H列: サイズ
        item.date,         // I列: 買った時期
        item.price,        // J列: 金額
        item.notes         // K列: 備考
      ];
      // 該当の行 (i + 1行目) に、1行分 (1行×11列) のデータをセットする
      sheet.getRange(i + 1, 1, 1, 11).setValues([rowData]);
      break;
    }
  }
  
  return getItemsData(); // 最新のデータをまとめて返す
}
