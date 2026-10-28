/**
 * スマート持ち物リスト Web API (Google Apps Script)
 * スプレッドシートと連携し、GitHub Pages (フロントエンド) と高速・安全に同期するためのバックエンドスクリプトです。
 *
 * 【0ベース再設計の改善点】
 * 1. LockService による排他制御で、連続操作（並び替え連打・連続追加/編集/削除）時のデータ競合・先祖返りを完全防止
 * 2. 更新系APIでの Read-after-Write (getDisplayValues) 呼び出しを廃止し、GAS特有の再計算フラッシュ遅延を解消
 * 3. 単一アイテムの追加・更新（Upsert）はA列（ID列）のみを検索して対象1行のみをピンポイント更新する超高速方式
 * 4. GETクエリパラメータ通信とPOSTボディ通信のハイブリッド対応（302リダイレクト時のCORS/通信失敗を回避）
 * 5. 重複ID行の自動マージ・クリーンアップと、並び替え時の末尾ゾンビ行残留バグの完全修正
 */

const COL_COUNT = 11; // A列(ID) 〜 K列(備考)

/**
 * 外部Webアプリ（GitHub Pages等）からのAPI処理共通ディスパッチャー
 */
function handleApiAction(action, payload) {
  payload = payload || {};
  action = action || 'get';

  if (action === 'get' || action === 'getInitialData') {
    return getItemsData();
  }

  if (action === 'ping') {
    return { status: 'ok', message: 'Smart Belongings API is active.' };
  }

  // 更新系リクエストは LockService で排他制御し、同時リクエストによるデータの先祖返り・競合を完全防止する
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (action === 'upsert' || action === 'add' || action === 'update') {
      return upsertItemInSheet(payload.item);
    } else if (action === 'delete') {
      return removeItemFromSheet(payload.id);
    } else if (action === 'reorder') {
      return reorderItemsInSheet(payload.orderedIds);
    } else {
      throw new Error('未定義のアクションです: ' + action);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * ① GETリクエストのハンドラー（データの取得・軽量APIアクション実行・稼働確認）
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'get';
    let payload = {};
    if (e && e.parameter) {
      if (e.parameter.payload) {
        try {
          payload = JSON.parse(e.parameter.payload);
        } catch (jsonErr) {
          payload = {};
        }
      } else {
        payload = { ...e.parameter };
      }
    }

    const result = handleApiAction(action, payload);
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ② POSTリクエストのハンドラー（追加・更新・削除・並び替え・取得）
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('リクエストボディが空です');
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'get';
    const result = handleApiAction(action, payload);

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --------------------------------------------------------
// 以下は、スプレッドシートのデータを高速・安全に読み書きする処理です
// --------------------------------------------------------

/**
 * 対象のシートを取得（先頭シートまたはアクティブシート）
 */
function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()[0] || ss.getActiveSheet();
}

/**
 * 買った時期（date）の文字列を「YYYY/MM」形式に正規化
 * 例: "2026/5" -> "2026/05", "2026/09/01" -> "2026/09", "2026-05" -> "2026/05"
 */
function normalizeDateString(rawDate) {
  if (!rawDate) return '';
  const str = String(rawDate).trim();
  if (!str) return '';

  const match = str.match(/^(\d{4})[\/\-](\d{1,2})(?:[\/\-]\d{1,2})?$/);
  if (match) {
    const year = match[1];
    const month = String(parseInt(match[2], 10)).padStart(2, '0');
    return `${year}/${month}`;
  }
  return str;
}

/**
 * 1行分の配列からフロントエンド用アイテムオブジェクトを生成
 */
function parseRowToItem(row) {
  return {
    id: String(row[0] || '').trim(),
    categoryMain: String(row[1] || '').trim(),
    categorySub: String(row[2] || '').trim(),
    categoryMinor: String(row[3] || '').trim(),
    name: String(row[4] || '').trim(),
    color: String(row[5] || '').trim(),
    brand: String(row[6] || '').trim(),
    size: String(row[7] || '').trim(),
    date: normalizeDateString(row[8]),
    price: String(row[9] || '').trim(),
    notes: String(row[10] || '').trim()
  };
}

/**
 * アイテムオブジェクトからスプレッドシート書き込み用の1行配列（11列）を生成
 */
function buildRowFromItem(item, targetId) {
  const normalizedDate = normalizeDateString(item.date);
  return [
    String(targetId || item.id || '').trim(),
    String(item.categoryMain || '').trim(),
    String(item.categorySub || '').trim(),
    String(item.categoryMinor || '').trim(),
    String(item.name || '').trim(),
    String(item.color || '').trim(),
    String(item.brand || '').trim(),
    String(item.size || '').trim(),
    normalizedDate,
    String(item.price !== undefined && item.price !== null ? item.price : '').trim(),
    String(item.notes || '').trim()
  ];
}

/**
 * 2つの同一IDアイテム（重複行）のデータをマージする（primaryをベースに、空欄項目があればsecondaryで補完）
 */
function mergeDuplicateRowArrays(primaryRow, secondaryRow) {
  const merged = primaryRow.slice(0, COL_COUNT);
  while (merged.length < COL_COUNT) merged.push('');
  for (let c = 0; c < COL_COUNT; c++) {
    const pVal = String(merged[c] || '').trim();
    const sVal = String(secondaryRow[c] || '').trim();
    if (!pVal && sVal) {
      merged[c] = secondaryRow[c];
    }
  }
  merged[8] = normalizeDateString(merged[8]);
  return merged;
}

/**
 * スプレッドシートから全アイテムを純粋読み取りで超高速取得（1回のgetDisplayValuesのみ）
 */
function getItemsData() {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COL_COUNT).getDisplayValues();
  const items = [];
  const idToIndexMap = new Map();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;

    if (idToIndexMap.has(id)) {
      // 万一重複ID行があった場合は上位アイテムの空欄項目を補完マージ
      const idx = idToIndexMap.get(id);
      const existing = items[idx];
      const candidate = parseRowToItem(row);
      Object.keys(existing).forEach(k => {
        if (!existing[k] && candidate[k]) existing[k] = candidate[k];
      });
    } else {
      idToIndexMap.set(id, items.length);
      items.push(parseRowToItem(row));
    }
  }

  return items;
}

/**
 * アイテムの追加・更新（Upsert）をピンポイント1行書き込みで超高速実行（Read-after-Writeフラッシュなし）
 */
function upsertItemInSheet(item) {
  if (!item) {
    return { ok: false, error: 'アイテムデータが空です' };
  }

  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  const targetId = String(item.id || Utilities.getUuid().split('-')[0]).trim();
  const newRowData = buildRowFromItem(item, targetId);

  if (lastRow < 2) {
    sheet.getRange(2, 1, 1, COL_COUNT).setValues([newRowData]);
    return { ok: true, id: targetId, mode: 'created' };
  }

  // A列（ID列）のみを高速取得して該当行を特定
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const matchedRows = [];
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === targetId) {
      matchedRows.push(i + 2);
    }
  }

  if (matchedRows.length > 0) {
    // 先頭の一致行を1回のsetValuesでピンポイント上書き更新
    const primaryRow = matchedRows[0];
    sheet.getRange(primaryRow, 1, 1, COL_COUNT).setValues([newRowData]);

    // 万が一同一IDの重複行が2行目以降に存在していた場合は下から削除
    for (let j = matchedRows.length - 1; j >= 1; j--) {
      sheet.deleteRow(matchedRows[j]);
    }
    return { ok: true, id: targetId, mode: 'updated' };
  } else {
    // 新規行として末尾の1行にピンポイント追加
    const targetRow = lastRow + 1;
    sheet.getRange(targetRow, 1, 1, COL_COUNT).setValues([newRowData]);
    return { ok: true, id: targetId, mode: 'created' };
  }
}

/**
 * アイテムの削除（同一IDが複数行あった場合もすべて確実に削除し、Read-after-Writeを排除）
 */
function removeItemFromSheet(id) {
  const targetId = String(id || '').trim();
  if (!targetId) return { ok: false };

  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, id: targetId, deletedCount: 0 };

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  let deletedCount = 0;

  // 下の行から順番に削除することで行番号ズレを防ぐ
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0] || '').trim() === targetId) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }

  return {
    ok: true,
    id: targetId,
    deletedCount: deletedCount
  };
}

/**
 * アイテムの並び順（行順）の一括更新（重複ID排除＆余剰末尾行の完全クリア対応）
 */
function reorderItemsInSheet(orderedIds) {
  if (!Array.isArray(orderedIds) || orderedIds.length === 0) {
    return { ok: true, count: 0 };
  }

  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const dataRows = sheet.getRange(2, 1, lastRow - 1, COL_COUNT).getDisplayValues();
  const rowMap = new Map();

  // 上から順に走査し、同一IDの重複行があれば先頭（上位）のデータを優先しつつ空欄をマージする
  dataRows.forEach(row => {
    const id = String(row[0] || '').trim();
    if (!id) return;
    const row11 = row.slice(0, COL_COUNT);
    row11[8] = normalizeDateString(row11[8]);

    if (rowMap.has(id)) {
      rowMap.set(id, mergeDuplicateRowArrays(rowMap.get(id), row11));
    } else {
      rowMap.set(id, row11);
    }
  });

  const newRows = [];
  orderedIds.forEach(id => {
    const key = String(id || '').trim();
    if (rowMap.has(key)) {
      newRows.push(rowMap.get(key));
      rowMap.delete(key);
    }
  });

  // 万が一 orderedIds に含まれなかった残りの行があれば末尾に追加
  rowMap.forEach(row => {
    newRows.push(row);
  });

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, COL_COUNT).setValues(newRows);
  }

  // 重複除去などで行数が減った場合、末尾に残った古い行（ゾンビ行）を確実に消去する
  const originalCount = lastRow - 1;
  if (originalCount > newRows.length) {
    sheet.getRange(2 + newRows.length, 1, originalCount - newRows.length, COL_COUNT).clearContent();
  }

  return {
    ok: true,
    count: newRows.length
  };
}

// 後方互換用関数ラッパー（GAS内部呼び出し等に対応）
function addNewItem(item) { return upsertItemInSheet(item); }
function updateItem(item) { return upsertItemInSheet(item); }
function removeItem(id) { return removeItemFromSheet(id); }
function reorderItems(orderedIds) { return reorderItemsInSheet(orderedIds); }
