// ==========================================
// 在庫管理システム (Stock Sync) バックエンドGAS
// ※Googleフォームと連携するバージョン
// ==========================================

const SHEET_NAME_STOCK = '在庫一覧';
// 通知先メールアドレス（未入力の場合はこのスクリプトの所有者に送られます）
const ALERT_EMAIL = ''; 

// ▼ 1. Googleフォーム送信時に実行される関数
// （スプレッドシートの「トリガー」設定から、フォーム送信時を条件にしてこの関数を指定してください）
function onFormSubmit(e) {
  // e.namedValues にフォームの回答が入っています
  // ※フォームの質問タイトルが「品目」「入庫数」「出庫数」であると仮定します。
  const responses = e.namedValues;
  const itemName = responses['品目'] ? responses['品目'][0] : '';
  const stockIn = Number(responses['入庫数'] ? responses['入庫数'][0] : 0) || 0;
  const stockOut = Number(responses['出庫数'] ? responses['出庫数'][0] : 0) || 0;

  if (!itemName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let stockSheet = ss.getSheetByName(SHEET_NAME_STOCK);
  
  if (!stockSheet) {
    stockSheet = ss.insertSheet(SHEET_NAME_STOCK);
    stockSheet.appendRow(['品目', '現在庫数', '最終更新日時']);
    stockSheet.getRange('A1:C1').setBackground('#f3f4f6').setFontWeight('bold');
  }

  const stockData = stockSheet.getDataRange().getValues();
  let rowIndex = -1;
  
  // 既存の品目を探す
  for (let i = 1; i < stockData.length; i++) {
    if (stockData[i][0] === itemName) {
      rowIndex = i + 1;
      break;
    }
  }

  const now = new Date();
  let currentStock = 0;
  
  if (rowIndex !== -1) {
    // 既存品目の在庫を更新
    const previousStock = Number(stockSheet.getRange(rowIndex, 2).getValue()) || 0;
    currentStock = previousStock + stockIn - stockOut;
    stockSheet.getRange(rowIndex, 2).setValue(currentStock);
    stockSheet.getRange(rowIndex, 3).setValue(now);
  } else {
    // 新規品目を追加
    currentStock = stockIn - stockOut;
    stockSheet.appendRow([itemName, currentStock, now]);
  }

  // 在庫が2以下の場合はメール通知
  if (currentStock <= 2) {
    sendAlertEmail(itemName, currentStock);
  }
}

// ▼ 2. Webサイト (index.html) からアクセスされた時に在庫データを返す処理
function doGet(e) {
  const output = ContentService.createTextOutput();
  // JSONを返し、Webアプリの実行ユーザー権限でアクセスさせます
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stockSheet = ss.getSheetByName(SHEET_NAME_STOCK);
    
    let stockList = [];
    if (stockSheet) {
      const data = stockSheet.getDataRange().getValues();
      // ヘッダー行を飛ばす
      for (let i = 1; i < data.length; i++) {
        if(data[i][0]) { // 空行は無視
          stockList.push({
            itemName: data[i][0],
            stock: data[i][1]
          });
        }
      }
    }

    const result = {
      status: 'success',
      data: stockList
    };
    
    output.setContent(JSON.stringify(result));
    return output;
  } catch (err) {
    const result = { status: 'error', message: err.message };
    output.setContent(JSON.stringify(result));
    return output;
  }
}

// ▼ メール送信処理
function sendAlertEmail(itemName, stock) {
  const email = ALERT_EMAIL || Session.getActiveUser().getEmail();
  const subject = `【在庫警告】発注が必要です：${itemName}`;
  const body = `在庫管理システムからの自動通知です。\n\n` +
               `「${itemName}」の現在庫数が「${stock}」になりました。\n` +
               `規定数（2以下）を下回っているため、至急発注をお願いします。\n\n` +
               `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
  MailApp.sendEmail(email, subject, body);
}

// ▼ 3. 毎朝の自動チェック用（時間主導型トリガーを設定）
function checkStockDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName(SHEET_NAME_STOCK);
  if (!stockSheet) return;

  const data = stockSheet.getDataRange().getValues();
  const lowStockItems = [];

  for (let i = 1; i < data.length; i++) {
    const itemName = data[i][0];
    const stock = Number(data[i][1]) || 0;
    if (stock <= 2 && itemName) {
      lowStockItems.push(`・${itemName} (残り: ${stock})`);
    }
  }

  if (lowStockItems.length > 0) {
    const email = ALERT_EMAIL || Session.getActiveUser().getEmail();
    const subject = `【定期警告】発注が必要な商品があります`;
    const body = `毎朝の自動チェック通知です。\n\n` +
                 `以下の商品の在庫が規定数（2以下）になっています。\n\n` +
                 `${lowStockItems.join('\n')}\n\n` +
                 `発注をお願いします。\n\n` +
                 `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
    MailApp.sendEmail(email, subject, body);
  }
}
