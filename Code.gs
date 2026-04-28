// ==========================================
// 在庫管理システム (Stock Sync) バックエンドGAS
// ==========================================

const SHEET_NAME_HISTORY = '入出庫履歴';
const SHEET_NAME_STOCK = '在庫一覧';
// 通知を送るメールアドレス（ご自身のGmailアドレスに変更してください。未入力の場合はGAS実行者のアドレスに送られます）
const ALERT_EMAIL = ''; 

// ▼ Webアプリとしてアクセスされたときの処理 (フロントエンドからのデータ受信)
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // 送信されたJSONデータをパース
    const data = JSON.parse(e.postData.contents);
    const itemName = data.itemName;
    const stockIn = Number(data.stockIn) || 0;
    const stockOut = Number(data.stockOut) || 0;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 入出庫履歴シートの更新
    let historySheet = ss.getSheetByName(SHEET_NAME_HISTORY);
    if (!historySheet) {
      historySheet = ss.insertSheet(SHEET_NAME_HISTORY);
      historySheet.appendRow(['タイムスタンプ', '品目名', '入庫数', '出庫数']);
      historySheet.getRange('A1:D1').setBackground('#f3f4f6').setFontWeight('bold');
    }
    
    const date = new Date();
    historySheet.appendRow([date, itemName, stockIn, stockOut]);

    // 2. 在庫一覧シートの更新
    let stockSheet = ss.getSheetByName(SHEET_NAME_STOCK);
    if (!stockSheet) {
      stockSheet = ss.insertSheet(SHEET_NAME_STOCK);
      stockSheet.appendRow(['品目名', '現在庫数']);
      stockSheet.getRange('A1:B1').setBackground('#f3f4f6').setFontWeight('bold');
    }

    const stockData = stockSheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // 品目が既に存在するか検索
    for (let i = 1; i < stockData.length; i++) {
      if (stockData[i][0] === itemName) {
        rowIndex = i + 1; // スプレッドシートの行番号は1始まり
        break;
      }
    }

    let currentStock = 0;
    if (rowIndex !== -1) {
      // 既存品目の場合
      const previousStock = Number(stockSheet.getRange(rowIndex, 2).getValue()) || 0;
      currentStock = previousStock + stockIn - stockOut;
      stockSheet.getRange(rowIndex, 2).setValue(currentStock);
    } else {
      // 新規品目の場合
      currentStock = stockIn - stockOut;
      stockSheet.appendRow([itemName, currentStock]);
    }

    // 3. 在庫が2以下の場合はメール通知
    if (currentStock <= 2) {
      sendAlertEmail(itemName, currentStock);
    }

    // 成功レスポンスを返す
    const result = { 
      status: 'success', 
      message: '在庫を更新しました', 
      currentStock: currentStock 
    };
    output.setContent(JSON.stringify(result));
    return output;

  } catch (err) {
    // エラーレスポンスを返す
    const result = { status: 'error', message: err.message };
    output.setContent(JSON.stringify(result));
    return output;
  }
}

// CORSエラー回避用（OPTIONSリクエストの処理）
function doOptions(e) {
  const output = ContentService.createTextOutput('');
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ▼ メール送信処理
function sendAlertEmail(itemName, stock) {
  const email = ALERT_EMAIL || Session.getActiveUser().getEmail();
  const subject = `【在庫警告】発注が必要です：${itemName}`;
  const body = `在庫管理システム (Stock Sync) からの自動通知です。\n\n` +
               `「${itemName}」の現在庫数が「${stock}」になりました。\n` +
               `規定数（2以下）を下回っているため、至急発注をお願いします。\n\n` +
               `※スプレッドシートをご確認ください：\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
               
  MailApp.sendEmail(email, subject, body);
}

// ▼ 【設定用】毎朝8時に自動チェックするトリガーを作成・実行する関数
// （この関数をGASエディタ上で手動実行するか、トリガー設定画面から設定します）
function checkStockDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName(SHEET_NAME_STOCK);
  if (!stockSheet) return;

  const data = stockSheet.getDataRange().getValues();
  const lowStockItems = [];

  // ヘッダーを飛ばして在庫数をチェック
  for (let i = 1; i < data.length; i++) {
    const itemName = data[i][0];
    const stock = Number(data[i][1]) || 0;
    if (stock <= 2) {
      lowStockItems.push(`・${itemName} (残り: ${stock})`);
    }
  }

  // 該当する品目があればメール通知
  if (lowStockItems.length > 0) {
    const email = ALERT_EMAIL || Session.getActiveUser().getEmail();
    const subject = `【定期警告】発注が必要な商品があります`;
    const body = `在庫管理システム (Stock Sync) からの毎朝の自動チェック通知です。\n\n` +
                 `以下の商品の在庫が規定数（2以下）になっています。\n\n` +
                 `${lowStockItems.join('\n')}\n\n` +
                 `ご確認の上、発注をお願いいたします。\n\n` +
                 `${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
                 
    MailApp.sendEmail(email, subject, body);
  }
}
