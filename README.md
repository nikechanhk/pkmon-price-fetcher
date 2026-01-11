## Step 1: 開啟 Google Spreadsheet

1. 開啟你想使用的 Google Spreadsheet
2. 點選上方選單 **Extensions** → **Apps Script**
3. 會開啟一個新的 Apps Script 編輯器視窗

## Step 2: 建立 Apps Script 函式

在 Apps Script 編輯器中，刪除預設的 `myFunction()`，貼上以下程式碼：

```javascript
/**
 * Fetch latest PSA10 transaction price for a given SNKR ID
 * @param {string|number} snkrId - The SNKR ID to fetch price for
 * @return {number|string} Latest PSA10 transaction price in HKD, or error message
 * @customfunction
 */
function GET_PSA10_PRICE(snkrId) {
  // Validate input
  if (!snkrId || snkrId === '') {
    return 'Error: SNKR ID is required';
  }
  
  // Convert to string and trim whitespace
  const cleanSnkrId = String(snkrId).trim();
  
  // Your API endpoint URL
  const apiUrl = 'https://pkmon-tracking.zeabur.app/api/fetch-latest-psa10-transaction-price';
  const url = `${apiUrl}?snkrId=${encodeURIComponent(cleanSnkrId)}`;
  
  try {
    // Log for debugging
    Logger.log('Fetching URL: ' + url);
    
    // Make HTTP GET request
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'Accept': 'application/json',
        'User-Agent': 'GoogleSheets/1.0'
      },
      validateHttpsCertificates: true
    });
    
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('Status Code: ' + statusCode);
    Logger.log('Response: ' + responseText);
    
    // Check if response is empty
    if (!responseText || responseText.trim() === '') {
      return 'Error: Empty response from server';
    }
    
    // Parse JSON response
    let data;
    try {
      data = JSON.parse(responseText);
    } catch (parseError) {
      Logger.log('Parse Error: ' + parseError.message);
      return 'Error: Invalid JSON response - ' + responseText.substring(0, 50);
    }
    
    // Handle different response scenarios
    if (statusCode === 200) {
      if (data.success === true) {
        // Return the price, or 0 if no history found
        const price = data.latestPriceHKD;
        if (price !== undefined && price !== null) {
          return Number(price);
        } else {
          return 0;
        }
      } else {
        // API returned success: false
        return 'Error: ' + (data.error || 'API returned success: false');
      }
    } else if (statusCode === 400) {
      return 'Error: Bad request - ' + (data.error || 'Invalid snkrId');
    } else if (statusCode === 500) {
      return 'Error: Server error - ' + (data.error || 'Internal server error');
    } else {
      return 'Error: HTTP ' + statusCode + ' - ' + (data.error || responseText.substring(0, 50));
    }
    
  } catch (error) {
    // Handle network or other errors
    Logger.log('Exception: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return 'Error: ' + error.message + ' (Check Apps Script logs)';
  }
}
```

## Step 3: 儲存並授權

1. 點選上方的 **💾 儲存** 圖示
2. 為專案命名（例如：`SNKR Price Fetcher`）
3. 在編輯器中點選 **▶️ Run** 按鈕
4. 會彈出授權視窗，點選 **Review Permissions**
5. 選擇你的 Google 帳號
6. 點選 **Advanced** → **Go to [專案名稱] (unsafe)**
7. 點選 **Allow**
8. 回到 Spreadsheet，函式現在應該可以正常運作

## Step 4: 在 Spreadsheet 中使用

回到你的 Google Spreadsheet，現在可以使用以下兩個自訂函式：

### 方法 1：取得純數字價格

在任意儲存格輸入, 參數為snkr ID：

```
=GET_PSA10_PRICE(115274)
```

或參照另一個儲存格的 SNKR ID：

```
=GET_PSA10_PRICE(A2)
```

**輸出範例**：`1730`


