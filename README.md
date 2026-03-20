# 跟團訂單自動化系統

使用 **Google Apps Script + Google Sheets + LINE Messaging API** 建立的跟團訂單自動化系統。  
當表單提交後，系統會自動將資料從「拜金表」搬移到「NIJI」彙總分頁，並透過 LINE Bot 即時推播通知管理者。

---

## 專案功能

- 監聽 Google 表單提交事件
- 自動將新訂單資料搬移到指定工作表
- 自動標記原始資料是否已搬移，避免重複處理
- 使用 LINE Messaging API 推播訂單通知
- 提供測試通知功能，方便除錯與驗證

---

## 使用技術

- Google Apps Script
- Google Sheets
- LINE Messaging API
- JavaScript

---

## 系統流程

1. 使用者提交 Google 表單
2. Apps Script 觸發 `onFormSubmit(e)`
3. 系統讀取表單資料並移除時間戳記
4. 檢查該筆資料是否已搬移
5. 將資料新增到 `NIJI` 工作表
6. 在原始工作表標記 `已搬移`
7. 呼叫 LINE Messaging API 發送通知給管理者

---

## 試算表欄位設計

### 原始工作表：`拜金表`

表單提交後的資料來源工作表。

建議欄位如下：

| 欄位 | 說明 |
|---|---|
| 時間戳記 | Google 表單自動產生 |
| 跟團日期 | 跟團日期 |
| 跟團 | 跟團者名稱 |
| 購買商品 | 商品名稱 |
| 金額 | 訂單金額 |
| 完成狀態 | 付款或處理狀態 |
| 已搬移標記 | 程式寫入 `已搬移` |

---

### 彙總工作表：`NIJI`

若不存在，系統會自動建立並加入標題列：

| 欄位 | 說明 |
|---|---|
| 跟團日期 | 跟團日期 |
| 跟團 | 跟團者名稱 |
| 購買商品 | 商品名稱 |
| 金額 | 訂單金額 |
| 完成狀態 | 處理狀態 |
| 特典 | 預留欄位 |

---

## 程式說明

### `onFormSubmit(e)`

當 Google 表單送出時自動觸發。

主要功能：

- 讀取表單資料
- 移除時間戳記
- 檢查是否已搬移
- 搬移資料到 `NIJI`
- 標記原始資料為 `已搬移`
- 發送 LINE 通知

---

### `sendLineNotify(formData)`

使用 LINE Messaging API 推播文字訊息給指定使用者。

通知內容包含：

- 跟團日期
- 跟團者
- 購買商品
- 金額
- 狀態

---

### `testLineNotify()`

測試用函式，不需送出表單即可直接發送一則範例通知。  
方便在開發階段確認 LINE 推播功能是否正常。

---

## 程式碼

```javascript
function onFormSubmit(e) {
  const formData = e.values.slice(1); // 移除時間戳記
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("拜金表");
  const product = formData[2]; // 第三欄為「購買商品」
  const rowIndex = e.range.getRow();
  const lastCol = formSheet.getLastColumn();

  // 檢查是否已搬移
  const movedFlag = formSheet.getRange(rowIndex, lastCol).getValue();
  if (movedFlag === "已搬移") return;

  const newRow = [...formData, '']; // 補空白「特典」欄

  // 搬資料到 NIJI
  let summarySheet = ss.getSheetByName("NIJI");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("NIJI");
    summarySheet.appendRow(['跟團日期', '跟團', '購買商品', '金額', '完成狀態', '特典']);
  }
  summarySheet.appendRow(newRow);

  // 標記為已搬移
  formSheet.getRange(rowIndex, lastCol).setValue("已搬移");

  // 發送 LINE 通知
  sendLineNotify(formData);
}

function sendLineNotify(formData) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  const message = `
📦 有人跟團啦！
🗓️ 日期：${formData[0]}
👤 跟團：${formData[1]}
🛒 商品：
${formData[2]}
💰 金額：${formData[3]}
📌 狀態：${formData[4]}
  `;

  const url = "https://api.line.me/v2/bot/message/push";
  const userId = "U0a2790f5c7aa4091ed16fdf265573acf"; // 替換為你自己的使用者 ID

  const payload = {
    to: userId,
    messages: [
      {
        type: "text",
        text: message
      }
    ]
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${token}`
    },
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(url, options);
}

// 測試通知用
function testLineNotify() {
  const sampleData = ["2025-04-15", "小A", "應援毛巾", "690", "已付款"];
  sendLineNotify(sampleData);
}
