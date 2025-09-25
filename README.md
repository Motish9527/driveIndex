# Google Drive 索引器 - 開發者文檔

> 基於 Google Apps Script 的 Google Drive 檔案索引系統，提供增量更新、智能篩選、性能優化等功能。

## 📋 目錄

- [架構概覽](#架構概覽)
- [核心模組](#核心模組)
- [性能優化](#性能優化)
- [開發指南](#開發指南)
- [故障排除](#故障排除)
- [資料結構](#資料結構)
- [部署指南](#部署指南)

---


## 🏗️ 架構概覽

### 系統架構
```
Google Drive → 掃描引擎 → 快取層 → 索引層 → UI 層
     ↓           ↓         ↓        ↓       ↓
  檔案系統    traverse_   Cache   Index  篩選器
```

### 核心流程
1. **掃描階段**: `traverse_()` 深度優先遍歷 Drive 結構
2. **快取階段**: 建立 `Cache` 工作表作為資料持久層
3. **索引階段**: 生成 `Index` 工作表供使用者操作
4. **UI 階段**: 提供篩選、搜尋、下拉選單等互動功能

---

## 📦 核心模組

### `global.js` - 全域配置與工具___________________
```javascript
// 核心常數
const INDEX_SHEET_NAME = 'Index';
const CACHE_SHEET_NAME = 'Cache';

// 主要函數
function onOpen()           // 初始化選單和 UI
function onEdit(e)          // 處理使用者互動
function getOrCreateSheet_() // 工作表管理
```

**職責**:
- 全域常數定義
- 選單系統初始化
- 事件處理器 (`onEdit`, `onOpen`)
- 共用工具函數

### `googleDriveIndex.js` - 索引引擎_______________________
```javascript
// 主要流程
function fullRebuild()           // 全量重建
function incrementalUpdate()     // 增量更新
function rebuildIndexFromCache() // 從快取重建

// 核心工具
function traverse_()             // 檔案系統遍歷
function postFormatIndexSheet_() // 格式化處理
```

**職責**:
- Drive 檔案系統掃描
- 增量更新邏輯
- 資料格式化與排序
- 快取管理

### `UI_selection.js` - 使用者介面______________________
```javascript
// 下拉選單管理
function updateDropdownB3_()     // Drive 資料夾選擇
function updateDropdownC3_()     // 層級篩選
function updateDropdownE3_()     // 檔案類型篩選

// 篩選邏輯
function applyFilterByC3_()      // 層級篩選
function applyFilterByE3_()      // 類型篩選
```

**職責**:
- 下拉選單動態生成
- 篩選器實作
- 使用者互動處理

### `KWDetection.js` - 搜尋功能__________________________
```javascript
// 搜尋系統
function setupF3Search_()        // 初始化搜尋
function performSearch_()        // 執行搜尋
function clearSearch_()          // 清除搜尋
```

**職責**:
- Google 風格搜尋實作
- 關鍵字匹配
- 搜尋結果顯示

---

## ⏰ 自動觸發器系統

### 觸發器配置
系統已配置自動觸發器，無需手動設定：

- **更新時間**: 台北時間每日 8:00 和 14:00
- **觸發函數**: `incrementalUpdate()`
- **自動安裝**: 系統啟動時自動檢查並安裝觸發器

### 觸發器管理
```javascript
// 重新安裝觸發器
installAutoTriggers_();

// 移除所有觸發器
removeTriggers_();

// 檢查觸發器狀態
initializeAutoTriggers();
```

### 時區設定
- 系統使用 UTC 時間設定觸發器
- UTC 0:00 = 台北時間 8:00
- UTC 6:00 = 台北時間 14:00

---

## 🔧 API 參考

### 核心函數

#### `incrementalUpdate()`
執行增量更新，只處理變更的檔案。

**性能特點**:
- 智能檢測變更
- 批量操作優化
- 智能格式化（只在有變更時執行）

```javascript
// 使用範例
incrementalUpdate(); // 增量更新
```

#### `traverse_(folder, path, parentId, onFolder, onFile)`
遞迴遍歷 Google Drive 資料夾結構。

**參數**:
- `folder`: 當前資料夾物件
- `path`: 路徑陣列
- `parentId`: 父層 ID
- `onFolder`: 資料夾回調函數
- `onFile`: 檔案回調函數

#### `postFormatIndexSheet_(sheet)`
格式化索引工作表，包含排序、合併、著色等。

**處理項目**:
- 依 Full Path 排序
- 層級欄位合併
- Folder 行著色
- 篩選器建立

### 配置常數

```javascript
// 工作表配置
const INDEX_SHEET_NAME = 'Index';
const CACHE_SHEET_NAME = 'Cache';
const SOURCE_START_ROW = 5;

// 欄位配置
const INDEX_HEADERS = [
  'ID*', '資料夾層級1', '資料夾層級2', '資料夾層級3', '資料夾層級4',
  '檔案名稱', '檔案類型', '連結', 'TAG', 'Full Path'
];

// 支援的檔案類型
const ALLOWED_MIME_SET = {
  'application/vnd.google-apps.document': true,
  'application/vnd.google-apps.spreadsheet': true,
  'application/vnd.google-apps.presentation': true,
  'application/pdf': true,
  // ... 更多類型
};
```

---

## ⚡ 性能優化

### 優化策略

#### 1. 智能格式化
```javascript
// 只在有變更時執行格式化
if (hasChanges) {
  postFormatIndexSheet_(indexSheet);
}
```

#### 2. 批量操作
```javascript
// 批量寫入而非逐行操作
writeBatchedRows_(sheet, updates, columnCount);
```

#### 3. 快取優化
- 使用 `Cache` 工作表作為持久層
- 增量比對減少重複掃描
- ID 映射加速查找

### 性能監控

#### 內建統計
```javascript
function checkPerformanceStats() {
  // 顯示資料量統計
  // 預估處理時間
  // 當前模式狀態
}
```

#### 時間追蹤
```javascript
const startTime = new Date().getTime();
// ... 處理邏輯
const endTime = new Date().getTime();
console.log(`處理耗時: ${endTime - startTime}ms`);
```

---

## 🛠️ 開發指南

### 新增檔案類型支援

1. **更新 MIME 集合**:
```javascript
// 在 global.js 中
const ALLOWED_MIME_SET = {
  // 現有類型...
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': true,
};
```

2. **添加友善名稱**:
```javascript
function toFriendlyType_(mime) {
  const typeMap = {
    // 現有映射...
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word 文件',
  };
  return typeMap[mime] || mime;
}
```

### 擴展篩選功能

1. **新增篩選欄位**:
```javascript
// 在 UI_selection.js 中
function updateDropdownG3_() {
  // 實作新的下拉選單
}

function applyFilterByG3_(selected) {
  // 實作新的篩選邏輯
}
```

2. **註冊事件處理**:
```javascript
// 在 global.js 的 onEdit 中
if (range.getColumn() === 7) { // G3
  applyFilterByG3_(selected);
}
```

### 自定義格式化

```javascript
function customFormatIndexSheet_(sheet) {
  // 自定義排序邏輯
  sheet.getRange(5, 1, lastRow - 4, lastCol)
    .sort([{ column: CUSTOM_SORT_COL, ascending: true }]);
  
  // 自定義合併邏輯
  mergeCustomColumns_(sheet);
  
  // 自定義著色邏輯
  applyCustomColors_(sheet);
}
```

---

## 🔍 故障排除

### 常見問題

#### 1. 增量更新耗時過長
**原因**: 大量資料 + 格式化操作
**解決方案**:
```javascript
// 系統已內建智能格式化，只在有變更時執行格式化
// 無需額外配置
```

#### 2. 篩選器無響應
**原因**: 函數名稱衝突或事件處理問題
**檢查項目**:
- 確認 `onEdit` 事件正確註冊
- 檢查篩選函數名稱唯一性
- 驗證工作表範圍設定

#### 3. 下拉選單不更新
**原因**: 資料來源變更但選單未刷新
**解決方案**:
```javascript
// 手動更新選單
updateDropdownMenu();
```

#### 4. 合併儲存格衝突
**原因**: 手動操作與自動格式化衝突
**解決方案**:
```javascript
// 先取消合併
preUnmergeForUpdate_(sheet);

// 執行操作
// ...

// 重新格式化
postFormatIndexSheet_(sheet);
```

### 除錯工具

#### 1. 日誌輸出
```javascript
console.log('處理階段:', stage);
console.log('資料量:', dataCount);
console.log('耗時:', duration);
```

#### 2. 狀態檢查
```javascript
function debugSystemState() {
  console.log('工作表狀態:', getSheetStatus());
  console.log('快取狀態:', getCacheStatus());
}
```

#### 3. 性能分析
```javascript
function analyzePerformance() {
  const stats = {
    cacheRows: cacheSheet.getLastRow(),
    indexRows: indexSheet.getLastRow(),
    estimatedTime: calculateEstimatedTime()
  };
  console.log('性能統計:', stats);
}
```

---

## 📊 資料結構

### Index 工作表結構
| 欄位 | 說明 | 類型 | 範例 |
|------|------|------|------|
| A (ID*) | 檔案唯一識別碼 | String | 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms |
| B-E (層級1-4) | 資料夾層級路徑 | String | 專案/子專案/文件 |
| F (檔案名稱) | 檔案顯示名稱 | String | 報告.pdf |
| G (檔案類型) | 友善類型名稱 | String | PDF |
| H (連結) | Google Drive 連結 | URL | https://drive.google.com/... |
| I (TAG) | 使用者標籤 | String | 重要 |
| J (Full Path) | 完整路徑 | String | 專案/子專案/文件/報告.pdf |

### Cache 工作表結構
| 欄位 | 說明 | 用途 |
|-----|------|------|
| ID  | 檔案唯一識別碼 | 增量比對 |
| 類型 | 檔案類型 | 分類統計 |
| 名稱 | 檔案名稱 | 顯示用 |
| MIME | MIME 類型 | 類型判斷 |
| 連結 | Drive 連結 | 快速存取 |
| 父層ID | 父資料夾 ID | 結構重建 |
| 層級1-4 | 路徑層級 | 篩選用 |
| TAG | 使用者標籤 | 自定義分類 |
| 快取時間 | 最後更新時間 | 變更檢測 |
| FullPath | 完整路徑 | 排序用 |

---


## 🔄 更新日誌

### v2.1.0 (2024-12)
- ✅ 實作自動化觸發器系統
- ✅ 設定台北時間每日 8:00 和 14:00 自動更新
- ✅ 移除手動觸發器安裝步驟
- ✅ 優化觸發器管理功能

### v2.0.0 (2024-12)
- ✅ 新增性能優化功能
- ✅ 實作智能格式化
- ✅ 添加性能監控工具
- ✅ 改善篩選器響應性
- ✅ 優化搜尋功能

### v1.0.0 (2024-11)
- ✅ 基礎索引功能
- ✅ 增量更新機制
- ✅ 篩選和搜尋功能
- ✅ 下拉選單系統

---

*最後更新: 2024年12月*