/***** 全域配置與常數 *****/

// 工作表名稱配置
const INDEX_SHEET_NAME = 'Index';
const CACHE_SHEET_NAME = 'Cache';
const DRIVEID_SHEET_NAME = 'DriveID';
const KWDATA_SHEET_NAME = 'KWData';

// 日期格式
const DATE_FORMAT = 'yyyy-MM-dd HH:mm:ss';

// 欄位座標配置（Index 表）
const LEVEL_START_COL = 2; // B：資料夾層級1
const LEVEL_END_COL = 5;   // E：資料夾層級4
const TYPE_COL = 7;        // G：檔案類型
const FORMAT_START_COL = 6; // F 欄（檔案名稱）起算
const FULL_PATH_COL = 10;   // J：Full Path
const LIGHT_GREEN3 = '#fff2cc'; // 淺黃色


// Google Drive 根資料夾 ID
const ROOT_FOLDER_ID = '#######';

// 允許的檔案類型 MIME 設定
const ALLOWED_MIME_SET = (() => {
  const set = {};
  [
    'application/vnd.google-apps.spreadsheet',
    'application/vnd.google-apps.document',
    'application/vnd.google-apps.presentation',
    'application/pdf',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.ms-powerpoint',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  ].forEach(m => (set[m] = true));
  return set;
})();

// Index 表頭
const INDEX_HEADERS = [
  'ID*',               // A（隱藏）
  '資料夾層級1',       // B
  '資料夾層級2',       // C
  '資料夾層級3',       // D
  '資料夾層級4',       // E
  '檔案名稱',          // F
  '檔案類型',          // G
  '連結',              // H
  'TAG',               // I（手動填也可）
  'Full Path'          // J（不限層級）
];

// Cache 表頭
const CACHE_HEADERS = [
  'ID', '類型', '名稱', 'MIME', '連結', '父層ID',
  '層級1', '層級2', '層級3', '層級4', 'TAG', '快取時間', 'FullPath'
];

// Selection.js 篩選器相關配置
const SELECTION_SHEET_NAME = 'Index'; // 統一使用 Index 工作表名稱
const SOURCE_START_ROW = 5;           // 來源資料從第 5 列開始（A5 起）
const HEADER_ROW = SOURCE_START_ROW - 1; // 篩選器的表頭列（第 4 列）

/***** 全域共用函式 *****/

// 統一的 onOpen 函式，整合所有模組的初始化
function onOpen() {
  // 建立 Drive 索引選單
  SpreadsheetApp.getUi()
    .createMenu('Drive 索引')
    .addItem('增量更新（掃描新增/更名/搬家/刪除）', 'incrementalUpdate')
    .addItem('重建索引（全掃描）', 'fullRebuild')
    .addItem('從快取重建索引', 'rebuildIndexFromCache')
    .addItem('排版(自動合併儲存格)', 'mergeCells')
    .addItem('更新下拉式選單', 'updateDropdownMenu')
    .addSeparator()
    .addItem('清除搜尋結果', 'clearSearch_')
    .addSeparator()
    .addItem('檢查性能統計', 'checkPerformanceStats')
    .addSeparator()
    .addToUi();
    
  // 自動安裝觸發器（如果尚未安裝）
  try {
    initializeAutoTriggers();
  } catch (e) {
    console.log('安裝自動觸發器時發生錯誤:', e);
  }
    
  // 初始化下拉式選單
  try {
    updateDropdownB3_(); // 新增：DriveID 選擇下拉式選單
    updateDropdownC3_();
    updateDropdownE3_(); // 關鍵字篩選
    updateDropdownG3_(); // 檔案類型篩選移到 G3
    setupF3Search_(); // F3 搜尋系統
  } catch (e) {
    console.log('下拉式選單初始化時發生錯誤（可能 selection.js 未載入）:', e);
  }
}

// 統一的 onEdit 函式，處理下拉式選單的篩選功能
function onEdit(e) {
  try {
    const range = e.range;
    const sh = range.getSheet();
    if (sh.getName() !== INDEX_SHEET_NAME) return;

    // 只在 B3、C3、E3、F3 或 G3 被改動時觸發
    if (range.getRow() === 3) {
      const selected = e.value || ''; // 當使用者清空時 e.value 會是 undefined
      
      if (range.getColumn() === 2) {
        // B3 被改動：選擇雲端資料夾
        console.log('雲端資料夾已變更為:', selected);
        // 可以在這裡添加額外的處理邏輯，例如清除現有的索引資料
      } else if (range.getColumn() === 3) {
        // C3 被改動：篩選資料夾層級1（B欄）
        console.log('C3 被改動，選擇值:', selected);
        applyFilterByC3_(selected);
      } else if (range.getColumn() === 5) {
        // E3 被改動：關鍵字篩選（J欄 Full Path）
        console.log('E3 被改動，選擇值:', selected);
        applyFilterByE3_(selected);
      } else if (range.getColumn() === 6) {
        // F3 被改動：Google 風格搜尋
        console.log('F3 被改動，選擇值:', selected);
        handleF3Search_(selected);
      } else if (range.getColumn() === 7) {
        // G3 被改動：篩選檔案類型（G欄）
        console.log('G3 被改動，選擇值:', selected);
        applyFilterByG3_(selected);
      }
    }
  } catch (err) {
    console.error('onEdit 處理時發生錯誤:', err);
  }
}

// 取得或建立工作表的通用函式
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// 取得根資料夾的函式（支援動態選擇）
function getRootFolder_() {
  try {
    // 嘗試從 B3 儲存格讀取選擇的資料夾名稱
    const sheet = getOrCreateSheet_(INDEX_SHEET_NAME);
    const selectedFolderName = sheet.getRange('B3').getValue();
    
    if (selectedFolderName && selectedFolderName !== '') {
      // 從 DriveID 工作表查找對應的 ID
      const driveIdSheet = getOrCreateSheet_(DRIVEID_SHEET_NAME);
      const lastRow = driveIdSheet.getLastRow();
      
      if (lastRow >= 3) {
        const data = driveIdSheet.getRange(3, 1, lastRow - 2, 2).getValues(); // A3:B 範圍
        
        for (let i = 0; i < data.length; i++) {
          const folderName = String(data[i][0] || '').trim();
          const folderId = String(data[i][1] || '').trim();
          
          if (folderName === selectedFolderName && folderId) {
            return DriveApp.getFolderById(folderId);
          }
        }
      }
    }
    
    // 如果找不到對應的 ID 或 B3 為空，使用預設的 ROOT_FOLDER_ID
    return DriveApp.getFolderById(ROOT_FOLDER_ID);
  } catch (e) {
    console.log('取得根資料夾時發生錯誤，使用預設值:', e);
    return DriveApp.getFolderById(ROOT_FOLDER_ID);
  }
}

// MIME 類型轉換為友善名稱
function toFriendlyType_(mime) {
  switch (mime) {
    case 'application/vnd.google-apps.spreadsheet': return 'Google 試算表';
    case 'application/vnd.google-apps.document': return 'Google 文件';
    case 'application/vnd.google-apps.presentation': return 'Google 簡報';
    case 'application/pdf': return 'PDF';
    case 'application/vnd.ms-excel':
    case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': return 'Excel';
    case 'application/msword':
    case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': return 'Word';
    case 'application/vnd.ms-powerpoint':
    case 'application/vnd.openxmlformats-officedocument.presentationml.presentation': return 'PowerPoint';
    default: return mime;
  }
}

// 填滿層級陣列的輔助函式
function fillLevels_(segments) {
  return [segments[0] || '', segments[1] || '', segments[2] || '', segments[3] || ''];
}

// 批次寫入資料的優化函式
function writeBatchedRows_(sheet, rowUpdates, width) {
  // rowUpdates: Array<{row, values}>
  // 先依 row 排序，並把連續 row 合併成區段
  const sorted = rowUpdates.slice().sort((a, b) => a.row - b.row);
  let i = 0;
  while (i < sorted.length) {
    let start = i;
    let end = i;
    // 擴展到相鄰的連續列
    while (end + 1 < sorted.length && sorted[end + 1].row === sorted[end].row + 1) {
      end++;
    }
    const count = end - start + 1;
    const startRow = sorted[start].row;
    const values = new Array(count);
    for (let k = 0; k < count; k++) values[k] = sorted[start + k].values;
    sheet.getRange(startRow, 1, count, width).setValues(values);
    i = end + 1;
  }
}

// 保存前3行的內容和格式
function saveTopRowsContent_(sheet) {
  const savedData = {};
  if (sheet.getLastRow() >= 3) {
    const lastCol = Math.max(sheet.getLastColumn(), 10); // 至少10欄
    
    // 保存第1-3行的內容
    for (let row = 1; row <= 3; row++) {
      const range = sheet.getRange(row, 1, 1, lastCol);
      savedData[`row${row}`] = {
        values: range.getValues(),
        backgrounds: range.getBackgrounds(),
        horizontalAlignments: range.getHorizontalAlignments(),
        verticalAlignments: range.getVerticalAlignments(),
        fontWeights: range.getFontWeights(),
        fontStyles: range.getFontStyles(),
        fontSizes: range.getFontSizes(),
        fontFamilies: range.getFontFamilies(),
        fontColors: range.getFontColors()
      };
    }
  }
  return savedData;
}

// 恢復前3行的內容和格式
function restoreTopRowsContent_(sheet, savedData) {
  if (!savedData) return;
  
  const lastCol = Math.max(sheet.getLastColumn(), 10); // 至少10欄
  
  // 恢復第1-3行的內容和格式
  for (let row = 1; row <= 3; row++) {
    const data = savedData[`row${row}`];
    if (data) {
      const range = sheet.getRange(row, 1, 1, lastCol);
      
      // 恢復內容
      range.setValues(data.values);
      
      // 恢復格式
      range.setBackgrounds(data.backgrounds);
      range.setHorizontalAlignments(data.horizontalAlignments);
      range.setVerticalAlignments(data.verticalAlignments);
      range.setFontWeights(data.fontWeights);
      range.setFontStyles(data.fontStyles);
      range.setFontSizes(data.fontSizes);
      range.setFontFamilies(data.fontFamilies);
      range.setFontColors(data.fontColors);
    }
  }
}

// 在 C3 建立下拉式選單，資料來源為 B5 以下的所有 B 欄資料
function createDropdownMenu_(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) {
      // 如果沒有資料，清除 C3 的內容
      sheet.getRange(3, 3).clear();
      return;
    }
    
    // 取得 B5 以下的所有 B 欄資料
    const dataRange = sheet.getRange(5, 2, lastRow - 4, 1); // B5 開始到最後一行
    const values = dataRange.getValues();
    
    // 過濾出非空且唯一的資料夾層級1名稱
    const uniqueValues = [...new Set(values
      .map(row => String(row[0] || '').trim())
      .filter(val => val !== '')
    )];
    
    if (uniqueValues.length === 0) {
      // 如果沒有有效資料，清除 C3 的內容
      sheet.getRange(3, 3).clear();
      return;
    }
    
    // 在 C3 建立下拉式選單
    const cell = sheet.getRange(3, 3);
    
    // 清除現有的資料驗證規則
    try {
      cell.getDataValidation().clearValidation();
    } catch (e) {}
    
    // 建立新的資料驗證規則
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueValues, true) // true 表示允許使用者輸入其他值
      .setAllowInvalid(false)
      .setHelpText('請選擇資料夾層級1')
      .build();
    
    cell.setDataValidation(rule);
    
    // 只有在 C3 為空時才設定預設值，避免覆蓋現有選擇
    const currentValue = cell.getValue();
    if (!currentValue || currentValue === '') {
      sheet.getRange(3, 3).setValue('選擇資料夾層級1');
    }
    
  } catch (e) {
    console.log('建立下拉式選單時發生錯誤:', e);
  }
}


// 檢查性能統計
function checkPerformanceStats() {
  console.log('檢查性能統計...');
  
  const cacheSheet = getOrCreateSheet_(CACHE_SHEET_NAME);
  const indexSheet = getOrCreateSheet_(INDEX_SHEET_NAME);
  
  const cacheRows = cacheSheet.getLastRow() - 1; // 減去表頭
  const indexRows = indexSheet.getLastRow() - 4; // 減去前4行
  
  console.log(`Cache 工作表: ${cacheRows} 個項目`);
  console.log(`Index 工作表: ${indexRows} 個項目`);
  // 估算掃描時間
  const estimatedScanTime = Math.ceil(cacheRows * 0.1); // 假設每個項目 0.1ms
  console.log(`預估掃描時間: ${estimatedScanTime}ms`);
  
  // 估算格式化時間
  const estimatedFormatTime = Math.ceil(indexRows * 0.5); // 假設每個項目 0.5ms
  console.log(`預估格式化時間: ${estimatedFormatTime}ms`);
  
  SpreadsheetApp.getUi().alert(
    `性能統計：\n` +
    `Cache 項目: ${cacheRows}\n` +
    `Index 項目: ${indexRows}\n` +
    `預估掃描時間: ${estimatedScanTime}ms\n` +
    `預估格式化時間: ${estimatedFormatTime}ms`
  );
}


