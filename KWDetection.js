/***** Google 風格搜尋系統 *****/
// 在 F3 儲存格提供類似 Google 搜尋的關鍵字搜尋功能

/***** 建立 F3 搜尋輸入框 *****/
function setupF3Search_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  
  // 目標儲存格：F3
  const targetCell = sh.getRange('F3');
  
  // 清除現有的資料驗證
  targetCell.clearDataValidations();
  
  // 設定儲存格提示文字
  targetCell.setNote('輸入關鍵字進行搜尋（支援多關鍵字，用空格分隔，部分匹配）');
  
  // 設定儲存格樣式
  targetCell.setBackground('#f8f9fa');
  targetCell.setFontStyle('italic');
  targetCell.setHorizontalAlignment('left');
  
  // 如果儲存格是空的，設定預設提示文字
  const currentValue = targetCell.getValue();
  if (!currentValue) {
    targetCell.setValue('輸入關鍵字搜尋...');
    targetCell.setFontColor('#999999');
  }
  
  console.log('F3 搜尋系統已設定完成');
}

/***** 執行搜尋功能 *****/
function performSearch_(searchQuery) {
  console.log('執行搜尋，查詢:', searchQuery);
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  
  // 如果搜尋查詢為空或只有提示文字，顯示所有資料
  if (!searchQuery || searchQuery.trim() === '' || searchQuery === '輸入關鍵字搜尋...') {
    console.log('搜尋查詢為空，顯示所有資料');
    const f = sh.getFilter();
    if (f) f.remove();
    
    // 顯示所有資料行
    const lastRow = sh.getLastRow();
    if (lastRow >= SOURCE_START_ROW) {
      sh.showRows(SOURCE_START_ROW, lastRow - SOURCE_START_ROW + 1);
    }
    return;
  }
  
  // 讀取 Cache 工作表的資料進行搜尋
  const cacheSheet = SpreadsheetApp.getActive().getSheetByName(CACHE_SHEET_NAME);
  if (!cacheSheet) {
    console.log('Cache 工作表不存在');
    const f = sh.getFilter();
    if (f) f.remove();
    return;
  }
  
  const cacheLastRow = cacheSheet.getLastRow();
  if (cacheLastRow < SOURCE_START_ROW) {
    console.log('Cache 工作表沒有足夠的資料');
    const f = sh.getFilter();
    if (f) f.remove();
    return;
  }
  
  // 讀取 Cache 資料：ID (A欄) 和 Full Path (M欄)
  const cacheData = cacheSheet.getRange(SOURCE_START_ROW, 1, cacheLastRow - SOURCE_START_ROW + 1, 13)
                     .getValues();
  
  console.log('Cache 資料總數:', cacheData.length);
  
  // 解析搜尋查詢（支援多關鍵字）
  const searchTerms = searchQuery.toLowerCase().trim().split(/\s+/).filter(term => term !== '');
  console.log('搜尋關鍵字:', searchTerms);
  
  // 找出包含關鍵字的檔案 ID 集合
  const matchingIds = [];
  
  for (let i = 0; i < cacheData.length; i++) {
    const row = cacheData[i];
    const id = String(row[0] || '').trim();
    const fullPath = String(row[12] || '').trim();
    
    if (id && fullPath) {
      const pathLower = fullPath.toLowerCase();
      
      // 檢查是否包含任何搜尋關鍵字（OR 邏輯）
      let isMatch = false;
      
      if (searchTerms.length === 1) {
        // 單一關鍵字：部分匹配
        isMatch = pathLower.includes(searchTerms[0]);
      } else {
        // 多個關鍵字：包含任何一個即可（OR 邏輯）
        isMatch = searchTerms.some(term => pathLower.includes(term));
      }
      
      if (isMatch) {
        matchingIds.push(id);
      }
    }
  }
  
  console.log('搜尋結果 ID 數量:', matchingIds.length);
  
  if (matchingIds.length === 0) {
    // 沒有匹配的結果，顯示所有資料
    console.log('沒有找到匹配的搜尋結果');
    const f = sh.getFilter();
    if (f) f.remove();
    return;
  }
  
  // 取得 Index 的 ID → row 映射
  const idRowMap = buildIndexIdRowMap();
  
  // 將匹配的 ID 轉換為 Index 的行號
  const matchingRows = matchingIds
    .map(id => idRowMap[id])
    .filter(row => row !== undefined)
    .sort((a, b) => a - b);
  
  console.log('匹配的 Index 行數:', matchingRows.length);
  
  // 使用 Filter 顯示結果
  applyFilterByRowsForSearch_(matchingRows);
  
  console.log('搜尋完成，顯示', matchingRows.length, '筆結果');
}

/***** 使用 Hide Row 顯示指定行號的結果（F3 搜索專用） *****/
function applyFilterByRowsForSearch_(matchingRows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  const lastRow = sh.getLastRow();
  
  if (lastRow < SOURCE_START_ROW) {
    return;
  }
  
  // 先顯示所有行
  sh.showRows(SOURCE_START_ROW, lastRow - SOURCE_START_ROW + 1);
  
  // 移除現有篩選器
  const existingFilter = sh.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  
  if (matchingRows.length === 0) {
    // 如果沒有匹配的行，隱藏所有資料行
    sh.hideRows(SOURCE_START_ROW, lastRow - SOURCE_START_ROW + 1);
    return;
  }
  
  // 找出所有需要隱藏的行
  const allDataRows = [];
  for (let row = SOURCE_START_ROW; row <= lastRow; row++) {
    allDataRows.push(row);
  }
  
  const rowsToHide = allDataRows.filter(row => !matchingRows.includes(row));
  
  // 隱藏不匹配的行
  if (rowsToHide.length > 0) {
    // 從後往前隱藏，避免行號變化
    rowsToHide.sort((a, b) => b - a);
    rowsToHide.forEach(row => {
      sh.hideRows(row, 1);
    });
  }
  
  console.log('F3 搜索完成，顯示', matchingRows.length, '行，隱藏', rowsToHide.length, '行');
}

/***** 清除搜尋結果 *****/
function clearSearch_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  const searchCell = sh.getRange('F3');
  
  // 清除搜尋框
  searchCell.setValue('輸入關鍵字搜尋...');
  searchCell.setFontColor('#999999');
  
  // 顯示所有資料行
  const lastRow = sh.getLastRow();
  if (lastRow >= SOURCE_START_ROW) {
    sh.showRows(SOURCE_START_ROW, lastRow - SOURCE_START_ROW + 1);
  }
  
  // 移除篩選器
  const f = sh.getFilter();
  if (f) f.remove();
  
  console.log('搜尋已清除，顯示所有資料');
}

/***** 處理 F3 儲存格編輯事件 *****/
function handleF3Search_(value) {
  console.log('F3 搜尋輸入:', value);
  
  // 如果輸入的是提示文字，不執行搜尋
  if (value === '輸入關鍵字搜尋...') {
    return;
  }
  
  // 設定儲存格樣式（移除提示文字樣式）
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  const searchCell = sh.getRange('F3');
  searchCell.setFontColor('#000000');
  searchCell.setFontStyle('normal');
  
  // 執行搜尋
  performSearch_(value);
}
