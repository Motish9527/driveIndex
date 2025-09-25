/***** 下拉式選單篩選功能 *****/
// 注意：所有全域常數和共用函式已移至 global.js
// 注意：onOpen 和 onEdit 函式已移至 global.js 並整合

/***** 雲端資料夾：建立/更新 B3 下拉（來源：DriveID 工作表的 A 欄） *****/
function updateDropdownB3_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  const driveIdSheet = SpreadsheetApp.getActive().getSheetByName(DRIVEID_SHEET_NAME);

  // 目標儲存格：B3
  const targetCell = sh.getRange('B3');

  // 檢查 DriveID 工作表是否存在
  if (!driveIdSheet) {
    console.log('DriveID 工作表不存在');
    targetCell.clearDataValidations();
    targetCell.setValue('請先建立 DriveID 工作表');
    return;
  }

  // 取得 DriveID 工作表的資料
  const lastRow = driveIdSheet.getLastRow();
  if (lastRow < 3) {
    console.log('DriveID 工作表沒有資料');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 DriveID 工作表添加資料');
    return;
  }

  // 讀取 A3 開始的資料夾名稱
  const folderNames = driveIdSheet.getRange(3, 1, lastRow - 2, 1).getValues()
    .map(row => String(row[0] || '').trim())
    .filter(name => name !== '');

  if (folderNames.length === 0) {
    console.log('DriveID 工作表沒有有效的資料夾名稱');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 DriveID 工作表添加資料夾名稱');
    return;
  }

  // 建立下拉規則
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(folderNames, true)
    .setAllowInvalid(true)
    .setHelpText('請選擇要索引的雲端資料夾')
    .build();

  targetCell.clearDataValidations();
  targetCell.setDataValidation(rule);

  // 如果儲存格是空的，設定預設值
  const currentValue = targetCell.getValue();
  if ((!currentValue || currentValue === '') && folderNames.length > 0) {
    targetCell.setValue(folderNames[0]); // 設定第一個選項為預設值
  }
}

/***** 目錄資料夾 建立/更新 C3 下拉（來源：Cache 工作表的 G 欄） *****/
function updateDropdownC3_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  const cacheSheet = SpreadsheetApp.getActive().getSheetByName(CACHE_SHEET_NAME);

  // 目標儲存格：C3
  const targetCell = sh.getRange('C3');

  // 檢查 Cache 工作表是否存在
  if (!cacheSheet) {
    console.log('Cache 工作表不存在');
    targetCell.clearDataValidations();
    targetCell.setValue('請先建立 Cache 工作表');
    return;
  }

  // 取得 Cache 工作表的資料
  const lastRow = cacheSheet.getLastRow();
  if (lastRow < SOURCE_START_ROW) {
    console.log('Cache 工作表沒有足夠的資料');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 Cache 工作表添加資料');
    return;
  }

  // 讀取 G 欄（從第5行開始）的所有資料夾名稱
  const folderNames = cacheSheet.getRange(SOURCE_START_ROW, 7, lastRow - SOURCE_START_ROW + 1, 1)
    .getValues()
    .map(row => String(row[0] || '').trim())
    .filter(name => name !== '');

  if (folderNames.length === 0) {
    console.log('Cache 工作表 G 欄沒有有效的資料夾名稱');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 Cache 工作表 G 欄添加資料');
    return;
  }

  // 添加 "ALL" 選項並去重
  const allOptions = ['ALL', ...Array.from(new Set(folderNames))];

  try {
    // 建立下拉規則
  const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allOptions, true)
    .setAllowInvalid(true)
      .setHelpText('請選擇資料夾層級1')
    .build();

  targetCell.clearDataValidations();
  targetCell.setDataValidation(rule);
    
    // 如果儲存格是空的，設定預設值
    const currentValue = targetCell.getValue();
    if (!currentValue || currentValue === '') {
      targetCell.setValue('ALL');
    }
    
    console.log('C3 下拉式選單已建立，資料夾數量:', allOptions.length - 1);
  } catch (e) {
    console.error('建立 C3 下拉式選單時發生錯誤:', e);
    targetCell.clearDataValidations();
    targetCell.setValue('建立下拉式選單失敗');
  }
}

/***** 檔案類型：建立/更新 G3 下拉（來源：Cache 工作表的 B 欄） *****/
function updateDropdownG3_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  const cacheSheet = SpreadsheetApp.getActive().getSheetByName(CACHE_SHEET_NAME);

  // 目標儲存格：G3
  const targetCell = sh.getRange('G3');

  // 檢查 Cache 工作表是否存在
  if (!cacheSheet) {
    console.log('Cache 工作表不存在');
    targetCell.clearDataValidations();
    targetCell.setValue('請先建立 Cache 工作表');
    return;
  }

  // 取得 Cache 工作表的資料
  const lastRow = cacheSheet.getLastRow();
  if (lastRow < SOURCE_START_ROW) {
    console.log('Cache 工作表沒有足夠的資料');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 Cache 工作表添加資料');
    return;
  }

  // 讀取 B 欄（從第5行開始）的所有檔案類型
  const fileTypes = cacheSheet.getRange(SOURCE_START_ROW, 2, lastRow - SOURCE_START_ROW + 1, 1)
    .getValues()
    .map(row => String(row[0] || '').trim())
    .filter(type => type !== '');

  if (fileTypes.length === 0) {
    console.log('Cache 工作表 B 欄沒有有效的檔案類型');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 Cache 工作表 B 欄添加檔案類型');
    return;
  }

  // 添加 "ALL" 選項並去重
  const allOptions = ['ALL', ...Array.from(new Set(fileTypes))];

  try {
    // 建立下拉規則
  const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allOptions, true)
    .setAllowInvalid(true)
      .setHelpText('請選擇檔案類型')
    .build();

  targetCell.clearDataValidations();
  targetCell.setDataValidation(rule);
    
    // 如果儲存格是空的，設定預設值
    const currentValue = targetCell.getValue();
    if (!currentValue || currentValue === '') {
      targetCell.setValue('ALL');
    }
    
    console.log('G3 下拉式選單已建立，檔案類型數量:', allOptions.length - 1);
  } catch (e) {
    console.error('建立 G3 下拉式選單時發生錯誤:', e);
    targetCell.clearDataValidations();
    targetCell.setValue('建立下拉式選單失敗');
  }
}

/***** 依 C3 選擇值套用篩選；若空字串則清除篩選 *****/
function applyFilterByC3_(selected) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);

  // 定義要加篩選器的總範圍：從「表頭列」到最後一列、涵蓋到最後一欄
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < SOURCE_START_ROW) {
    // 沒資料就清篩選並離開
    const f = sh.getFilter();
    if (f) f.remove();
    return;
  }
  const filterRange = sh.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, lastCol);

  // 若沒有篩選器就建立；若已有就覆用
  let filter = sh.getFilter();
  if (!filter) {
    // createFilter 需要第一列是表頭，因此用 HEADER_ROW 當表頭
    filterRange.createFilter();
    filter = sh.getFilter();
  } else {
    // 若現有的篩選器範圍不包含我們的資料，則重建
    const fr = filter.getRange();
    const same =
      fr.getRow() === HEADER_ROW &&
      fr.getColumn() === 1 &&
      fr.getNumRows() === (lastRow - HEADER_ROW + 1) &&
      fr.getNumColumns() === lastCol;
    if (!same) {
      filter.remove();
      filterRange.createFilter();
      filter = sh.getFilter();
    }
  }

  // 若 C3 留白或選擇 ALL → 清除所有條件、移除篩選器
  if (!selected || selected === 'ALL') {
    filter.remove();
    return;
  }

  // 讀取 Cache 工作表的 G 欄（從第5行開始）所有值
  const cacheSheet = SpreadsheetApp.getActive().getSheetByName(CACHE_SHEET_NAME);
  if (!cacheSheet) {
    console.log('Cache 工作表不存在');
    filter.remove();
    return;
  }
  
  const cacheLastRow = cacheSheet.getLastRow();
  if (cacheLastRow < SOURCE_START_ROW) {
    console.log('Cache 工作表沒有足夠的資料');
    filter.remove();
    return;
  }
  
  const values = cacheSheet.getRange(SOURCE_START_ROW, 7, cacheLastRow - SOURCE_START_ROW + 1, 1)
                  .getDisplayValues()
                  .map(r => String(r[0] || '').trim())
                  .filter(v => v !== '');

  // 只保留唯一值
  const uniq = Array.from(new Set(values));

  // 準備要「隱藏」的值（= 除了 selected 以外的全部）
  const hidden = uniq.filter(v => v !== selected);

  const criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(hidden)  // 隱藏其他值，等效於「只顯示 selected」
    .build();

  // 套用到第 2 欄（B 欄）- 資料夾層級1
  filter.setColumnFilterCriteria(2, criteria);
}

/***** 依 G3 選擇值套用篩選；若空字串則清除篩選 *****/
function applyFilterByG3_(selected) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);

  // 定義要加篩選器的總範圍：從「表頭列」到最後一列、涵蓋到最後一欄
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  
  if (lastRow < SOURCE_START_ROW) {
    // 沒資料就清篩選並離開
    console.log('沒有足夠的資料進行篩選');
    const f = sh.getFilter();
    if (f) f.remove();
    return;
  }
  const filterRange = sh.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, lastCol);

  // 若沒有篩選器就建立；若已有就覆用
  let filter = sh.getFilter();
  if (!filter) {
    // createFilter 需要第一列是表頭，因此用 HEADER_ROW 當表頭
    filterRange.createFilter();
    filter = sh.getFilter();
  } else {
    // 若現有的篩選器範圍不包含我們的資料，則重建
    const fr = filter.getRange();
    const same =
      fr.getRow() === HEADER_ROW &&
      fr.getColumn() === 1 &&
      fr.getNumRows() === (lastRow - HEADER_ROW + 1) &&
      fr.getNumColumns() === lastCol;
    if (!same) {
      filter.remove();
      filterRange.createFilter();
      filter = sh.getFilter();
    }
  }

  // 若 G3 留白或選擇 ALL → 清除所有條件、移除篩選器
  if (!selected || selected === 'ALL') {
    filter.remove();
    return;
  }

  // 讀取 Cache 工作表的 B 欄（從第5行開始）所有值
  const cacheSheet = SpreadsheetApp.getActive().getSheetByName(CACHE_SHEET_NAME);
  if (!cacheSheet) {
    console.log('Cache 工作表不存在');
    filter.remove();
    return;
  }
  
  const cacheLastRow = cacheSheet.getLastRow();
  if (cacheLastRow < SOURCE_START_ROW) {
    console.log('Cache 工作表沒有足夠的資料');
    filter.remove();
    return;
  }
  
  const values = cacheSheet.getRange(SOURCE_START_ROW, 2, cacheLastRow - SOURCE_START_ROW + 1, 1)
                  .getDisplayValues()
                  .map(r => String(r[0] || '').trim())
                  .filter(v => v !== '');

  // 只保留唯一值
  const uniq = Array.from(new Set(values));

  // 準備要「隱藏」的值（= 除了 selected 以外的全部）
  const hidden = uniq.filter(v => v !== selected);

  const criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(hidden)  // 隱藏其他值，等效於「只顯示 selected」
    .build();

  // 套用到第 7 欄（G 欄）- 檔案類型
  filter.setColumnFilterCriteria(7, criteria);
}

/***** 關鍵字篩選：建立/更新 E3 下拉（來源：KWData 工作表的 F 欄） *****/
function updateDropdownE3_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
  const kwDataSheet = SpreadsheetApp.getActive().getSheetByName(KWDATA_SHEET_NAME);

  // 目標儲存格：E3
  const targetCell = sh.getRange('E3');

  // 檢查 KWData 工作表是否存在
  if (!kwDataSheet) {
    console.log('KWData 工作表不存在');
    targetCell.clearDataValidations();
    targetCell.setValue('請先建立 KWData 工作表');
    return;
  }

  // 取得 KWData 工作表的資料
  const lastRow = kwDataSheet.getLastRow();
  if (lastRow < 1) {
    console.log('KWData 工作表沒有資料');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 KWData 工作表添加資料');
    return;
  }

  // 讀取 F 欄的關鍵字（從第3行開始，排除第1行和第2行）
  if (lastRow < 3) {
    console.log('KWData 工作表沒有足夠的資料（需要至少3行）');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 KWData 工作表添加資料（從第3行開始）');
    return;
  }
  
  const keywords = kwDataSheet.getRange(3, 6, lastRow - 2, 1).getValues()
    .map(row => String(row[0] || '').trim())
    .filter(keyword => keyword !== '');

  if (keywords.length === 0) {
    console.log('KWData 工作表沒有有效的關鍵字');
    targetCell.clearDataValidations();
    targetCell.setValue('請先在 KWData 工作表 F 欄添加關鍵字');
    return;
  }

  // 添加 "ALL" 選項
  const allOptions = ['ALL', ...keywords];

  try {
    // 建立下拉規則
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allOptions, true)
      .setAllowInvalid(true)
      .setHelpText('請選擇關鍵字進行篩選')
    .build();

    targetCell.clearDataValidations();
    targetCell.setDataValidation(rule);
    
    // 如果儲存格是空的，設定預設值
    const currentValue = targetCell.getValue();
    if (!currentValue || currentValue === '') {
      targetCell.setValue('ALL');
    }
    
    console.log('E3 關鍵字下拉式選單已建立，關鍵字數量:', keywords.length);
  } catch (e) {
    console.error('建立 E3 關鍵字下拉式選單時發生錯誤:', e);
    targetCell.clearDataValidations();
    targetCell.setValue('建立下拉式選單失敗');
  }
}

/***** 依 E3 選擇值套用關鍵字篩選；若空字串則清除篩選 *****/
function applyFilterByE3_(selected) {
  console.log('E3 篩選，選擇值:', selected);
  
  // 若 E3 留白或選擇 ALL → 清除所有條件並顯示所有資料
  if (!selected || selected === 'ALL') {
    const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
    const f = sh.getFilter();
    if (f) f.remove();
    
    // 顯示所有資料行
    const lastRow = sh.getLastRow();
    if (lastRow >= SOURCE_START_ROW) {
      sh.showRows(SOURCE_START_ROW, lastRow - SOURCE_START_ROW + 1);
    }
    
    console.log('E3 篩選已清除，顯示所有資料');
    return;
  }

  // 讀取 Cache 工作表的資料
  const cacheSheet = SpreadsheetApp.getActive().getSheetByName(CACHE_SHEET_NAME);
  if (!cacheSheet) {
    console.log('Cache 工作表不存在');
    const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
    const f = sh.getFilter();
    if (f) f.remove();
    return;
  }
  
  const cacheLastRow = cacheSheet.getLastRow();
  if (cacheLastRow < SOURCE_START_ROW) {
    console.log('Cache 工作表沒有足夠的資料');
    const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
    const f = sh.getFilter();
    if (f) f.remove();
    return;
  }
  
  // 讀取 Cache 資料：ID (A欄) 和 Full Path (M欄)
  const cacheData = cacheSheet.getRange(SOURCE_START_ROW, 1, cacheLastRow - SOURCE_START_ROW + 1, 13)
                     .getValues();
  
  // 找出包含關鍵字的檔案 ID
  const matchingIds = [];
  const selectedLower = selected.toLowerCase();
  
  for (let i = 0; i < cacheData.length; i++) {
    const row = cacheData[i];
    const id = String(row[0] || '').trim();
    const fullPath = String(row[12] || '').trim();
    
    if (id && fullPath && fullPath.toLowerCase().includes(selectedLower)) {
      matchingIds.push(id);
    }
  }
  
  console.log('E3 匹配的 ID 數量:', matchingIds.length);
  
  if (matchingIds.length === 0) {
    const sh = SpreadsheetApp.getActive().getSheetByName(SELECTION_SHEET_NAME);
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
  
  // 使用 Filter 顯示結果
  applyFilterByRowsForE3_(matchingRows);
}

/***** 使用 Hide Row 顯示指定行號的結果（E3 篩選專用） *****/
function applyFilterByRowsForE3_(matchingRows) {
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
  
  console.log('E3 篩選完成，顯示', matchingRows.length, '行，隱藏', rowsToHide.length, '行');
}


