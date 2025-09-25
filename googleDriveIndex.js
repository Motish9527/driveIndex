/***** Google Drive 索引主要功能 *****/
// 注意：所有全域常數和共用函式已移至 global.js

/***** ID→row 映射系統 *****/
function buildIndexIdRowMap() {
  const sheet = getOrCreateSheet_(INDEX_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < SOURCE_START_ROW) {
    return {};
  }
  
  const idMap = {};
  try {
    // 讀取 ID 欄（A 欄）從第5行開始
    const ids = sheet.getRange(SOURCE_START_ROW, 1, lastRow - SOURCE_START_ROW + 1, 1)
                    .getDisplayValues()
                    .map(row => String(row[0] || '').trim())
                    .filter(id => id !== '');
    
    // 建立 { id: rowNumber } 映射
    for (let i = 0; i < ids.length; i++) {
      const id = ids[i];
      if (id) {
        idMap[id] = SOURCE_START_ROW + i;
      }
    }
    
    console.log('ID 映射已建立，共', Object.keys(idMap).length, '個項目');
  } catch (e) {
    console.error('建立 ID 映射失敗:', e);
  }
  
  return idMap;
}

/***** 主流程 *****/

// 全掃描重建（會重置表頭）
function fullRebuild() {
  const root = getRootFolder_();
  const indexSheet = prepareIndexSheet_(true);
  const cacheSheet = prepareCacheSheet_(true);
  const timeZone = Session.getScriptTimeZone();

  const outIndex = [];
  const outCache = [];

  traverse_(root, [], '', (folder, path, parentId) => {
    const levels = fillLevels_(path);
    const fullPath = path.join('/');
    outIndex.push([
      folder.getId(), levels[0], levels[1], levels[2], levels[3],
      folder.getName(), 'Folder', folder.getUrl(), '', fullPath
    ]);
    outCache.push(buildCacheRowForFolderFast_(folder, path, timeZone, '', fullPath, parentId));
  }, (file, path, parentId) => {
    const mime = file.getMimeType();
    if (!ALLOWED_MIME_SET[mime]) return;
    const levels = fillLevels_(path);
    const fullPath = path.concat([file.getName()]).join('/');
    outIndex.push([
      file.getId(), levels[0], levels[1], levels[2], levels[3],
      file.getName(), toFriendlyType_(mime), file.getUrl(), '', fullPath
    ]);
    outCache.push(buildCacheRowForFileFast_(file, path, timeZone, '', fullPath, parentId));
  });

  appendRowsIfAny_(indexSheet, outIndex);
  appendRowsIfAny_(cacheSheet, outCache);
  postFormatIndexSheet_(indexSheet);
}

// 增量更新：新增、更新（改名/搬家）、刪除
function incrementalUpdate() {
  console.log('開始增量更新...');
  const startTime = new Date().getTime();
  
  const root = getRootFolder_();
  const indexSheet = prepareIndexSheet_(false); // 明確指定不清除內容
  // 先把層級欄位的既有合併拆掉，避免刪行/排序卡住
  preUnmergeForUpdate_(indexSheet);
  const cacheSheet = prepareCacheSheet_();  // 不清空
  const timeZone = Session.getScriptTimeZone();

  // 讀取既有快取 -> Map
  console.log('讀取快取資料...');
  const cache = readCacheAsMap_(cacheSheet);              // id -> {row, values[]}
  const seenNow = {};                                     // 現場掃到的 id

  const toAddIndex = [];
  const toAddCache = [];
  const toUpdateCache = [];    // {row, values[]}
  const updatesForIndex = [];  // {row, values[]} 需先建 indexRowMap

  // 建立 Index 的 ID -> row 快取，用於就地更新與刪除
  console.log('建立 ID 映射...');
  const indexMap = buildIndexIdRowMap_(indexSheet);       // id -> row

  // 掃描（資料夾 + 檔案）
  console.log('開始掃描 Google Drive...');
  const scanStartTime = new Date().getTime();
  
  traverse_(root, [], '', (folder, path, parentId) => {
    const id = folder.getId();
    seenNow[id] = true;
    const levels = fillLevels_(path);
    const fullPath = path.join('/');

    const newCacheRow = buildCacheRowForFolderFast_(folder, path, timeZone, cache[id]?.values[10] || '', fullPath, parentId);
    const newIndexRow = [
      id, levels[0], levels[1], levels[2], levels[3],
      folder.getName(), 'Folder', folder.getUrl(),
      cache[id]?.values[10] || '', fullPath
    ];

    if (!cache[id]) {
      // 新增
      toAddCache.push(newCacheRow);
      toAddIndex.push(newIndexRow);
    } else {
      // 可能更名/搬家 -> 比對核心欄位（名稱/父層/層級/FullPath/連結 等）
      if (needsCacheUpdate_(cache[id].values, newCacheRow)) {
        toUpdateCache.push({ row: cache[id].row, values: newCacheRow });
      }
      if (needsIndexUpdate_(indexSheet, indexMap[id], newIndexRow)) {
        updatesForIndex.push({ row: indexMap[id], values: newIndexRow });
      }
    }
  }, (file, path, parentId) => {
    const mime = file.getMimeType();
    if (!ALLOWED_MIME_SET[mime]) return;

    const id = file.getId();
    seenNow[id] = true;

    const levels = fillLevels_(path);
    const fullPath = path.concat([file.getName()]).join('/');

    const newCacheRow = buildCacheRowForFileFast_(file, path, timeZone, cache[id]?.values[10] || '', fullPath, parentId);
    const newIndexRow = [
      id, levels[0], levels[1], levels[2], levels[3],
      file.getName(), toFriendlyType_(mime), file.getUrl(),
      cache[id]?.values[10] || '', fullPath
    ];

    if (!cache[id]) {
      toAddCache.push(newCacheRow);
      toAddIndex.push(newIndexRow);
    } else {
      if (needsCacheUpdate_(cache[id].values, newCacheRow)) {
        toUpdateCache.push({ row: cache[id].row, values: newCacheRow });
      }
      if (needsIndexUpdate_(indexSheet, indexMap[id], newIndexRow)) {
        updatesForIndex.push({ row: indexMap[id], values: newIndexRow });
      }
    }
  });

  const scanEndTime = new Date().getTime();
  console.log(`掃描完成，耗時: ${scanEndTime - scanStartTime}ms`);

  // 刪除：在舊快取有、這次沒看到的 ID
  console.log('處理刪除項目...');
  const toDeleteIds = [];
  Object.keys(cache).forEach(id => {
    if (!seenNow[id]) toDeleteIds.push(id);
  });

  // 先處理刪除（Index & Cache 同步刪）
  if (toDeleteIds.length > 0) {
    console.log(`需要刪除 ${toDeleteIds.length} 個項目`);
    // 由於刪列要由後往前，先換成 row 陣列並排序
    const cacheRows = toDeleteIds
      .map(id => cache[id]?.row)
      .filter(r => r && r >= 2)
      .sort((a, b) => b - a);

    const indexRows = toDeleteIds
      .map(id => indexMap[id])
      .filter(r => r && r >= 5)
      .sort((a, b) => b - a);

    cacheRows.forEach(r => cacheSheet.deleteRow(r));
    indexRows.forEach(r => indexSheet.deleteRow(r));
  }

  // 新增
  console.log(`新增 ${toAddCache.length} 個項目到 Cache`);
  appendRowsIfAny_(cacheSheet, toAddCache);
  console.log(`新增 ${toAddIndex.length} 個項目到 Index`);
  appendRowsIfAny_(indexSheet, toAddIndex);

  // 更新（就地覆寫）
  if (toUpdateCache.length > 0) {
    console.log(`更新 ${toUpdateCache.length} 個 Cache 項目`);
    writeBatchedRows_(cacheSheet, toUpdateCache, CACHE_HEADERS.length);
  }
  if (updatesForIndex.length > 0) {
    console.log(`更新 ${updatesForIndex.length} 個 Index 項目`);
    writeBatchedRows_(indexSheet, updatesForIndex, INDEX_HEADERS.length);
  }

  // 只在有變更時才進行格式化
  if (toAddIndex.length > 0 || toUpdateCache.length > 0 || updatesForIndex.length > 0 || toDeleteIds.length > 0) {
    console.log('執行格式化...');
    const formatStartTime = new Date().getTime();
    
    postFormatIndexSheet_(indexSheet);
    
    const formatEndTime = new Date().getTime();
    console.log(`格式化完成，耗時: ${formatEndTime - formatStartTime}ms`);
  } else {
    console.log('沒有變更，跳過格式化');
  }

  const endTime = new Date().getTime();
  console.log(`增量更新完成，總耗時: ${endTime - startTime}ms`);
  console.log(`- 新增: ${toAddIndex.length} 個項目`);
  console.log(`- 更新: ${toUpdateCache.length} 個項目`);
  console.log(`- 刪除: ${toDeleteIds.length} 個項目`);
}

// 由 Cache 重建 Index（保持 TAG）
function rebuildIndexFromCache() {
  const cacheSheet = prepareCacheSheet_();
  const indexSheet = prepareIndexSheet_(true);
  const lastRow = cacheSheet.getLastRow();
  if (lastRow < 2) return;
  const cacheValues = cacheSheet.getRange(2, 1, lastRow - 1, CACHE_HEADERS.length).getValues();

  const indexRows = cacheValues.map(row => {
    const id = row[0], type = row[1], name = row[2], mimeOrFriendly = row[3], link = row[4];
    const level1 = row[6], level2 = row[7], level3 = row[8], level4 = row[9];
    const tag = row[10] || '';
    const fullPath = row[12] || '';
    const friendlyType = (type === 'Folder') ? 'Folder' : toFriendlyType_(mimeOrFriendly);
    return [id, level1, level2, level3, level4, name, friendlyType, link, tag, fullPath];
  });

  appendRowsIfAny_(indexSheet, indexRows);
  postFormatIndexSheet_(indexSheet);
}

// 合併層級欄
function mergeCells() {
  const sheet = getOrCreateSheet_(INDEX_SHEET_NAME);
  const cacheSheet = getOrCreateSheet_(CACHE_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;
  const startCol = 1; // A
  const endCol = 5;   // E
  const dataRows = lastRow - 4;

  // 1) 解除 Index A~E 既有合併
  const rng = sheet.getRange(5, startCol, dataRows, endCol - startCol + 1);
  try { rng.breakApart(); } catch (e) {}

  // 2) 先依 Full Path 對 Index 做排序（視覺與 Cache 對齊）
  const lastCol = sheet.getLastColumn();
  sheet.getRange(5, 1, dataRows, lastCol).sort([{ column: FULL_PATH_COL, ascending: true }]);

  // 3) 讀取 Cache 並建立 id -> levels 的映射（以 Cache 為事實來源）
  const cacheLastRow = cacheSheet.getLastRow();
  const cacheMap = {};
  if (cacheLastRow >= 2) {
    const cacheValues = cacheSheet.getRange(2, 1, cacheLastRow - 1, CACHE_HEADERS.length).getValues();
    for (let i = 0; i < cacheValues.length; i++) {
      const id = String(cacheValues[i][0] || '').trim();
      if (!id) continue;
      // 層級1..4 位於 6..9（0-based）
      cacheMap[id] = [
        String(cacheValues[i][6] || ''),
        String(cacheValues[i][7] || ''),
        String(cacheValues[i][8] || ''),
        String(cacheValues[i][9] || '')
      ];
    }
  }

  // 4) 依照 Index 當前（已排序）行序列，取出對應的 Cache 層級鍵值
  const indexIds = sheet.getRange(5, 1, dataRows, 1).getValues().map(r => String(r[0] || ''));
  const keysB = new Array(dataRows);
  const keysC = new Array(dataRows);
  const keysD = new Array(dataRows);
  const keysE = new Array(dataRows);
  for (let i = 0; i < dataRows; i++) {
    const levels = cacheMap[indexIds[i]] || ['', '', '', ''];
    keysB[i] = levels[0];
    keysC[i] = levels[1];
    keysD[i] = levels[2];
    keysE[i] = levels[3];
  }

  // 5) 依 Cache 鍵值合併（B~E 欄），空字串不合併
  mergeByKeysDown_(sheet, 5, 2, keysB); // B
  mergeByKeysDown_(sheet, 5, 3, keysC); // C
  mergeByKeysDown_(sheet, 5, 4, keysD); // D
  mergeByKeysDown_(sheet, 5, 5, keysE); // E
}

// 依鍵值陣列在指定欄位向下合併連續且相同的值（忽略空字串）
function mergeByKeysDown_(sheet, startRow, col, keys) {
  const n = keys.length;
  if (n <= 1) return;
  let groupStart = 0; // index within keys
  for (let i = 1; i <= n; i++) {
    const prev = String(keys[i - 1] || '');
    const cur  = (i < n) ? String(keys[i] || '') : '__END__';
    const same = (i < n) && prev !== '' && cur === prev;
    if (!same) {
      const height = i - groupStart;
      if (height > 1 && prev !== '') {
        sheet.getRange(startRow + groupStart, col, height, 1).merge();
      }
      groupStart = i;
    }
  }
}



/***** 掃描與共用工具 *****/

function traverse_(folder, pathSegments, parentFolderId, onFolder, onFile) {
  if (folder.isTrashed && folder.isTrashed()) return;

  // 非 root 層的資料夾本身
  if (pathSegments.length > 0) onFolder(folder, pathSegments, parentFolderId);

  // 檔案
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.isTrashed && f.isTrashed()) continue;
    // 檔案的父層即為目前的資料夾
    onFile(f, pathSegments, folder.getId());
  }

  // 子資料夾
  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const sub = folders.next();
    if (sub.isTrashed && sub.isTrashed()) continue;
    const nextPath = pathSegments.concat([sub.getName()]);
    // 子資料夾的父層為目前資料夾
    traverse_(sub, nextPath, folder.getId(), onFolder, onFile);
  }
}

function prepareIndexSheet_(clear) {
  const sheet = getOrCreateSheet_(INDEX_SHEET_NAME);
  
  if (clear) {
    // 保存前3行的內容和格式
    const savedTopRows = saveTopRowsContent_(sheet);
    
    // 只清除從第 5 行開始的內容，保留前 4 行
    const lastRow = sheet.getLastRow();
    if (lastRow >= 5) {
      const lastCol = sheet.getLastColumn();
      if (lastCol > 0) {
        sheet.getRange(5, 1, lastRow - 4, lastCol).clearContent();
      }
    }
    
    // 處理表頭
    ensureHeaders_(sheet, INDEX_HEADERS);
    
    // 恢復前3行的內容和格式
    restoreTopRowsContent_(sheet, savedTopRows);
  } else {
    ensureHeaders_(sheet, INDEX_HEADERS);
  }

  // 隱藏 ID 欄（A 欄）
  sheet.hideColumns(1);

  return sheet;
}

function prepareCacheSheet_(clear) {
  const sheet = getOrCreateSheet_(CACHE_SHEET_NAME);
  if (clear) sheet.clearContents();
  ensureHeaders_(sheet, CACHE_HEADERS);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const width = headers.length;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(4, 1, 1, width).setValues([headers]);
    sheet.setFrozenRows(4);
  } else {
    // 檢查第4行是否為空或表頭不正確
    const cur = sheet.getRange(4, 1, 1, width).getValues()[0];
    const isEmpty = cur.every(cell => !cell || cell.toString().trim() === '');
    const curStr = cur.join('\t');
    
    if (isEmpty || curStr !== headers.join('\t')) {
      // 只清除第4行和第5行以下的內容，保留前3行
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow >= 4) {
        // 清除第4行
        if (lastCol > 0) {
          sheet.getRange(4, 1, 1, lastCol).clearContent();
        }
        // 清除第5行以下的所有內容
        if (lastRow >= 5) {
          sheet.getRange(5, 1, lastRow - 4, lastCol).clearContent();
        }
      }
      sheet.getRange(4, 1, 1, width).setValues([headers]);
      sheet.setFrozenRows(4);
    }
  }
  
  // 如果是 Index 表，建立 C3 的下拉式選單
  if (headers === INDEX_HEADERS) {
    createDropdownMenu_(sheet);
  }
}

function appendRowsIfAny_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = Math.max(5, sheet.getLastRow() + 1);
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

// 手動更新下拉式選單
function updateDropdownMenu() {
  const sheet = getOrCreateSheet_(INDEX_SHEET_NAME);
  createDropdownMenu_(sheet);
  
  // 同時更新 B3 的雲端資料夾選擇下拉式選單
  try {
    updateDropdownB3_();
  } catch (e) {
    console.log('更新 B3 下拉式選單時發生錯誤:', e);
  }
  
  SpreadsheetApp.getUi().alert('所有下拉式選單已更新！');
}


// 注意：writeBatchedRows_ 已移至 global.js

function postFormatIndexSheet_(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 5) return;

    // 先把層級欄任何舊的合併拆掉
    preUnmergeForUpdate_(sheet);

    // 1) 依「Full Path」（J 欄）做 A→Z 排序，之後再合併層級欄
    sheet.getRange(5, 1, lastRow - 4, lastCol).sort([{ column: FULL_PATH_COL, ascending: true }]);

    // 2) 合併層級欄中連續相同值（滿足你的「上下相同就合併」）
    for (let col = LEVEL_START_COL; col <= LEVEL_END_COL; col++) {
      mergeEqualDown_(sheet, col);
    }

    // 3) 檔案類型為 Folder 底色（從 F 欄起）
    applyFolderRowColor_(sheet);

    // 4) 只在未合併的欄位建立篩選器（F~最後欄），避免與合併衝突
    applyFilterForNonMergedCols_(sheet);

    // 5) 自動欄寬（已停用以保留使用者設定的欄寬）
    // for (let c = 1; c <= lastCol; c++) {
    //   try { sheet.autoResizeColumn(c); } catch (e) {}
    // }
    
  // 6) 更新 C3 的下拉式選單
  createDropdownMenu_(sheet);
  } catch (e) {
    // 靜默失敗即可（例如使用者同時操作時）
  }
}


function readCacheAsMap_(cacheSheet) {
  const map = {};
  const lastRow = cacheSheet.getLastRow();
  if (lastRow < 2) return map;
  const values = cacheSheet.getRange(2, 1, lastRow - 1, CACHE_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const id = String(values[i][0] || '').trim();
    if (!id) continue;
    map[id] = { row: i + 2, values: values[i] };
  }
  return map;
}

function buildIndexIdRowMap_(indexSheet) {
  const map = {};
  const lastRow = indexSheet.getLastRow();
  if (lastRow < 5) return map;
  const ids = indexSheet.getRange(5, 1, lastRow - 4, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const id = String(ids[i][0] || '').trim();
    if (id) map[id] = i + 5; // row index
  }
  return map;
}

function needsCacheUpdate_(oldRow, newRow) {
  // 比對：名稱(2)、MIME/Folder(3)、連結(4)、父層ID(5)、層級1-4(6..9)、FullPath(12)
  const keys = [2, 3, 4, 5, 6, 7, 8, 9, 12];
  for (const k of keys) {
    if (String(oldRow[k] || '') !== String(newRow[k] || '')) return true;
  }
  return false;
}

function needsIndexUpdate_(indexSheet, row, newIndexRow) {
  if (!row || row < 5) return false;
  const width = INDEX_HEADERS.length;
  const cur = indexSheet.getRange(row, 1, 1, width).getValues()[0];
  // 比對 B..J（ID(A) 固定）、實務上最常變動：名稱/層級/連結/FullPath/類型
  for (let c = 2; c <= width; c++) {
    if (String(cur[c - 1] || '') !== String(newIndexRow[c - 1] || '')) return true;
  }
  return false;
}

function buildCacheRowForFolder_(folder, path, timeZone, tag, fullPath) {
  return [
    folder.getId(),
    'Folder',
    folder.getName(),
    'application/vnd.google-apps.folder',
    folder.getUrl(),
    folder.getParents().hasNext() ? folder.getParents().next().getId() : '',
    path[0] || '',
    path[1] || '',
    path[2] || '',
    path[3] || '',
    tag || '',
    Utilities.formatDate(new Date(), timeZone, DATE_FORMAT),
    fullPath || path.join('/')
  ];
}

function buildCacheRowForFile_(file, path, timeZone, tag, fullPath) {
  return [
    file.getId(),
    toFriendlyType_(file.getMimeType()),
    file.getName(),
    file.getMimeType(),
    file.getUrl(),
    file.getParents().hasNext() ? file.getParents().next().getId() : '',
    path[0] || '',
    path[1] || '',
    path[2] || '',
    path[3] || '',
    tag || '',
    Utilities.formatDate((file.getLastUpdated ? file.getLastUpdated() : new Date()), timeZone, DATE_FORMAT),
    fullPath || path.concat([file.getName()]).join('/')
  ];
}

// 快速版本：避免在遍歷時再次呼叫 getParents / getLastUpdated
function buildCacheRowForFolderFast_(folder, path, timeZone, tag, fullPath, parentId) {
  return [
    folder.getId(),
    'Folder',
    folder.getName(),
    'application/vnd.google-apps.folder',
    folder.getUrl(),
    parentId || '',
    path[0] || '',
    path[1] || '',
    path[2] || '',
    path[3] || '',
    tag || '',
    Utilities.formatDate(new Date(), timeZone, DATE_FORMAT),
    fullPath || path.join('/')
  ];
}

function buildCacheRowForFileFast_(file, path, timeZone, tag, fullPath, parentId) {
  return [
    file.getId(),
    toFriendlyType_(file.getMimeType()),
    file.getName(),
    file.getMimeType(),
    file.getUrl(),
    parentId || '',
    path[0] || '',
    path[1] || '',
    path[2] || '',
    path[3] || '',
    tag || '',
    Utilities.formatDate((file.getLastUpdated ? file.getLastUpdated() : new Date()), timeZone, DATE_FORMAT),
    fullPath || path.concat([file.getName()]).join('/')
  ];
}

// 注意：fillLevels_, toFriendlyType_, getOrCreateSheet_, getRootFolder_ 已移至 gobal.js

/***** 觸發器 *****/
// 自動安裝定時觸發器（台北時間每日 8:00 和 14:00）
function installAutoTriggers_() {
  removeTriggers_();
  
  // 設定台北時間每日 8:00 觸發
  ScriptApp.newTrigger('incrementalUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(0) // UTC 0:00 = 台北時間 8:00
    .create();
  
  // 設定台北時間每日 14:00 觸發  
  ScriptApp.newTrigger('incrementalUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(6) // UTC 6:00 = 台北時間 14:00
    .create();
    
  console.log('已設定自動觸發器：台北時間每日 8:00 和 14:00 執行增量更新');
}

// 移除所有現有觸發器
function removeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const fn = t.getHandlerFunction && t.getHandlerFunction();
    if (fn === 'incrementalUpdate' || fn === 'fullRebuild') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// 初始化時自動安裝觸發器（由 global.js 的 onOpen 函數調用）
function initializeAutoTriggers() {
  // 檢查是否已有觸發器，如果沒有則自動安裝
  const triggers = ScriptApp.getProjectTriggers();
  const hasAutoTrigger = triggers.some(t => {
    const fn = t.getHandlerFunction && t.getHandlerFunction();
    return fn === 'incrementalUpdate';
  });
  
  if (!hasAutoTrigger) {
    installAutoTriggers_();
  }
}

/***** 更新合併層級欄 *****/

// 將層級欄（B~E）任何舊的合併拆開，避免更新/刪除衝突
function preUnmergeForUpdate_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;
  const rng = sheet.getRange(5, LEVEL_START_COL, lastRow - 4, LEVEL_END_COL - LEVEL_START_COL + 1);
  try { rng.breakApart(); } catch (e) {}
}

// 合併指定欄位中，連續且相同文字的儲存格（從第5列開始）
function mergeEqualDown_(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return; // 沒有資料或只有一列
  let start = 5;
  const values = sheet.getRange(5, col, lastRow - 4, 1).getValues().map(r => String(r[0] || ''));
  for (let i = 6; i <= lastRow + 1; i++) {
    const prev = values[i - 6];
    const cur  = (i <= lastRow) ? values[i - 5] : '__END__'; // 末尾哨兵
    const same = (cur === prev && prev !== '');
    if (!same) {
      const height = i - start;
      if (height > 1 && prev !== '') {
        sheet.getRange(start, col, height - 1, 1).merge();
      }
      start = i;
    }
  }
}

// 設定條件格式：檔案類型為 Folder 的整列底色為淺綠色3
function applyFolderRowColor_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 5) return;

  const rules = sheet.getConditionalFormatRules() || [];
  // 移除舊的同類規則（避免重複）
  const filtered = rules.filter(r => {
    const cond = r.getBooleanCondition && r.getBooleanCondition();
    if (!cond) return true;
    const isCustom = cond.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA;
    const formula = isCustom ? cond.getCriteriaValues()[0] : '';
    return formula !== '=$G5="Folder"';
  });

  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G5="Folder"')
    .setBackground(LIGHT_GREEN3)
    // 只在未合併欄位（F~最後）上色，避免與 B~E 合併衝突
    .setRanges([sheet.getRange(5, FORMAT_START_COL, lastRow - 4, lastCol - FORMAT_START_COL + 1)])
    .build();

  sheet.setConditionalFormatRules([...filtered, rule]);
}

// 只在不合併的欄位建立篩選器（F~最後一欄：檔案名稱～Full Path）
function applyFilterForNonMergedCols_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 4 || lastCol < 6) return;
  try {
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.getRange(4, 6, lastRow - 3, lastCol - 5).createFilter(); // F..last
  } catch (e) {
    // 若因為其他合併導致無法建立篩選，忽略即可
  }
}

