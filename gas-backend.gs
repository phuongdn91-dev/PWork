// ============================================================
// PWork — Google Apps Script Backend v2
// ============================================================
// HƯỚNG DẪN DEPLOY:
//   1. Mở Google Sheet > Extensions > Apps Script
//   2. Xóa toàn bộ code mặc định, paste toàn bộ file này vào
//   3. Nhấn Save (Ctrl+S)
//   4. Deploy > New deployment > Type: Web app
//      - Description: PWork API v2
//      - Execute as: Me
//      - Who has access: Anyone
//   5. Click Deploy > Copy Web app URL
//   6. Paste URL vào biến GAS_URL trong index.html
//
// KHI CẬP NHẬT CODE (sau lần deploy đầu):
//   Deploy > Manage deployments > Edit (✏️) > Version: New version > Deploy
// ============================================================

const SHEET_TASKS   = 'CongViec';
const SHEET_FILES   = 'TaiLieu';
const DRIVE_FOLDER  = 'PWork_Attachments';

// Cột trong sheet CongViec (1-based index)
const COL = {
  ID: 1, NAME: 2, CONTENT: 3, LOCATION: 4,
  LEADER: 5, COWORKER: 6, CREATED: 7, DEADLINE: 8,
  NOTES: 9, PROGRESS: 10, DEPLOYED: 11, PLAN: 12,
  REPORT: 13, COMPLETED: 14, FILE_COUNT: 15,
  DRIVE_LINKS: 16, STATUS: 17
};
const NCOLS = 17;

// ============================================================
// ENTRY POINTS
// ============================================================

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'getTasks') {
    return jsonResponse(getAllTasks());
  }
  return jsonResponse({ status: 'ok', version: 'PWork GAS v2' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    switch (data.action) {
      case 'syncAll':    result = syncAll(data.tasks, data.files); break;
      case 'upsertTask': result = upsertTask(data.task, data.files); break;
      case 'deleteTask': result = deleteTask(data.id); break;
      case 'uploadFiles': result = uploadFileBatch(data.files); break;
      default: result = { status: 'error', message: 'Unknown action: ' + data.action };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message, stack: err.stack });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SYNC ALL — Upsert toàn bộ task từ localStorage
// files: [ { taskId, name, type, data (base64), isReport } ]
// ============================================================
function syncAll(tasks, files) {
  if (!tasks || !tasks.length) return { status: 'ok', synced: 0 };

  ensureSheets();

  // Upload files trước, lấy map taskId -> [driveLinks]
  const driveMap = {};
  if (files && files.length) {
    const uploadResults = uploadFileBatch(files);
    uploadResults.results.forEach(r => {
      if (r.status === 'ok') {
        if (!driveMap[r.taskId]) driveMap[r.taskId] = [];
        driveMap[r.taskId].push(r.viewUrl + ' (' + r.fileName + ')');
      }
    });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TASKS);

  // Build lookup map id -> rowIndex từ dữ liệu hiện tại
  const existing = getIdRowMap(sheet);

  let inserted = 0, updated = 0;

  tasks.forEach(t => {
    const driveLinks = (driveMap[t.id] || []).join('\n');
    const row = buildRow(t, driveLinks);

    if (existing[t.id]) {
      // UPDATE existing row
      const rowIdx = existing[t.id];
      sheet.getRange(rowIdx, 1, 1, NCOLS).setValues([row]);
      updated++;
    } else {
      // INSERT new row after header
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, NCOLS).setValues([row]);
      inserted++;
      // Re-apply header style (insertRowAfter có thể kế thừa format)
    }

    // Màu status
    colorStatusRow(sheet, existing[t.id] || 2, t);
  });

  applyHeaderStyle(sheet);
  sheet.autoResizeColumns(1, NCOLS);

  return { status: 'ok', inserted, updated, total: tasks.length };
}

// ============================================================
// UPSERT 1 TASK (auto-sync sau mỗi cập nhật tiến độ)
// ============================================================
function upsertTask(task, files) {
  if (!task || !task.id) return { status: 'error', message: 'Missing task.id' };
  ensureSheets();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TASKS);
  const existing = getIdRowMap(sheet);

  // Upload files nếu có
  let driveLinks = '';
  if (files && files.length) {
    const uploadResults = uploadFileBatch(files);
    const links = uploadResults.results
      .filter(r => r.status === 'ok' && r.taskId === task.id)
      .map(r => r.viewUrl + ' (' + r.fileName + ')');
    driveLinks = links.join('\n');

    // Merge với link cũ đã có trong sheet
    if (existing[task.id]) {
      const old = sheet.getRange(existing[task.id], COL.DRIVE_LINKS).getValue();
      if (old) driveLinks = old + (driveLinks ? '\n' + driveLinks : '');
    }
  }

  const row = buildRow(task, driveLinks);

  if (existing[task.id]) {
    sheet.getRange(existing[task.id], 1, 1, NCOLS).setValues([row]);
    colorStatusRow(sheet, existing[task.id], task);
    return { status: 'ok', action: 'updated', id: task.id };
  } else {
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, NCOLS).setValues([row]);
    applyHeaderStyle(sheet);
    colorStatusRow(sheet, 2, task);
    return { status: 'ok', action: 'inserted', id: task.id };
  }
}

// ============================================================
// DELETE TASK
// ============================================================
function deleteTask(id) {
  if (!id) return { status: 'error', message: 'Missing id' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TASKS);
  if (!sheet) return { status: 'ok', deleted: false };

  const existing = getIdRowMap(sheet);
  if (existing[id]) {
    sheet.deleteRow(existing[id]);
    return { status: 'ok', deleted: true, id };
  }
  return { status: 'ok', deleted: false, id };
}

// ============================================================
// UPLOAD FILES TO GOOGLE DRIVE
// files: [ { taskId, name, type, data (base64 dataURL), isReport } ]
// ============================================================
function uploadFileBatch(files) {
  if (!files || !files.length) return { status: 'ok', results: [] };

  const root = getOrCreateFolder(DRIVE_FOLDER);
  const results = [];

  files.forEach(f => {
    try {
      // Bỏ prefix "data:...;base64,"
      const b64 = f.data.indexOf(',') >= 0 ? f.data.split(',')[1] : f.data;
      const bytes = Utilities.base64Decode(b64);
      const mimeType = f.type || 'application/octet-stream';
      const blob = Utilities.newBlob(bytes, mimeType, f.name);

      // Sub-folder theo taskId
      const taskFolder = getOrCreateFolder('Task_' + f.taskId, root);
      const file = taskFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // Ghi vào sheet TaiLieu
      logFileToSheet(f.taskId, f.name, file.getId(), file.getUrl(), f.isReport || false);

      results.push({
        status: 'ok',
        taskId: f.taskId,
        fileName: f.name,
        fileId: file.getId(),
        viewUrl: file.getUrl(),
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId()
      });
    } catch (err) {
      results.push({ status: 'error', taskId: f.taskId, fileName: f.name, message: err.message });
    }
  });

  return { status: 'ok', results };
}

// ============================================================
// GET ALL TASKS (cho tính năng load từ Sheet về app)
// ============================================================
function getAllTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TASKS);
  if (!sheet) return { status: 'ok', tasks: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'ok', tasks: [] };

  const tasks = data.slice(1).map(row => ({
    id:         row[COL.ID - 1],
    name:       row[COL.NAME - 1],
    content:    row[COL.CONTENT - 1],
    location:   row[COL.LOCATION - 1],
    leader:     row[COL.LEADER - 1],
    coworker:   row[COL.COWORKER - 1],
    createdAt:  row[COL.CREATED - 1],
    deadline:   row[COL.DEADLINE - 1],
    notes:      row[COL.NOTES - 1],
    progress:   row[COL.PROGRESS - 1] || 0,
    deployed:   row[COL.DEPLOYED - 1],
    plan:       row[COL.PLAN - 1],
    report:     row[COL.REPORT - 1],
    completedAt: row[COL.COMPLETED - 1],
    driveLinks: row[COL.DRIVE_LINKS - 1],
    status:     row[COL.STATUS - 1]
  })).filter(t => t.id);

  return { status: 'ok', tasks };
}

// ============================================================
// HELPERS
// ============================================================

function buildRow(t, driveLinks) {
  return [
    t.id || '',
    t.name || '',
    t.content || '',
    t.location || '',
    t.leader || '',
    t.coworker || '',
    t.createdAt  ? formatDate(t.createdAt)  : '',
    t.deadline   ? t.deadline                : '',
    t.notes || '',
    t.progress || 0,
    t.deployed || '',
    t.plan || '',
    t.report || '',
    t.completedAt ? formatDate(t.completedAt) : '',
    t.fileCount || 0,
    driveLinks || '',
    calcStatus(t)
  ];
}

function calcStatus(t) {
  if (t.completedAt || t.progress === 100) return 'Hoàn thành';
  if (t.deadline) {
    const now = new Date(); now.setHours(0,0,0,0);
    if (new Date(t.deadline) < now) return 'Quá hạn';
  }
  if (t.progress > 0) return 'Đang triển khai';
  return 'Chưa bắt đầu';
}

function formatDate(iso) {
  try { return new Date(iso).toLocaleString('vi-VN'); } catch(e) { return iso || ''; }
}

function getIdRowMap(sheet) {
  const map = {};
  if (!sheet) return map;
  const nRows = sheet.getLastRow();
  if (nRows <= 1) return map;
  const ids = sheet.getRange(2, COL.ID, nRows - 1, 1).getValues();
  ids.forEach((row, i) => { if (row[0]) map[row[0]] = i + 2; });
  return map;
}

function colorStatusRow(sheet, rowIdx, task) {
  const status = calcStatus(task);
  const cell = sheet.getRange(rowIdx, COL.STATUS);
  if (status === 'Hoàn thành')      { cell.setBackground('#D1FAE5'); cell.setFontColor('#065F46'); }
  else if (status === 'Quá hạn')    { cell.setBackground('#FEE2E2'); cell.setFontColor('#991B1B'); }
  else if (status === 'Đang triển khai') { cell.setBackground('#FEF3C7'); cell.setFontColor('#92400E'); }
  else                               { cell.setBackground('#EFF6FF'); cell.setFontColor('#1E40AF'); }

  // Progress bar màu
  const pCell = sheet.getRange(rowIdx, COL.PROGRESS);
  const p = task.progress || 0;
  if (p >= 100)     pCell.setBackground('#D1FAE5');
  else if (p >= 66) pCell.setBackground('#DBEAFE');
  else if (p >= 33) pCell.setBackground('#FEF3C7');
  else if (p > 0)   pCell.setBackground('#FEE2E2');
  else              pCell.setBackground('#F3F4F6');
}

function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet CongViec
  let sheet = ss.getSheetByName(SHEET_TASKS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TASKS);
    writeTaskHeader(sheet);
  } else if (sheet.getLastRow() === 0) {
    writeTaskHeader(sheet);
  }

  // Sheet TaiLieu
  let fileSheet = ss.getSheetByName(SHEET_FILES);
  if (!fileSheet) {
    fileSheet = ss.insertSheet(SHEET_FILES);
    writeFileHeader(fileSheet);
  }
}

function writeTaskHeader(sheet) {
  const headers = [
    'ID', 'Tên công việc', 'Nội dung', 'Nơi công tác',
    'Người chủ trì', 'Người phối hợp', 'Thời gian tạo',
    'Deadline', 'Ghi chú', 'Tiến độ (%)',
    'Đã triển khai', 'Kế hoạch', 'Báo cáo kết quả',
    'Ngày hoàn thành', 'Số file', 'Link Drive', 'Trạng thái'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  applyHeaderStyle(sheet);
}

function writeFileHeader(sheet) {
  const headers = ['ID Công việc', 'Tên file', 'File ID', 'Link Drive', 'Loại', 'Ngày upload'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const r = sheet.getRange(1, 1, 1, headers.length);
  r.setBackground('#1565C0'); r.setFontColor('#ffffff');
  r.setFontWeight('bold'); r.setFontSize(10);
}

function applyHeaderStyle(sheet) {
  const r = sheet.getRange(1, 1, 1, NCOLS);
  r.setBackground('#1565C0');
  r.setFontColor('#FFFFFF');
  r.setFontWeight('bold');
  r.setFontSize(10);
  r.setWrap(false);
}

function logFileToSheet(taskId, fileName, fileId, viewUrl, isReport) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_FILES);
  if (!sheet) return;
  sheet.appendRow([
    taskId, fileName, fileId, viewUrl,
    isReport ? 'Báo cáo' : 'Đính kèm',
    new Date().toLocaleString('vi-VN')
  ]);
}

function getOrCreateFolder(name, parent) {
  const base = parent || DriveApp;
  const iter = base.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : base.createFolder(name);
}
