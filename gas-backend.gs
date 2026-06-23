// ============================================================
// PWork — Google Apps Script Backend
// Dán code này vào Extensions > Apps Script trong Google Sheet
// Deploy: Deploy > New deployment > Web App > Execute as: Me > Anyone
// Lấy URL dán vào biến GAS_URL trong index.html
// ============================================================

const SHEET_NAME = 'CongViec';
const DRIVE_FOLDER_NAME = 'PWork_Attachments';

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'PWork GAS is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };
  
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    
    switch (data.action) {
      case 'syncAll':
        result = syncAllTasks(data.tasks);
        break;
      case 'updateTask':
        result = updateTaskInSheet(data.task);
        break;
      case 'uploadFile':
        result = uploadFileToDrive(data.file);
        break;
      default:
        result = { status: 'error', message: 'Unknown action' };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// SYNC ALL TASKS — Ghi toàn bộ danh sách công việc vào Sheet
// ============================================================
function syncAllTasks(tasks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  // Header row
  const headers = [
    'ID', 'Tên công việc', 'Nội dung', 'Nơi công tác',
    'Người chủ trì', 'Người phối hợp', 'Thời gian tạo',
    'Deadline', 'Ghi chú', 'Tiến độ (%)',
    'Đã triển khai', 'Kế hoạch', 'Báo cáo kết quả',
    'Ngày hoàn thành', 'Số file đính kèm', 'Trạng thái'
  ];
  
  // Clear and rewrite
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Style header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0A2540');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  
  if (!tasks || !tasks.length) {
    return { status: 'ok', message: 'No tasks to sync' };
  }
  
  const rows = tasks.map(t => {
    const status = getStatus(t);
    return [
      t.id || '',
      t.name || '',
      t.content || '',
      t.location || '',
      t.leader || '',
      t.coworker || '',
      t.createdAt ? new Date(t.createdAt).toLocaleString('vi-VN') : '',
      t.deadline || '',
      t.notes || '',
      t.progress || 0,
      t.deployed || '',
      t.plan || '',
      t.report || '',
      t.completedAt ? new Date(t.completedAt).toLocaleString('vi-VN') : '',
      t.fileCount || 0,
      status
    ];
  });
  
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
  // Color code status column (col 16)
  rows.forEach((row, i) => {
    const cell = sheet.getRange(i + 2, 16);
    const status = row[15];
    if (status === 'Hoàn thành') { cell.setBackground('#E8F5E9'); cell.setFontColor('#2E7D32'); }
    else if (status === 'Quá hạn') { cell.setBackground('#FFEBEE'); cell.setFontColor('#C62828'); }
    else if (status === 'Đang triển khai') { cell.setBackground('#FFF3E0'); cell.setFontColor('#E65100'); }
    
    // Progress bar via conditional formatting alternative - set progress cell
    const progCell = sheet.getRange(i + 2, 10);
    progCell.setNote(`Tiến độ: ${row[9]}%`);
  });
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  return { status: 'ok', synced: tasks.length };
}

// ============================================================
// UPDATE SINGLE TASK
// ============================================================
function updateTaskInSheet(task) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return syncAllTasks([task]);
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === task.id) {
      if (task.progress !== undefined) sheet.getRange(i+1, 10).setValue(task.progress);
      if (task.deployed !== undefined) sheet.getRange(i+1, 11).setValue(task.deployed);
      if (task.plan !== undefined) sheet.getRange(i+1, 12).setValue(task.plan);
      if (task.report !== undefined) sheet.getRange(i+1, 13).setValue(task.report);
      if (task.completedAt) sheet.getRange(i+1, 14).setValue(new Date(task.completedAt).toLocaleString('vi-VN'));
      sheet.getRange(i+1, 16).setValue(getStatus(task));
      return { status: 'ok', updated: task.id };
    }
  }
  return { status: 'notfound', id: task.id };
}

// ============================================================
// UPLOAD FILE TO GOOGLE DRIVE
// ============================================================
function uploadFileToDrive(fileData) {
  // fileData: { name, type, data (base64), taskId }
  let folder;
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
  
  // Sub-folder per task
  const taskFolderName = 'Task_' + fileData.taskId;
  let taskFolder;
  const taskFolders = folder.getFoldersByName(taskFolderName);
  if (taskFolders.hasNext()) {
    taskFolder = taskFolders.next();
  } else {
    taskFolder = folder.createFolder(taskFolderName);
  }
  
  // Decode base64 and create file
  const base64Data = fileData.data.split(',')[1];
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, fileData.type, fileData.name);
  const file = taskFolder.createFile(blob);
  
  // Make publicly readable
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return {
    status: 'ok',
    fileId: file.getId(),
    fileName: file.getName(),
    viewUrl: file.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId()
  };
}

// ============================================================
// HELPERS
// ============================================================
function getStatus(t) {
  if (t.completedAt || t.progress === 100) return 'Hoàn thành';
  const now = new Date();
  const dl = t.deadline ? new Date(t.deadline) : null;
  if (dl && dl < now) return 'Quá hạn';
  if (t.progress > 0) return 'Đang triển khai';
  return 'Chưa bắt đầu';
}
