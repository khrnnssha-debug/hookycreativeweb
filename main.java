/**
 * HOOKY CREATIVE AGENCY - DESKTOP BACKEND & DATABASE SETUP
 */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet 1: Users
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['id', 'username', 'password', 'name', 'avatar']);
    userSheet.appendRow(['usr_1', 'admin', 'hooky123', 'Admin Hooky', 'https://images.unsplash.com/photo-1534528741775-53994a69daeb?w=150']);
  }
  
  // Sheet 2: Schedule (Jadwal Upload)
  let scheduleSheet = ss.getSheetByName('Schedule');
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet('Schedule');
    scheduleSheet.appendRow(['id', 'title', 'platform', 'category', 'date', 'time', 'pic', 'status', 'thumbnail', 'mediaType', 'mediaUrl', 'postUrl', 'views', 'likes', 'comments', 'shares', 'saves']);
    const today = new Date().toISOString().split('T')[0];
    scheduleSheet.appendRow(['sch_1', 'Tips Financial Client 2026', 'Instagram', 'Edukasi', today, '18:00', 'Budi', 'Belum Diupload', 'https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?w=300', 'image', 'https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?w=800', '', 0, 0, 0, 0, 0]);
  }
  
  // Sheet 3: Drafts (Draft Konten dengan Media)
  let draftSheet = ss.getSheetByName('Drafts');
  if (!draftSheet) {
    draftSheet = ss.insertSheet('Drafts');
    draftSheet.appendRow(['id', 'title', 'concept', 'hook', 'caption', 'hashtag', 'platform', 'deadline', 'status', 'mediaType', 'mediaUrl']);
    const today = new Date().toISOString().split('T')[0];
    draftSheet.appendRow(['drf_1', 'Ide Campaign Klien A', 'Review produk baru', 'Bisa dapet 10jt pertama?', 'Detail promosi...', '#hooky #campaign', 'Instagram', today, 'Scripting', 'image', 'https://images.unsplash.com/photo-1618005182384-a83a8bd57fbe?w=600']);
  }
  
  // Sheet 4: References
  let refSheet = ss.getSheetByName('References');
  if (!refSheet) {
    refSheet = ss.insertSheet('References');
    refSheet.appendRow(['id', 'title', 'link', 'thumbnail', 'platform', 'tag', 'whyInteresting', 'isFavorite']);
    refSheet.appendRow(['ref_1', 'Transisi Video Aesthetic', 'https://instagram.com/p/example1', 'https://images.unsplash.com/photo-1516035069371-29a1b244cc32?w=300', 'Instagram', 'Editing', 'Peralihan mulus', true]);
  }
  
  return "Database Setup Berhasil!";
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Hooky Creative Agency Workspace')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    users: sheetToJSON(ss.getSheetByName('Users')),
    schedules: sheetToJSON(ss.getSheetByName('Schedule')),
    drafts: sheetToJSON(ss.getSheetByName('Drafts')),
    references: sheetToJSON(ss.getSheetByName('References'))
  };
}

function sheetToJSON(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = val.toISOString().split('T')[0];
      }
      obj[h] = val;
    });
    return obj;
  });
}

function saveItem(sheetName, item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  if (!item.id) {
    item.id = sheetName.toLowerCase().substring(0, 3) + '_' + Date.now();
  }

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == item.id) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowValues = headers.map(h => item[h] !== undefined ? item[h] : '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  
  return getAppData();
}

function deleteItem(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getAppData();
}