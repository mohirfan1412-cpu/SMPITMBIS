function doGet(e) {
  var htmlOutput;
  try {
    htmlOutput = HtmlService.createHtmlOutputFromFile('index');
  } catch (err1) {
    try {
      htmlOutput = HtmlService.createHtmlOutputFromFile('index.html');
    } catch (err2) {
      return HtmlService.createHtmlOutput(
        "Gagal memuat file HTML. Pastikan file HTML di panel kiri bernama 'index' atau 'index.html'."
      );
    }
  }
  
  return htmlOutput
      .setTitle('SMART CBT')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    
    // 1. Setup Sheet Soal
    if (!ss.getSheetByName("Soal")) {
      var sheetSoal = ss.insertSheet("Soal");
      sheetSoal.appendRow(["ID Soal", "Mata Pelajaran", "Tipe", "Teks Soal", "Opsi (JSON)", "Kunci Jawaban", "Pasangan Match (JSON)", "Bobot", "URL Gambar", "Ukuran Gambar", "Posisi Gambar", "Judul Soal", "Kelas", "Pembuat"]);
      sheetSoal.getRange("A1:N1").setFontWeight("bold").setBackground("#d9ead3");
      sheetSoal.setFrozenRows(1);
    }
    
    // 2. Setup Sheet Hasil
    if (!ss.getSheetByName("Hasil")) {
      var sheetHasil = ss.insertSheet("Hasil");
      sheetHasil.appendRow(["Waktu Submit", "Nama Siswa", "Kelas", "Mata Pelajaran", "Skor CBT", "Data Jawaban (JSON)"]);
      sheetHasil.getRange("A1:F1").setFontWeight("bold").setBackground("#c9daf8");
      sheetHasil.setFrozenRows(1);
    }

    // 3. Setup Sheet Penugasan
    if (!ss.getSheetByName("Penugasan")) {
      var sheetTugas = ss.insertSheet("Penugasan");
      sheetTugas.appendRow(["ID Penugasan", "Mata Pelajaran", "Target Kelas", "Durasi (Menit)", "Batas Waktu", "Jam Mulai", "Jam Berakhir", "Judul Soal", "Kelas Soal", "Siswa Diizinkan (JSON)", "Pembuat", "Pengawas (JSON)", "Tanggal Ujian"]);
      sheetTugas.getRange("A1:M1").setFontWeight("bold").setBackground("#fff2cc");
      sheetTugas.setFrozenRows(1);
    }

    // 4. Setup Sheet Pengguna
    if (!ss.getSheetByName("Pengguna")) {
      var sheetUser = ss.insertSheet("Pengguna");
      sheetUser.appendRow(["ID User", "Username", "Password", "Role", "Nama Lengkap", "Info Tambahan"]);
      sheetUser.appendRow(["user_admin", "admin", "admin123", "admin", "Administrator", "Sekolah"]);
      sheetUser.getRange("A1:F1").setFontWeight("bold").setBackground("#f4cccc");
      sheetUser.setFrozenRows(1);
    }

    // 5. Setup Sheet Pengaturan
    if (!ss.getSheetByName("Pengaturan")) {
      var sheetSetting = ss.insertSheet("Pengaturan");
      sheetSetting.appendRow(["Key", "Value"]);
      sheetSetting.appendRow(["appName", "SMART CBT SMPIT MODERN BIS"]);
      sheetSetting.appendRow(["logoUrl", ""]);
      sheetSetting.appendRow(["showScoreToStudents", "true"]);
      sheetSetting.getRange("A1:B1").setFontWeight("bold").setBackground("#cfe2f3");
      sheetSetting.setFrozenRows(1);
    }

    // 6. Setup Sheet Monitoring (Progres Siswa Real-Time)
    if (!ss.getSheetByName("Monitoring")) {
      var sheetMon = ss.insertSheet("Monitoring");
      sheetMon.appendRow(["ID Siswa", "Nama Siswa", "Kelas", "Mata Pelajaran", "Judul Ujian", "Total Soal", "Terjawab", "Pelanggaran", "Status", "Waktu Terakhir (WITA)", "ID Penugasan", "Timestamp"]);
      sheetMon.getRange("A1:L1").setFontWeight("bold").setBackground("#d0e0e3");
      sheetMon.setFrozenRows(1);
    }
  } catch (e) {
    Logger.log("Error pada setupDatabase: " + e.message);
  }
}

// Fungsi utama: Ekstraksi berkecepatan tinggi dalam satu kali proses
function getInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { settings: { appName: "SMART CBT", logoUrl: "", showScoreToStudents: true }, questions: [], assignments: [], users: [], results: [], monitoring: [] };
    
    var sheets = ss.getSheets();
    var sheetMap = {};
    for (var s = 0; s < sheets.length; s++) {
      sheetMap[sheets[s].getName()] = sheets[s];
    }
    
    // Inisialisasi otomatis jika sheet belum dibuat
    if (!sheetMap["Soal"] || !sheetMap["Hasil"] || !sheetMap["Penugasan"] || !sheetMap["Pengguna"] || !sheetMap["Pengaturan"] || !sheetMap["Monitoring"]) {
      setupDatabase();
      sheets = ss.getSheets();
      sheetMap = {};
      for (var s = 0; s < sheets.length; s++) {
        sheetMap[sheets[s].getName()] = sheets[s];
      }
    }
    
    return {
      serverTime: new Date().getTime(),
      settings: extractSettings(sheetMap["Pengaturan"]),
      questions: extractQuestions(sheetMap["Soal"]),
      assignments: extractAssignments(sheetMap["Penugasan"]),
      users: extractUsers(sheetMap["Pengguna"]),
      results: extractResults(sheetMap["Hasil"]),
      monitoring: extractMonitoring(sheetMap["Monitoring"])
    };
  } catch (err) {
    Logger.log("Error getInitialData: " + err.message);
    return {
      serverTime: new Date().getTime(),
      settings: { appName: "SMART CBT", logoUrl: "", showScoreToStudents: true },
      questions: [],
      assignments: [],
      users: [],
      results: [],
      monitoring: []
    };
  }
}

function extractSettings(sheet) {
  if (!sheet) return { appName: "SMART CBT", logoUrl: "", showScoreToStudents: true };
  var data = sheet.getDataRange().getValues();
  var settings = { appName: "SMART CBT SMPIT MODERN BIS", logoUrl: "", showScoreToStudents: true };
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var val = data[i][1];
    if (key === 'appName') settings.appName = val ? val.toString() : "SMART CBT";
    if (key === 'logoUrl') settings.logoUrl = val ? val.toString() : "";
    if (key === 'showScoreToStudents') settings.showScoreToStudents = (val.toString() === 'true');
  }
  return settings;
}

function formatTimeString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+8", "HH:mm");
  }
  var str = val.toString().trim();
  var match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    var hh = match[1].length === 1 ? '0' + match[1] : match[1];
    return hh + ':' + match[2];
  }
  return str;
}

function formatCleanDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+8", "yyyy-MM-dd");
  }
  var str = val.toString().trim();
  var match = str.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) {
    var m = match[2].length === 1 ? '0' + match[2] : match[2];
    var d = match[3].length === 1 ? '0' + match[3] : match[3];
    return match[1] + '-' + m + '-' + d;
  }
  return str;
}

function extractQuestions(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var questions = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row[0] === '') continue;
    var opts = [];
    try { opts = row[4] ? JSON.parse(row[4]) : []; } catch (e) { opts = []; }
    var matchPairs = [];
    try { matchPairs = row[6] ? JSON.parse(row[6]) : []; } catch (e) { matchPairs = []; }
    questions.push({
      id: (row[0] || '').toString(),
      mapel: (row[1] || '-').toString(),
      type: (row[2] || 'PG').toString(),
      text: (row[3] || '').toString(),
      options: opts,
      answer: (row[5] || '').toString(),
      pairs: matchPairs,
      weight: row[7] ? Number(row[7]) : 1,
      imageUrl: row[8] ? row[8].toString() : '',
      imageSize: row[9] ? row[9].toString() : '100',
      imagePosition: row[10] ? row[10].toString() : 'atas',
      judul: row[11] ? row[11].toString() : 'Paket Utama',
      kelas: row[12] ? row[12].toString() : 'Semua Kelas',
      author: row[13] ? row[13].toString() : ''
    });
  }
  return questions;
}

function extractAssignments(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var assignments = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row[0] === '') continue;
    var allowed = [];
    try { allowed = row[9] ? JSON.parse(row[9]) : []; } catch (e) { allowed = []; }
    var proctors = [];
    try { 
      proctors = row[11] ? JSON.parse(row[11]) : []; 
    } catch (e) { 
      proctors = row[11] ? row[11].toString().split(',').map(function(s){ return s.trim(); }).filter(Boolean) : []; 
    }
    
    // Ambil tanggal ujian dari kolom 13 secara konsisten
    var rawDeadline = row[4] ? row[4].toString() : '';
    var rawTanggal = row[12] ? formatCleanDateString(row[12]) : '';
    if (!rawTanggal && rawDeadline) {
      var dateMatch = rawDeadline.match(/\d{4}-\d{2}-\d{2}/);
      if (dateMatch) rawTanggal = dateMatch[0];
    }

    assignments.push({
      id: row[0].toString(),
      mapel: (row[1] || '').toString(),
      kelas: (row[2] || '').toString(),
      duration: row[3] ? Number(row[3]) : 0,
      deadline: rawDeadline,
      jamMulai: formatTimeString(row[5]),
      jamSelesai: formatTimeString(row[6]),
      judul: row[7] ? row[7].toString() : '',
      kelasSoal: row[8] ? row[8].toString() : '',
      allowedStudents: allowed,
      author: (row[10] || '').toString(),
      proctors: proctors,
      tanggal: rawTanggal
    });
  }
  return assignments;
}

function extractUsers(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return [{
      id: "user_admin",
      username: "admin",
      password: "admin123",
      role: "admin",
      nama: "Administrator",
      info: "Admin Utama"
    }];
  }

  var users = [];
  var colMap = { id: 0, username: 1, password: 2, role: 3, nama: 4, info: 5 };
  var header = data[0];
  for (var h = 0; h < header.length; h++) {
    var hText = (header[h] || '').toString().toLowerCase().trim();
    if (hText === 'id' || hText.indexOf('id user') !== -1) colMap.id = h;
    else if (hText.indexOf('user') !== -1) colMap.username = h;
    else if (hText.indexOf('pass') !== -1 || hText.indexOf('sandi') !== -1) colMap.password = h;
    else if (hText.indexOf('role') !== -1 || hText.indexOf('peran') !== -1) colMap.role = h;
    else if (hText.indexOf('nama') !== -1) colMap.nama = h;
    else if (hText.indexOf('info') !== -1 || hText.indexOf('kelas') !== -1) colMap.info = h;
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;

    var rawUser = row[colMap.username];
    var rawPass = row[colMap.password];
    var rawNama = row[colMap.nama];
    var rawRole = row[colMap.role];
    var rawId = row[colMap.id];
    var rawInfo = row[colMap.info];

    if ((rawUser === '' || rawUser === undefined) && (rawNama === '' || rawNama === undefined) && (rawId === '' || rawId === undefined)) {
      continue;
    }

    var uUsername = (rawUser !== undefined && rawUser !== null ? rawUser.toString() : '').trim();
    var uPassword = (rawPass !== undefined && rawPass !== null ? rawPass.toString() : '').trim();
    var uNama = (rawNama !== undefined && rawNama !== null ? rawNama.toString() : uUsername).trim();
    var uRole = (rawRole ? rawRole.toString().toLowerCase() : 'siswa').trim();
    var uId = (rawId ? rawId.toString() : ('user_' + i)).trim();
    var uInfo = (rawInfo !== undefined && rawInfo !== null ? rawInfo.toString() : '').trim();

    if (!uUsername && uNama) uUsername = uNama.toLowerCase().replace(/\s+/g, '');
    if (!uNama && uUsername) uNama = uUsername;

    users.push({
      id: uId,
      username: uUsername,
      password: uPassword,
      role: uRole,
      nama: uNama,
      info: uInfo
    });
  }

  if (users.length === 0) {
    users.push({
      id: "user_admin",
      username: "admin",
      password: "admin123",
      role: "admin",
      nama: "Administrator",
      info: "Admin Utama"
    });
  }
  return users;
}

function extractResults(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row[0] === '') continue;
    var userAnswers = {};
    try { userAnswers = row[5] ? JSON.parse(row[5]) : {}; } catch (e) { userAnswers = {}; }
    results.push({
      time: row[0].toString(),
      name: (row[1] || '').toString(),
      class: (row[2] || '').toString(),
      mapel: (row[3] || '').toString(),
      score: Number(row[4]) || 0,
      answers: userAnswers
    });
  }
  return results;
}

function extractMonitoring(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row[1] === '') continue;
    list.push({
      id: (row[0] || '').toString(),
      name: (row[1] || '').toString(),
      class: (row[2] || '').toString(),
      mapel: (row[3] || '').toString(),
      judul: (row[4] || '').toString(),
      totalQuestions: Number(row[5]) || 0,
      answeredCount: Number(row[6]) || 0,
      violations: Number(row[7]) || 0,
      status: (row[8] || 'Sedang Mengerjakan').toString(),
      lastActive: (row[9] || '').toString(),
      assignmentId: (row[10] || '').toString(),
      timestamp: Number(row[11]) || 0
    });
  }
  return list;
}

function getMonitoringData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Monitoring");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Monitoring"); }
    if (!sheet) return [];
    return extractMonitoring(sheet);
  } catch (err) {
    Logger.log("Error getMonitoringData: " + err.message);
    return [];
  }
}

function updateStudentProgress(progress) {
  try {
    if (!progress || !progress.name) return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("Monitoring");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Monitoring"); }
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var cleanName = (progress.name || '').toString().toLowerCase().trim();
    var cleanAssignId = (progress.assignmentId || progress.mapel || '').toString().toLowerCase().trim();
    var foundRow = -1;

    for (var i = 1; i < data.length; i++) {
      var rowName = (data[i][1] || '').toString().toLowerCase().trim();
      var rowAssignId = (data[i][10] || data[i][3] || '').toString().toLowerCase().trim();
      if (rowName === cleanName && rowAssignId === cleanAssignId) {
        foundRow = i + 1;
        break;
      }
    }

    var rowValues = [
      progress.id || ('mon_' + Date.now()),
      progress.name,
      progress.class || '',
      progress.mapel || '',
      progress.judul || '',
      progress.totalQuestions || 0,
      progress.answeredCount || 0,
      progress.violations || 0,
      progress.status || 'Sedang Mengerjakan',
      progress.lastActive || new Date().toLocaleTimeString('id-ID'),
      progress.assignmentId || '',
      Date.now()
    ];

    if (foundRow > 1) {
      sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }
  } catch (err) {
    Logger.log("Error updateStudentProgress: " + err.message);
  }
}

function clearMonitoringData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("Monitoring");
    if (!sheet) return;
    var maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, 12).clearContent();
    }
  } catch (e) {
    Logger.log("Error clearMonitoringData: " + e.message);
  }
}

function getSettings() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { appName: "SMART CBT", logoUrl: "", showScoreToStudents: true };
    
    var sheet = ss.getSheetByName("Pengaturan");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("Pengaturan");
    }
    if (!sheet) return { appName: "SMART CBT", logoUrl: "", showScoreToStudents: true };
    
    var data = sheet.getDataRange().getValues();
    var settings = { appName: "SMART CBT SMPIT MODERN BIS", logoUrl: "", showScoreToStudents: true };
    
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0];
      var val = data[i][1];
      if (key === 'appName') settings.appName = val ? val.toString() : "SMART CBT";
      if (key === 'logoUrl') settings.logoUrl = val ? val.toString() : "";
      if (key === 'showScoreToStudents') settings.showScoreToStudents = (val.toString() === 'true');
    }
    return settings;
  } catch (err) {
    return { appName: "SMART CBT", logoUrl: "", showScoreToStudents: true };
  }
}

function saveSettings(settings) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("Pengaturan");
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName("Pengaturan");
    }
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    var foundShowScore = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === 'appName') sheet.getRange(i + 1, 2).setValue(settings.appName);
      if (data[i][0] === 'logoUrl') sheet.getRange(i + 1, 2).setValue(settings.logoUrl);
      if (data[i][0] === 'showScoreToStudents') {
        sheet.getRange(i + 1, 2).setValue(settings.showScoreToStudents ? 'true' : 'false');
        foundShowScore = true;
      }
    }

    if (!foundShowScore) {
      sheet.appendRow(["showScoreToStudents", settings.showScoreToStudents ? 'true' : 'false']);
    }
  } catch (err) {
    Logger.log("Error saveSettings: " + err.message);
  }
}

function getQuestions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Soal");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Soal"); }
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var questions = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || row[0] === '') continue;
      
      var opts = [];
      try {
        opts = row[4] ? JSON.parse(row[4]) : [];
      } catch (e) {
        opts = [];
      }

      var matchPairs = [];
      try {
        matchPairs = row[6] ? JSON.parse(row[6]) : [];
      } catch (e) {
        matchPairs = [];
      }

      questions.push({
        id: (row[0] || '').toString(),
        mapel: (row[1] || '-').toString(),
        type: (row[2] || 'PG').toString(),
        text: (row[3] || '').toString(),
        options: opts,
        answer: (row[5] || '').toString(),
        pairs: matchPairs,
        weight: row[7] ? Number(row[7]) : 1,
        imageUrl: row[8] ? row[8].toString() : '',
        imageSize: row[9] ? row[9].toString() : '100',
        imagePosition: row[10] ? row[10].toString() : 'atas',
        judul: row[11] ? row[11].toString() : 'Paket Utama',
        kelas: row[12] ? row[12].toString() : 'Semua Kelas',
        author: row[13] ? row[13].toString() : ''
      });
    }
    return questions;
  } catch (err) {
    Logger.log("Error getQuestions: " + err.message);
    return [];
  }
}

function saveQuestions(questions) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("Soal");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Soal"); }
    if (!sheet) return;
    
    var maxRows = sheet.getMaxRows();
    var maxCols = Math.max(sheet.getMaxColumns(), 14);
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
    }
    
    if (questions && questions.length > 0) {
      var rows = questions.map(function(q) {
        return [
          q.id,
          q.mapel || '-',
          q.type,
          q.text,
          JSON.stringify(q.options || []),
          q.answer,
          JSON.stringify(q.pairs || []),
          q.weight || 1,
          q.imageUrl || '',
          q.imageSize || '100',
          q.imagePosition || 'atas',
          q.judul || 'Paket Utama',
          q.kelas || 'Semua Kelas',
          q.author || ''
        ];
      });
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Error saveQuestions: " + err.message);
  }
}

function getAssignments() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Penugasan");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Penugasan"); }
    if (!sheet) return [];
    
    return extractAssignments(sheet);
  } catch (err) {
    Logger.log("Error getAssignments: " + err.message);
    return [];
  }
}

function saveAssignments(assignments) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("Penugasan");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Penugasan"); }
    if (!sheet) return;
    
    var maxRows = sheet.getMaxRows();
    var maxCols = Math.max(sheet.getMaxColumns(), 13);
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
    }
    
    if (assignments && assignments.length > 0) {
      var rows = assignments.map(function(a) {
        return [
          a.id,
          a.mapel,
          a.kelas,
          a.duration || 0,
          a.deadline || '',
          formatTimeString(a.jamMulai),
          formatTimeString(a.jamSelesai),
          a.judul || '',
          a.kelasSoal || '',
          JSON.stringify(a.allowedStudents || []),
          a.author || '',
          JSON.stringify(a.proctors || a.pengawas || []),
          formatCleanDateString(a.tanggal)
        ];
      });
      // Set kolom jam dan tanggal sebagai Plain Text agar tidak terkonversi otomatis oleh Google Sheets
      sheet.getRange(2, 6, rows.length, 2).setNumberFormat("@");
      sheet.getRange(2, 13, rows.length, 1).setNumberFormat("@");
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Error saveAssignments: " + err.message);
  }
}

function getResults() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Hasil");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Hasil"); }
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var results = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || row[0] === '') continue;
      
      var userAnswers = {};
      try {
        userAnswers = row[5] ? JSON.parse(row[5]) : {};
      } catch (e) {
        userAnswers = {};
      }

      results.push({
        time: row[0].toString(),
        name: (row[1] || '').toString(),
        class: (row[2] || '').toString(),
        mapel: (row[3] || '').toString(),
        score: Number(row[4]) || 0,
        answers: userAnswers
      });
    }
    return results;
  } catch (err) {
    Logger.log("Error getResults: " + err.message);
    return [];
  }
}

function saveResult(resultData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return false;
    var sheet = ss.getSheetByName("Hasil");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Hasil"); }
    if (!sheet) return false;
    
    // Periksa apakah data ujian siswa ini sudah ada di sheet Hasil untuk memperbarui baris atau menambah baru (mencegah duplikasi data)
    var data = sheet.getDataRange().getValues();
    var cleanName = (resultData.name || '').toString().toLowerCase().trim();
    var cleanMapel = (resultData.mapel || '').toString().toLowerCase().trim();
    var cleanClass = (resultData.class || '').toString().toLowerCase().trim();
    var cleanJudul = (resultData.judul || '').toString().toLowerCase().trim();
    var foundRow = -1;

    for (var i = 1; i < data.length; i++) {
      var rowName = (data[i][1] || '').toString().toLowerCase().trim();
      var rowClass = (data[i][2] || '').toString().toLowerCase().trim();
      var rowMapel = (data[i][3] || '').toString().toLowerCase().trim();
      if (rowName === cleanName && (!cleanClass || rowClass === cleanClass) && rowMapel === cleanMapel) {
        foundRow = i + 1;
        break;
      }
    }

    var rowValues = [
      resultData.time || Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"),
      resultData.name,
      resultData.class,
      resultData.mapel,
      resultData.score !== undefined ? resultData.score : 0,
      JSON.stringify(resultData.answers || {})
    ];

    if (foundRow > 1) {
      sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    // Sekaligus perbarui status siswa pada sheet Monitoring secara langsung (atomik) ke status Selesai
    var sheetMon = ss.getSheetByName("Monitoring");
    if (sheetMon) {
      var monData = sheetMon.getDataRange().getValues();
      var cleanAssignId = (resultData.assignmentId || resultData.mapel || '').toString().toLowerCase().trim();
      for (var j = 1; j < monData.length; j++) {
        var mName = (monData[j][1] || '').toString().toLowerCase().trim();
        var mAssignId = (monData[j][10] || monData[j][3] || '').toString().toLowerCase().trim();
        if (mName === cleanName && (mAssignId === cleanAssignId || mAssignId === cleanMapel)) {
          sheetMon.getRange(j + 1, 7).setValue(Object.keys(resultData.answers || {}).length);
          sheetMon.getRange(j + 1, 9).setValue('Selesai');
          sheetMon.getRange(j + 1, 10).setValue(Utilities.formatDate(new Date(), "GMT+8", "HH:mm") + " WITA");
          break;
        }
      }
    }

    // Paksa flush agar data langsung tersimpan di disk Google Spreadsheet tanpa delay
    SpreadsheetApp.flush();
    return true;
  } catch (err) {
    Logger.log("Error saveResult: " + err.message);
    return false;
  }
}

function saveResults(results) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("Hasil");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Hasil"); }
    if (!sheet) return;
    
    var maxRows = sheet.getMaxRows();
    var maxCols = Math.max(sheet.getMaxColumns(), 6);
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
    }
    
    if (results && results.length > 0) {
      var rows = results.map(function(r) {
        return [
          r.time || '',
          r.name || '',
          r.class || '',
          r.mapel || '',
          r.score !== undefined ? r.score : 0,
          JSON.stringify(r.answers || {})
        ];
      });
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Error saveResults: " + err.message);
  }
}

function getUsers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];
    var sheet = ss.getSheetByName("Pengguna");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Pengguna"); }
    if (!sheet) return [];
    
    return extractUsers(sheet);
  } catch (err) {
    Logger.log("Error getUsers: " + err.message);
    return [];
  }
}

function saveUsers(users) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("Pengguna");
    if (!sheet) { setupDatabase(); sheet = ss.getSheetByName("Pengguna"); }
    if (!sheet) return;
    
    var maxRows = sheet.getMaxRows();
    var maxCols = Math.max(sheet.getMaxColumns(), 6);
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
    }
    
    if (users && users.length > 0) {
      var rows = users.map(function(u) {
        return [u.id, u.username, u.password, u.role, u.nama, u.info];
      });
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log("Error saveUsers: " + err.message);
  }
}

function resetStudentExam(name, className, mapel, assignmentId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return {};
    
    var cleanName = (name || '').toString().toLowerCase().trim();
    var cleanClass = (className || '').toString().toLowerCase().trim();
    var cleanMapel = (mapel || '').toString().toLowerCase().trim();
    var cleanAssignId = (assignmentId || '').toString().toLowerCase().trim();
    
    var preservedAnswers = {};

    // 1. Ekstrak dan amankan jawaban siswa dari sheet Hasil sebelum baris nilai dihapus
    var sheetHasil = ss.getSheetByName("Hasil");
    if (sheetHasil) {
      var dataH = sheetHasil.getDataRange().getValues();
      for (var i = dataH.length - 1; i >= 1; i--) {
        var rName = (dataH[i][1] || '').toString().toLowerCase().trim();
        var rClass = (dataH[i][2] || '').toString().toLowerCase().trim();
        var rMapel = (dataH[i][3] || '').toString().toLowerCase().trim();
        if (rName === cleanName && (!cleanClass || rClass === cleanClass) && (!cleanMapel || rMapel === cleanMapel)) {
          if (dataH[i][5]) {
            try {
              var parsed = JSON.parse(dataH[i][5]);
              if (parsed && typeof parsed === 'object') {
                preservedAnswers = parsed;
              }
            } catch(e) {}
          }
          sheetHasil.deleteRow(i + 1);
        }
      }
    }
    
    // 2. Reset pelanggaran menjadi 0 dan ubah status menjadi 'Dibuka Kembali' pada sheet Monitoring
    var sheetMon = ss.getSheetByName("Monitoring");
    if (sheetMon) {
      var dataM = sheetMon.getDataRange().getValues();
      for (var j = dataM.length - 1; j >= 1; j--) {
        var mName = (dataM[j][1] || '').toString().toLowerCase().trim();
        var mAssignId = (dataM[j][10] || dataM[j][3] || '').toString().toLowerCase().trim();
        if (mName === cleanName && (!cleanAssignId || mAssignId === cleanAssignId || mAssignId === cleanMapel)) {
          // Kolom 8 adalah Pelanggaran, Kolom 9 adalah Status
          sheetMon.getRange(j + 1, 8).setValue(0);
          sheetMon.getRange(j + 1, 9).setValue('Dibuka Kembali');
        }
      }
    }
    SpreadsheetApp.flush();
    return preservedAnswers;
  } catch (e) {
    Logger.log("Error resetStudentExam: " + e.message);
    return {};
  }
}

function uploadImageToDrive(base64Data, fileName, mimeType) {
  try {
    var folderName = "Gambar_CBT";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    var data = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(data, mimeType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var directUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
    return directUrl;
  } catch (e) {
    return "ERROR: " + e.message;
  }
}
