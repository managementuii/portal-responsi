// ==========================================
// VARIABEL GLOBAL UNTUK FORMAT TANGGAL
// ==========================================
var NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
var NAMA_HARI = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

function toIndoDateString(yyyy_mm_dd) {
  if (!yyyy_mm_dd) return "";
  var parts = yyyy_mm_dd.split("-");
  if (parts.length !== 3) return yyyy_mm_dd;
  var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  return NAMA_HARI[d.getDay()] + ", " + d.getDate() + " " + NAMA_BULAN[d.getMonth()] + " " + d.getFullYear();
}

function fromIndoDateString(indoStr) {
  if (!indoStr || typeof indoStr !== "string") return indoStr;
  var cleanStr = indoStr.replace(/,/g, "").trim();
  var parts = cleanStr.split(/\s+/);
  if (parts.length < 3) return indoStr;
  var tglStr = parts.length >= 4 ? parts[1] : parts[0];
  var blnStr = parts.length >= 4 ? parts[2] : parts[1];
  var thnStr = parts.length >= 4 ? parts[3] : parts[2];
  
  var blnIdx = NAMA_BULAN.indexOf(blnStr);
  if (blnIdx !== -1) {
    return thnStr + "-" + (blnIdx + 1).toString().padStart(2, '0') + "-" + tglStr.padStart(2, '0');
  }
  return indoStr;
}

function formatWA(wa) {
  if (!wa) return "";
  var str = wa.toString().replace(/[^0-9]/g, '');
  if (str.startsWith("0")) return "62" + str.substring(1);
  if (str.startsWith("8")) return "62" + str;
  return str;
}


// ==========================================
// OTOMATISASI KAMUS HEADER (PENGAMAN STRUKTUR KOLOM)
// ==========================================
function getKamusKolom(sheet) {
  if (sheet.getLastColumn() === 0) return {};

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  function cari(namaJudul, indeksDefault) {
    var indeks = headers.indexOf(namaJudul);
    return indeks !== -1 ? indeks : indeksDefault;
  }

  return {
    NO: cari("No", 0),
    TANGGAL: cari("Tanggal", 1),
    JAM: cari("Jam", 2),
    NIM: cari("NIM", 3),
    NAMA_MHS: cari("Nama Mahasiswa", 4),
    WA_MHS: cari("No. WA Mahasiswa", 5),
    PERUSAHAAN: cari("Perusahaan", 6),
    NAMA_SPV: cari("Nama Supervisor", 7),
    WA_SPV: cari("No. WA Supervisor", 8),
    DOSBING: cari("Dosen Pembimbing", 10),
    WA_DOSEN: cari("No. WA Dosen", 11),
    PIC: cari("PIC Teknis", 13),
    WA_PIC: cari("WA PIC", 14),
    STATUS: cari("Status Responsi", 15),
    EMAIL_SPV: cari("Email SPV", 16),
    EMAIL_SITA: cari("Email SITA", 17),
    KEHADIRAN_SPV: cari("Kehadiran SPV", 18),
    NAMA_PENGGANTI: cari("Pengganti SPV", 19),
    WA_PENGGANTI: cari("WA Pengganti", 20),
    EMAIL_PENGGANTI: cari("Email Pengganti", 21),
    CENTANG_W: cari("Kolom W", 22),       
    CENTANG_X: cari("Kolom X", 23),       
    ALAMAT: cari("Alamat Perusahaan", 24),
    ID_SPV: cari("ID SPV", 25),           
    ID_DOSEN: cari("ID Dosen", 26),       
    STATUS_SURAT: cari("Status Surat", 27),
    // --- PENAMBAHAN KOLOM EVALUASI BARU ---
    JABATAN_PENGGANTI: cari("Jabatan Pengganti", 28),
    SPV_BERDAMPAK: cari("SPV Berdampak", 29),
    JABATAN_SPV_BERDAMPAK: cari("Jabatan SPV Berdampak", 30),
    // --- KOLOM FLAG NOTIFIKASI PIC ---
    STATUS_NOTIF: cari("Status Notifikasi", 31)
  };
}


// ==========================================
// 1. ROUTING MULTI-PAGE DENGAN IZIN IFRAME
// ==========================================
function doGet(e) {
  var params = Object.keys(e.parameter);

  if (e.parameter.page == 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
        .setTitle('Admin Panel - Portal Magang')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (e.parameter.page == 'pic-123' || params.includes('pic-123')) {
    return HtmlService.createHtmlOutputFromFile('pic-123')
        .setTitle('Command Center PIC - UII')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  for (var i = 0; i < params.length; i++) {
    var p = params[i].toLowerCase();
    if (p.startsWith('dos') || p.startsWith('spv')) {
      var template = HtmlService.createTemplateFromFile('Portal');
      template.targetId = p; 
      
      return template.evaluate()
          .setTitle('Jadwal Responsi Magang - UII')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Portal Administrasi Magang')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 2. FUNGSI PENGATURAN FASE
// ==========================================
function getPhaseStatus() {
  var props = PropertiesService.getScriptProperties();
  var phase = props.getProperty('ACTIVE_PHASE');
  if (!phase) { props.setProperty('ACTIVE_PHASE', '1'); return 1; }
  return parseInt(phase);
}

function setPhaseStatus(phaseNumber) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ACTIVE_PHASE', phaseNumber.toString());
  return parseInt(phaseNumber);
}

// ==========================================
// 3. FUNGSI MENGAMBIL DATA (OPTIMIZED)
// ==========================================
function getPreviewData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet); 
  var rawData = sheet.getDataRange().getValues(); // SUPER CEPAT
  var previewData = [];
 
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][KOLOM.NAMA_MHS] !== "") { 
      var rawTanggal = rawData[i][KOLOM.TANGGAL]; 
      var rawJam = rawData[i][KOLOM.JAM]; 
      var timestamp = 0;
      var tglInput = "";
      var tglDisplay = "";
      var jamDisplay = "";
      var dateObj = null;

      if (rawJam !== "") jamDisplay = (rawJam instanceof Date) ? Utilities.formatDate(rawJam, Session.getScriptTimeZone(), "HH:mm") : rawJam.toString();

      if (rawTanggal && rawTanggal !== "") {
        if (rawTanggal instanceof Date) {
          dateObj = new Date(rawTanggal.getTime());
          tglInput = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
          tglDisplay = toIndoDateString(tglInput);
        } else {
          tglInput = fromIndoDateString(rawTanggal.toString());
          tglDisplay = rawTanggal.toString();
          dateObj = new Date(tglInput);
          if (isNaN(dateObj.getTime())) dateObj = null;
        }
      }

      if (dateObj) {
        if (jamDisplay && jamDisplay.includes(':')) {
          var timeParts = jamDisplay.split(':');
          dateObj.setHours(parseInt(timeParts[0], 10), parseInt(timeParts[1], 10), 0, 0);
        }
        timestamp = dateObj.getTime();
      }

      var isWChecked = rawData[i][KOLOM.CENTANG_W] === true || String(rawData[i][KOLOM.CENTANG_W]).toUpperCase() === "TRUE"; 
      var isXChecked = rawData[i][KOLOM.CENTANG_X] === true || String(rawData[i][KOLOM.CENTANG_X]).toUpperCase() === "TRUE"; 
      var finalStatus = "Belum Daftar Ulang"; 

      if (timestamp > 0) {
        if (isWChecked && isXChecked) finalStatus = "Responsi Selesai";
        else if (isWChecked && !isXChecked) finalStatus = "Belum Isi Form Pasca Responsi";
        else finalStatus = "Belum Responsi";
      }

      var picTeknis = rawData[i][KOLOM.PIC] ? rawData[i][KOLOM.PIC].toString() : "-";
      var waPicRaw = rawData[i][KOLOM.WA_PIC] || "";
      var waMhsRaw = rawData[i][KOLOM.WA_MHS] || "";
      var waSpvRaw = rawData[i][KOLOM.WA_SPV] || "";
      var waDosenRaw = rawData[i][KOLOM.WA_DOSEN] || "";
      
      previewData.push({ 
        row: i + 1,
        timestamp: timestamp, 
        tanggal: tglDisplay, 
        tanggalInput: tglInput,
        jam: jamDisplay, 
        nim: rawData[i][KOLOM.NIM] ? rawData[i][KOLOM.NIM].toString() : "-",
        nama: rawData[i][KOLOM.NAMA_MHS].toString(), 
        perusahaan: rawData[i][KOLOM.PERUSAHAAN] ? rawData[i][KOLOM.PERUSAHAAN].toString() : "-",
        spv: rawData[i][KOLOM.NAMA_SPV] ? rawData[i][KOLOM.NAMA_SPV].toString() : "-",
        dosbing: rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString() : "-",
        waMhs: waMhsRaw !== "" ? formatWA(waMhsRaw) : "",
        waSpv: waSpvRaw !== "" ? formatWA(waSpvRaw) : "",
        waDosen: waDosenRaw !== "" ? formatWA(waDosenRaw) : "",
        pic: picTeknis, 
        waPic: waPicRaw !== "" ? formatWA(waPicRaw) : "", 
        status: finalStatus,
        isWChecked: isWChecked 
      });
    }
  }
  
  previewData.sort(function(a, b) {
    if (a.timestamp > 0 && b.timestamp === 0) return -1;
    if (a.timestamp === 0 && b.timestamp > 0) return 1; 
    if (a.timestamp > 0 && b.timestamp > 0) return a.timestamp - b.timestamp; 
    return 0; 
  });
  return previewData;
}

function getPhase1Data() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var rawData = sheet.getDataRange().getValues(); // SUPER CEPAT
  var phase1Data = [];
 
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][KOLOM.NAMA_MHS] !== "") { 
      var statusKolomAB = rawData[i][KOLOM.STATUS_SURAT];
      phase1Data.push({ 
        nama: rawData[i][KOLOM.NAMA_MHS] ? rawData[i][KOLOM.NAMA_MHS].toString() : "-", 
        dosbing: rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString() : "-",
        perusahaan: rawData[i][KOLOM.PERUSAHAAN] ? rawData[i][KOLOM.PERUSAHAAN].toString() : "-", 
        spv: rawData[i][KOLOM.NAMA_SPV] ? rawData[i][KOLOM.NAMA_SPV].toString() : "-", 
        status: (statusKolomAB && statusKolomAB.toString().trim() !== "") ? "Sudah Konfirmasi" : "Belum Konfirmasi" 
      });
    }
  }
  phase1Data.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return phase1Data;
}


// ==========================================
// PENGAMAN 1: CLASH CHECKER DOSEN (PINTU DEPAN)
// ==========================================
function cekBentrokDosen(sheet, KOLOM, nimToIgnore, dosbing, tglIndoStr, jamStr) {
  if (!dosbing || dosbing === "" || dosbing === "-") return false;
  
  var rawData = sheet.getDataRange().getValues();
  for (var i = 1; i < rawData.length; i++) {
    // 1. Abaikan baris mahasiswa itu sendiri
    if (rawData[i][KOLOM.NIM] == nimToIgnore) continue;
    
    // 2. Abaikan jika dosennya berbeda
    var rDosen = rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString().trim() : "";
    if (rDosen !== dosbing.trim()) continue;
    
    // 3. Cek Tanggal & Jam
    var rTgl = rawData[i][KOLOM.TANGGAL];
    var rawTglStr = "";
    if (rTgl !== "") {
      rawTglStr = (rTgl instanceof Date) ? toIndoDateString(Utilities.formatDate(rTgl, Session.getScriptTimeZone(), "yyyy-MM-dd")) : rTgl.toString().trim();
    }
    
    var rJam = rawData[i][KOLOM.JAM];
    var rawJamStr = "";
    if (rJam !== "") {
      rawJamStr = (rJam instanceof Date) ? Utilities.formatDate(rJam, Session.getScriptTimeZone(), "HH:mm") : rJam.toString().trim();
    }
    
    // BENTROK TERDETEKSI: Dosen sama, Hari sama, Jam sama!
    if (rawTglStr === tglIndoStr && rawJamStr === jamStr) {
      return rawData[i][KOLOM.NAMA_MHS]; // Kembalikan nama mahasiswa lawannya
    }
  }
  return false; // Aman, tidak ada bentrok
}


// ==========================================
// PENGAMAN 2: ALGORITMA AUTO-ASSIGN PIC (MESIN BELAKANG)
// ==========================================
function autoAssignPIC(tanggalRequestStr, jamRequestStr, namaDosen, picLama) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMain = ss.getSheets()[0];
  var sheetPic = ss.getSheetByName("PIC Teknis");
  
  if (!sheetPic) return {nama: "Menunggu Admin", wa: ""};

  var picData = sheetPic.getDataRange().getValues();
  var masterPics = [];
  for(var i = 1; i < picData.length; i++) {
      if(picData[i][1]) {
          masterPics.push({
              nama: picData[i][1].toString().trim(), 
              wa: picData[i][2] ? picData[i][2].toString().trim() : ""
          });
      }
  }
  if (masterPics.length === 0) return {nama: "Menunggu Admin", wa: ""};

  var rawData = sheetMain.getDataRange().getValues();
  var KOLOM = getKamusKolom(sheetMain);

  var busyPics = [];          
  var dosenTally = {};        
  var workloadTally = {};     
  
  masterPics.forEach(function(p) { workloadTally[p.nama] = 0; });

  for (var j = 1; j < rawData.length; j++) {
      var rTgl = rawData[j][KOLOM.TANGGAL];
      var rJam = rawData[j][KOLOM.JAM];
      var rDosen = rawData[j][KOLOM.DOSBING] ? rawData[j][KOLOM.DOSBING].toString().trim() : "";
      var rPic = rawData[j][KOLOM.PIC] ? rawData[j][KOLOM.PIC].toString().trim() : "";

      if (!rPic || rPic === "-" || rPic === "Menunggu Admin") continue;

      var rTglStr = "";
      if (rTgl && rTgl !== "") rTglStr = (rTgl instanceof Date) ? toIndoDateString(Utilities.formatDate(rTgl, Session.getScriptTimeZone(), "yyyy-MM-dd")) : rTgl.toString().trim();
      
      var rJamStr = "";
      if (rJam && rJam !== "") rJamStr = (rJam instanceof Date) ? Utilities.formatDate(rJam, Session.getScriptTimeZone(), "HH:mm") : rJam.toString().trim();

      if (workloadTally[rPic] !== undefined) workloadTally[rPic]++;

      if (rTglStr === tanggalRequestStr && rJamStr === jamRequestStr) {
          busyPics.push(rPic);
      }

      if (rDosen === namaDosen) {
          dosenTally[rPic] = (dosenTally[rPic] || 0) + 1;
      }
  }

  function getWaPIC(namaPIC) {
      for (var k=0; k < masterPics.length; k++) {
          if (masterPics[k].nama === namaPIC) return masterPics[k].wa;
      }
      return "";
  }

  // 1. Prioritas PIC Lama
  if (picLama && picLama !== "-" && picLama !== "Menunggu Admin") {
      var picLamaTerdaftar = false;
      for (var k=0; k < masterPics.length; k++) { if (masterPics[k].nama === picLama) picLamaTerdaftar = true; }
      
      if (picLamaTerdaftar && busyPics.indexOf(picLama) === -1) {
          return { nama: picLama, wa: getWaPIC(picLama) }; 
      }
  }

  // 2. Auto-Learning Dosen
  var favoritePic = "";
  var maxAssisted = 0;
  for (var p in dosenTally) {
      if (dosenTally[p] > maxAssisted) {
          maxAssisted = dosenTally[p];
          favoritePic = p;
      }
  }
  
  if (favoritePic !== "") {
      var favTerdaftar = false;
      for (var k=0; k < masterPics.length; k++) { if (masterPics[k].nama === favoritePic) favTerdaftar = true; }
      
      if (favTerdaftar && busyPics.indexOf(favoritePic) === -1) {
          return { nama: favoritePic, wa: getWaPIC(favoritePic) }; 
      }
  }

  // 3. Load Balancing
  var availablePics = [];
  for (var i=0; i < masterPics.length; i++) {
      if (busyPics.indexOf(masterPics[i].nama) === -1) {
          availablePics.push(masterPics[i]);
      }
  }

  if (availablePics.length > 0) {
      availablePics.sort(function(a, b) {
          return workloadTally[a.nama] - workloadTally[b.nama];
      });
      return { nama: availablePics[0].nama, wa: availablePics[0].wa }; 
  }

  // 4. Skenario Terburuk
  return { nama: "Menunggu Admin", wa: "" };
}


// ==========================================
// 5. FUNGSI CARI & SIMPAN DATA (MAHASISWA & PIC)
// ==========================================
function updateByPIC(info) {
  if (!info || !info.row) return { status: "error", message: "Gagal: Data baris tidak valid." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
    var KOLOM = getKamusKolom(sheet);
    var row = info.row;

    if (info.tanggal) sheet.getRange(row, KOLOM.TANGGAL + 1).setValue(toIndoDateString(info.tanggal)); 
    if (info.jam) sheet.getRange(row, KOLOM.JAM + 1).setValue(info.jam);        
    if (info.spv) sheet.getRange(row, KOLOM.NAMA_SPV + 1).setValue(info.spv);
    if (info.waMhs) sheet.getRange(row, KOLOM.WA_MHS + 1).setValue(formatWA(info.waMhs)); 
    if (info.waSpv) sheet.getRange(row, KOLOM.WA_SPV + 1).setValue(formatWA(info.waSpv));      
    if (info.waDosen !== undefined) sheet.getRange(row, KOLOM.WA_DOSEN + 1).setValue(formatWA(info.waDosen));

    if (info.isWChecked !== undefined) {
        sheet.getRange(row, KOLOM.CENTANG_W + 1).setValue(info.isWChecked ? true : false);
    }

    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#fef9c3');
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function getDataByNIM(nim) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var data = sheet.getDataRange().getValues();
 
  for (var i = 1; i < data.length; i++) {
    if (data[i][KOLOM.NIM] == nim) {
      var tglFormat = (data[i][KOLOM.TANGGAL] instanceof Date) ? Utilities.formatDate(data[i][KOLOM.TANGGAL], Session.getScriptTimeZone(), "yyyy-MM-dd") : fromIndoDateString(data[i][KOLOM.TANGGAL].toString());
      var jamFormat = (data[i][KOLOM.JAM] instanceof Date) ? Utilities.formatDate(data[i][KOLOM.JAM], Session.getScriptTimeZone(), "HH:mm") : data[i][KOLOM.JAM];

      var isWChecked = data[i][KOLOM.CENTANG_W] === true || String(data[i][KOLOM.CENTANG_W]).toUpperCase() === "TRUE"; 
      var isXChecked = data[i][KOLOM.CENTANG_X] === true || String(data[i][KOLOM.CENTANG_X]).toUpperCase() === "TRUE"; 
      
      var rawWaMhs = data[i][KOLOM.WA_MHS] ? data[i][KOLOM.WA_MHS].toString() : ""; if(rawWaMhs.startsWith("62")) rawWaMhs = "0" + rawWaMhs.substring(2); else if(rawWaMhs.startsWith("8")) rawWaMhs = "0" + rawWaMhs;
      var rawWaSpv = data[i][KOLOM.WA_SPV] ? data[i][KOLOM.WA_SPV].toString() : ""; if(rawWaSpv.startsWith("62")) rawWaSpv = "0" + rawWaSpv.substring(2); else if(rawWaSpv.startsWith("8")) rawWaSpv = "0" + rawWaSpv;
      var rawWaPengganti = data[i][KOLOM.WA_PENGGANTI] ? data[i][KOLOM.WA_PENGGANTI].toString() : ""; if(rawWaPengganti.startsWith("62")) rawWaPengganti = "0" + rawWaPengganti.substring(2); else if(rawWaPengganti.startsWith("8")) rawWaPengganti = "0" + rawWaPengganti;
      var rawWaPic = data[i][KOLOM.WA_PIC] ? data[i][KOLOM.WA_PIC].toString() : ""; if(rawWaPic.startsWith("62")) rawWaPic = "0" + rawWaPic.substring(2); else if(rawWaPic.startsWith("8")) rawWaPic = "0" + rawWaPic;

      return {
        isNew: false, isFinished: (isWChecked && isXChecked), row: i + 1, nim: nim, nama: data[i][KOLOM.NAMA_MHS], waMhs: rawWaMhs,    
        tanggal: tglFormat, jam: jamFormat, perusahaan: data[i][KOLOM.PERUSAHAAN], spv: data[i][KOLOM.NAMA_SPV], waSpv: rawWaSpv, dosbing: data[i][KOLOM.DOSBING],        
        emailSpv: data[i][KOLOM.EMAIL_SPV] || "", emailSita: data[i][KOLOM.EMAIL_SITA] || "", kehadiranSpv: data[i][KOLOM.KEHADIRAN_SPV] || "", namaPengganti: data[i][KOLOM.NAMA_PENGGANTI] || "",
        waPengganti: rawWaPengganti, emailPengganti: data[i][KOLOM.EMAIL_PENGGANTI] || "", alamatPerusahaan: data[i][KOLOM.ALAMAT] || "", pic: data[i][KOLOM.PIC] || "-", waPic: rawWaPic
      };
    }
  }
  return { isNew: true, isFinished: false, nim: nim };
}

function simpanUpdate(info) {
  if (!info || !info.nim) return { status: "error", message: "Gagal: Data kosong." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
    var KOLOM = getKamusKolom(sheet);
    var row = info.row;
    
    var isWChecked = sheet.getRange(row, KOLOM.CENTANG_W + 1).getValue() === true || String(sheet.getRange(row, KOLOM.CENTANG_W + 1).getValue()).toUpperCase() === "TRUE"; 
    var isXChecked = sheet.getRange(row, KOLOM.CENTANG_X + 1).getValue() === true || String(sheet.getRange(row, KOLOM.CENTANG_X + 1).getValue()).toUpperCase() === "TRUE"; 
    if (isWChecked && isXChecked) return { status: "error", message: "Responsi telah selesai." };

    var tglIndoStr = toIndoDateString(info.tanggal);

    // ===============================================
    // 🛡️ FRONT-DOOR VALIDATION: CEK BENTROK DOSEN
    // ===============================================
    var mhsBentrok = cekBentrokDosen(sheet, KOLOM, info.nim, info.dosbing, tglIndoStr, info.jam);
    if (mhsBentrok) {
      return { 
        status: "error", 
        message: "Jadwal Bentrok! Dosen pembimbing ini sudah dijadwalkan untuk menguji mahasiswa bernama '" + mhsBentrok + "' pada hari dan jam tersebut. Silakan pilih waktu yang lain." 
      };
    }

    // BACA PIC SAAT INI SEBELUM UPDATE
    var picSaatIni = sheet.getRange(row, KOLOM.PIC + 1).getValue() ? sheet.getRange(row, KOLOM.PIC + 1).getValue().toString().trim() : "";
    
    // TEMBAK MESIN AUTO-ASSIGN
    var assignedPIC = autoAssignPIC(tglIndoStr, info.jam, info.dosbing, picSaatIni);

    sheet.getRange(row, KOLOM.TANGGAL + 1).setValue(tglIndoStr); 
    sheet.getRange(row, KOLOM.JAM + 1).setValue(info.jam);        
    sheet.getRange(row, KOLOM.WA_MHS + 1).setValue(formatWA(info.waMhs)); 
    sheet.getRange(row, KOLOM.PERUSAHAAN + 1).setValue(info.perusahaan); 
    sheet.getRange(row, KOLOM.NAMA_SPV + 1).setValue(info.spv); 
    sheet.getRange(row, KOLOM.WA_SPV + 1).setValue(formatWA(info.waSpv));      
    sheet.getRange(row, KOLOM.STATUS + 1).setValue("Belum Responsi"); 
    sheet.getRange(row, KOLOM.EMAIL_SPV + 1).setValue(info.emailSpv);  
    sheet.getRange(row, KOLOM.EMAIL_SITA + 1).setValue(info.emailSita); 
    sheet.getRange(row, KOLOM.KEHADIRAN_SPV + 1).setValue(info.kehadiranSpv);
    sheet.getRange(row, KOLOM.NAMA_PENGGANTI + 1).setValue(info.namaPengganti); 
    sheet.getRange(row, KOLOM.WA_PENGGANTI + 1).setValue(formatWA(info.waPengganti));
    sheet.getRange(row, KOLOM.EMAIL_PENGGANTI + 1).setValue(info.emailPengganti);

    // TERAPKAN HASIL KEPUTUSAN AUTO-ASSIGN
    sheet.getRange(row, KOLOM.PIC + 1).setValue(assignedPIC.nama);
    sheet.getRange(row, KOLOM.WA_PIC + 1).setValue(formatWA(assignedPIC.wa));
    
    // 🔥 TAMBAH FLAG NOTIFIKASI UNTUK PIC 🔥
    sheet.getRange(row, KOLOM.STATUS_NOTIF + 1).setValue("BARU");
    
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e0f2fe');
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function simpanBaru(info) {
  if (!info || !info.nim) return { status: "error", message: "Gagal: Data kosong." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
    var KOLOM = getKamusKolom(sheet);
    var newRow = sheet.getLastRow() + 1;
    var maxCol = Math.max(31, sheet.getLastColumn()); // Pastikan mengcover indeks baru
    var rowData = new Array(maxCol).fill(""); 

    var tglIndoStr = toIndoDateString(info.tanggal);

    // ===============================================
    // 🛡️ FRONT-DOOR VALIDATION: CEK BENTROK DOSEN
    // ===============================================
    var mhsBentrok = cekBentrokDosen(sheet, KOLOM, info.nim, info.dosbing, tglIndoStr, info.jam);
    if (mhsBentrok) {
      return { 
        status: "error", 
        message: "Jadwal Bentrok! Dosen pembimbing ini sudah dijadwalkan untuk menguji mahasiswa bernama '" + mhsBentrok + "' pada hari dan jam tersebut. Silakan pilih waktu yang lain." 
      };
    }

    // TEMBAK MESIN AUTO-ASSIGN (Kosongkan Pic Lama karena ini input baru)
    var assignedPIC = autoAssignPIC(tglIndoStr, info.jam, info.dosbing, "");
    
    rowData[KOLOM.NO] = newRow - 1; 
    rowData[KOLOM.TANGGAL] = tglIndoStr; 
    rowData[KOLOM.JAM] = info.jam; 
    rowData[KOLOM.NIM] = info.nim;                       
    rowData[KOLOM.NAMA_MHS] = info.nama; 
    rowData[KOLOM.WA_MHS] = formatWA(info.waMhs); 
    rowData[KOLOM.PERUSAHAAN] = info.perusahaan; 
    rowData[KOLOM.NAMA_SPV] = info.spv;                       
    rowData[KOLOM.WA_SPV] = formatWA(info.waSpv); 
    rowData[KOLOM.DOSBING] = info.dosbing; 
    rowData[KOLOM.STATUS] = "Belum Responsi"; 
    rowData[KOLOM.EMAIL_SPV] = info.emailSpv;                 
    rowData[KOLOM.EMAIL_SITA] = info.emailSita; 
    rowData[KOLOM.KEHADIRAN_SPV] = info.kehadiranSpv; 
    rowData[KOLOM.NAMA_PENGGANTI] = info.namaPengganti; 
    rowData[KOLOM.WA_PENGGANTI] = formatWA(info.waPengganti);
    rowData[KOLOM.EMAIL_PENGGANTI] = info.emailPengganti;

    // TERAPKAN HASIL KEPUTUSAN AUTO-ASSIGN & FLAG NOTIFIKASI
    rowData[KOLOM.PIC] = assignedPIC.nama;
    rowData[KOLOM.WA_PIC] = formatWA(assignedPIC.wa);
    rowData[KOLOM.STATUS_NOTIF] = "BARU";
    
    sheet.getRange(newRow, 1, 1, maxCol).setValues([rowData]); 
    sheet.getRange(newRow, 1, 1, maxCol).setBackground('#ecfdf5');
    
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function simpanSurat(info) {
  if (!info || !info.nim) return { status: "error", message: "Gagal: NIM kosong." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
    var KOLOM = getKamusKolom(sheet);
    var row = info.row;
    
    if (sheet.getRange(row, KOLOM.NIM + 1).getValue() != info.nim) return { status: "error", message: "Akses Ditolak: Data tidak sinkron." };
    
    sheet.getRange(row, KOLOM.WA_MHS + 1).setValue(formatWA(info.waMhs)); 
    sheet.getRange(row, KOLOM.PERUSAHAAN + 1).setValue(info.perusahaan);      
    sheet.getRange(row, KOLOM.NAMA_SPV + 1).setValue(info.spv); 
    sheet.getRange(row, KOLOM.WA_SPV + 1).setValue(formatWA(info.waSpv));              
    sheet.getRange(row, KOLOM.EMAIL_SPV + 1).setValue(info.emailSpv); 
    sheet.getRange(row, KOLOM.EMAIL_SITA + 1).setValue(info.emailSita);      
    sheet.getRange(row, KOLOM.ALAMAT + 1).setValue(info.alamat); 
    sheet.getRange(row, KOLOM.STATUS_SURAT + 1).setValue("Sudah Konfirmasi");  
    
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function getDropdownOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
  var KOLOM = getKamusKolom(sheet);
  var rawData = sheet.getDataRange().getValues();
  var companies = []; var spvs = [];
  
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][KOLOM.PERUSAHAAN] && rawData[i][KOLOM.PERUSAHAAN] !== "" && companies.indexOf(rawData[i][KOLOM.PERUSAHAAN]) === -1) {
      companies.push(rawData[i][KOLOM.PERUSAHAAN]);
    }
    if (rawData[i][KOLOM.NAMA_SPV] && rawData[i][KOLOM.NAMA_SPV] !== "" && spvs.indexOf(rawData[i][KOLOM.NAMA_SPV]) === -1) {
      spvs.push(rawData[i][KOLOM.NAMA_SPV]);
    }
  }
  companies.sort(); spvs.sort();
  return { companies: companies, spvs: spvs };
}


// ==========================================
// 6. FUNGSI PORTAL PERSONALISASI (DOSEN & SPV)
// ==========================================
function getJadwalPersonalisasi(targetId) {
  if (!targetId) return { isFound: false };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var rawData = sheet.getDataRange().getValues(); 

  var isDosen = targetId.toLowerCase().startsWith('dos');
  var idColumnIndex = isDosen ? KOLOM.ID_DOSEN : KOLOM.ID_SPV;

  var result = {
    isFound: false,
    role: isDosen ? "Dosen Pembimbing" : "Supervisor",
    name: "",
    isDosen: isDosen,
    company: "-",
    email: "-",
    data: []
  };

  for (var i = 1; i < rawData.length; i++) {
    var currentId = rawData[i][idColumnIndex] ? rawData[i][idColumnIndex].toString().trim().toLowerCase() : "";

    if (currentId === targetId.toLowerCase()) {
      
      if (!result.isFound) {
        result.isFound = true;
        if (isDosen) {
          result.name = rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString() : "-"; 
        } else {
          result.name = rawData[i][KOLOM.NAMA_SPV] ? rawData[i][KOLOM.NAMA_SPV].toString() : "-";  
          result.company = rawData[i][KOLOM.PERUSAHAAN] ? rawData[i][KOLOM.PERUSAHAAN].toString() : "-"; 
          result.email = rawData[i][KOLOM.EMAIL_SPV] ? rawData[i][KOLOM.EMAIL_SPV].toString() : "-";  
        }
      }

      var isWChecked = rawData[i][KOLOM.CENTANG_W] === true || String(rawData[i][KOLOM.CENTANG_W]).toUpperCase() === "TRUE";
      var isXChecked = rawData[i][KOLOM.CENTANG_X] === true || String(rawData[i][KOLOM.CENTANG_X]).toUpperCase() === "TRUE";
      var finalStatus = "Belum Responsi";

      if (isWChecked && isXChecked) {
        finalStatus = "Responsi Selesai";
      } else if (isWChecked && !isXChecked) {
        finalStatus = "Belum Isi Form"; 
      }

      var rawTanggal = rawData[i][KOLOM.TANGGAL]; 
      var rawJam = rawData[i][KOLOM.JAM]; 
      var tglDisplay = "Belum ditentukan";
      var jamDisplay = "-";

      if (rawJam !== "") {
        jamDisplay = (rawJam instanceof Date) ? Utilities.formatDate(rawJam, Session.getScriptTimeZone(), "HH:mm") : rawJam.toString();
      }

      if (rawTanggal && rawTanggal !== "") {
        if (rawTanggal instanceof Date) {
          var ymd = Utilities.formatDate(rawTanggal, Session.getScriptTimeZone(), "yyyy-MM-dd");
          tglDisplay = toIndoDateString(ymd);
        } else {
          tglDisplay = rawTanggal.toString();
        }
      }

      var waPicRaw = rawData[i][KOLOM.WA_PIC] || "";

      result.data.push({
        tanggal: tglDisplay,
        jam: jamDisplay,
        nim: rawData[i][KOLOM.NIM] ? rawData[i][KOLOM.NIM].toString() : "-",
        nama: rawData[i][KOLOM.NAMA_MHS] ? rawData[i][KOLOM.NAMA_MHS].toString() : "-",
        mitra: isDosen ? (rawData[i][KOLOM.NAMA_SPV] ? rawData[i][KOLOM.NAMA_SPV].toString() : "-") : (rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString() : "-"), 
        perusahaan: rawData[i][KOLOM.PERUSAHAAN] ? rawData[i][KOLOM.PERUSAHAAN].toString() : "-",
        spv: rawData[i][KOLOM.NAMA_SPV] ? rawData[i][KOLOM.NAMA_SPV].toString() : "-",
        dosbing: rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString() : "-",
        status: finalStatus,
        pic: rawData[i][KOLOM.PIC] ? rawData[i][KOLOM.PIC].toString() : "-",
        waPic: waPicRaw !== "" ? formatWA(waPicRaw) : ""
      });
    }
  }

  return result;
}

// ==========================================
// 7. FUNGSI EVALUASI / KUESIONER PASCA RESPONSI
// ==========================================
function getDataForEvaluasi(nim) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var data = sheet.getDataRange().getValues();
 
  for (var i = 1; i < data.length; i++) {
    if (data[i][KOLOM.NIM] == nim) {
      var isWChecked = data[i][KOLOM.CENTANG_W] === true || String(data[i][KOLOM.CENTANG_W]).toUpperCase() === "TRUE"; 
      var isXChecked = data[i][KOLOM.CENTANG_X] === true || String(data[i][KOLOM.CENTANG_X]).toUpperCase() === "TRUE"; 
      
      // VALIDASI: Belum responsi (Kolom W false)
      if (!isWChecked) {
        return { status: "error", message: "Anda belum melaksanakan responsi atau status belum diupdate oleh PIC." };
      }

      // JIKA SUDAH PERNAH MENGISI SEBELUMNYA, KITA KEMBALIKAN DATA LAMA UNTUK DIEDIT
      return {
        status: "success",
        row: i + 1, 
        nim: nim, 
        nama: data[i][KOLOM.NAMA_MHS], 
        perusahaan: data[i][KOLOM.PERUSAHAAN], 
        spv: data[i][KOLOM.NAMA_SPV], 
        alamat: data[i][KOLOM.ALAMAT] || "",
        
        isEditEvaluasi: isXChecked, // True jika sudah pernah ngisi (Kolom X = True)
        
        // Data lama jika ingin diedit
        kehadiranSpvEvaluasi: data[i][KOLOM.KEHADIRAN_SPV] || "",
        namaPenggantiEvaluasi: data[i][KOLOM.NAMA_PENGGANTI] || "",
        jabatanPenggantiEvaluasi: data[i][KOLOM.JABATAN_PENGGANTI] || "",
        spvBerdampak: data[i][KOLOM.SPV_BERDAMPAK] || "",
        jabatanSpvBerdampakManual: data[i][KOLOM.JABATAN_SPV_BERDAMPAK] || ""
      };
    }
  }
  return { status: "error", message: "Data NIM tidak ditemukan." };
}

function simpanEvaluasi(info) {
  if (!info || !info.row) return { status: "error", message: "Data tidak valid." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
    var KOLOM = getKamusKolom(sheet);
    var row = info.row;

    // Tulis ke kolom-kolom baru
    sheet.getRange(row, KOLOM.KEHADIRAN_SPV + 1).setValue(info.kehadiranSpvEvaluasi); 
    
    if(info.kehadiranSpvEvaluasi === "Tidak Hadir") {
      sheet.getRange(row, KOLOM.NAMA_PENGGANTI + 1).setValue(info.namaPenggantiEvaluasi); 
      sheet.getRange(row, KOLOM.JABATAN_PENGGANTI + 1).setValue(info.jabatanPenggantiEvaluasi); 
    } else {
      // Kosongkan kolom jika ternyata hadir (jaga-jaga kalau dia edit dari tidak hadir ke hadir)
      sheet.getRange(row, KOLOM.NAMA_PENGGANTI + 1).setValue(""); 
      sheet.getRange(row, KOLOM.JABATAN_PENGGANTI + 1).setValue(""); 
    }

    sheet.getRange(row, KOLOM.SPV_BERDAMPAK + 1).setValue(info.spvBerdampak); 
    if(info.spvBerdampak === "Lainnya") {
      // Jika pilihannya Lainnya, timpa nama SPV berdamapak dengan input manual
      sheet.getRange(row, KOLOM.SPV_BERDAMPAK + 1).setValue(info.namaSpvBerdampakManual); 
      sheet.getRange(row, KOLOM.JABATAN_SPV_BERDAMPAK + 1).setValue(info.jabatanSpvBerdampakManual); 
    } else {
      // Kosongkan jabatan kalau dia pilih SPV utama miliknya
      sheet.getRange(row, KOLOM.JABATAN_SPV_BERDAMPAK + 1).setValue(""); 
    }

    // AUTO-CENTANG KOLOM X (Tanda Kuesioner Selesai)
    sheet.getRange(row, KOLOM.CENTANG_X + 1).setValue(true); 
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#dcfce3'); // Ubah warna jadi hijau sukses

    return { status: "success" };
  } catch (e) { 
    return { status: "error", message: e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}

// ==========================================
// 8. FUNGSI POLLING NOTIFIKASI PIC
// ==========================================
function cekNotifPIC(picName) {
  if (!picName) return [];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var rawData = sheet.getDataRange().getValues();
  var notifs = [];

  for (var i = 1; i < rawData.length; i++) {
    var picRaw = rawData[i][KOLOM.PIC] ? rawData[i][KOLOM.PIC].toString().trim() : "";
    var statusNotif = rawData[i][KOLOM.STATUS_NOTIF] ? rawData[i][KOLOM.STATUS_NOTIF].toString().trim() : "";

    if (picRaw === picName && statusNotif === "BARU") {
      notifs.push({
        row: i + 1,
        nama: rawData[i][KOLOM.NAMA_MHS],
        nim: rawData[i][KOLOM.NIM]
      });
    }
  }
  return notifs;
}

function tandaiNotifDibaca(row) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var KOLOM = getKamusKolom(sheet);
    sheet.getRange(row, KOLOM.STATUS_NOTIF + 1).setValue(""); 
    return {status: "success"};
  } catch (e) {
    return {status: "error", message: e.message};
  } finally {
    lock.releaseLock();
  }
}
