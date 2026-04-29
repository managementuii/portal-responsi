// ==========================================
// VARIABEL GLOBAL UNTUK FORMAT TANGGAL
// ==========================================
var NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
var NAMA_HARI = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
var TIMEZONE_JKT = "Asia/Jakarta"; // HARDCODE ZONA WAKTU (Mencegah Bug Timezone Shift)

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
    JABATAN_PENGGANTI: cari("Jabatan Pengganti", 28),
    SPV_BERDAMPAK: cari("SPV Berdampak", 29),
    JABATAN_SPV_BERDAMPAK: cari("Jabatan SPV Berdampak", 30),
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
// 🚀 FUNGSI PEMBERSIH CACHE
// ==========================================
function clearCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(["CACHE_PREVIEW", "CACHE_DROPDOWN"]);
}


// ==========================================
// 3. FUNGSI MENGAMBIL DATA (OPTIMIZED DENGAN CACHE)
// ==========================================
function getPreviewData() {
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get("CACHE_PREVIEW");
  
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet); 
  var rawData = sheet.getDataRange().getValues(); 
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

      if (rawJam !== "") jamDisplay = (rawJam instanceof Date) ? Utilities.formatDate(rawJam, TIMEZONE_JKT, "HH:mm") : rawJam.toString();

      if (rawTanggal && rawTanggal !== "") {
        if (rawTanggal instanceof Date) {
          dateObj = new Date(rawTanggal.getTime());
          tglInput = Utilities.formatDate(dateObj, TIMEZONE_JKT, "yyyy-MM-dd");
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

  cache.put("CACHE_PREVIEW", JSON.stringify(previewData), 1800);

  return previewData;
}

function getPhase1Data() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var rawData = sheet.getDataRange().getValues();
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

function getDropdownOptions() {
  var cache = CacheService.getScriptCache();
  var cachedData = cache.get("CACHE_DROPDOWN");
  
  if (cachedData) {
    return JSON.parse(cachedData);
  }

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
  
  var result = { companies: companies, spvs: spvs };
  cache.put("CACHE_DROPDOWN", JSON.stringify(result), 3600);

  return result;
}


// ==========================================
// PENGAMAN 1: CLASH CHECKER DOSEN
// ==========================================
function cekBentrokDosen(sheet, KOLOM, nimToIgnore, dosbing, tglIndoStr, jamStr) {
  if (!dosbing || dosbing === "" || dosbing === "-") return false;
  
  var rawData = sheet.getDataRange().getValues();
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][KOLOM.NIM] == nimToIgnore) continue;
    
    var rDosen = rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString().trim() : "";
    if (rDosen !== dosbing.trim()) continue;
    
    var rTgl = rawData[i][KOLOM.TANGGAL];
    var rawTglStr = "";
    if (rTgl !== "") {
      rawTglStr = (rTgl instanceof Date) ? toIndoDateString(Utilities.formatDate(rTgl, TIMEZONE_JKT, "yyyy-MM-dd")) : rTgl.toString().trim();
    }
    
    var rJam = rawData[i][KOLOM.JAM];
    var rawJamStr = "";
    if (rJam !== "") {
      rawJamStr = (rJam instanceof Date) ? Utilities.formatDate(rJam, TIMEZONE_JKT, "HH:mm") : rJam.toString().trim();
    }
    
    if (rawTglStr === tglIndoStr && rawJamStr === jamStr) {
      return rawData[i][KOLOM.NAMA_MHS]; 
    }
  }
  return false; 
}


// ==========================================
// PENGAMAN 2: ALGORITMA AUTO-ASSIGN PIC
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
      if (rTgl && rTgl !== "") rTglStr = (rTgl instanceof Date) ? toIndoDateString(Utilities.formatDate(rTgl, TIMEZONE_JKT, "yyyy-MM-dd")) : rTgl.toString().trim();
      
      var rJamStr = "";
      if (rJam && rJam !== "") rJamStr = (rJam instanceof Date) ? Utilities.formatDate(rJam, TIMEZONE_JKT, "HH:mm") : rJam.toString().trim();

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

  if (picLama && picLama !== "-" && picLama !== "Menunggu Admin") {
      var picLamaTerdaftar = false;
      for (var k=0; k < masterPics.length; k++) { if (masterPics[k].nama === picLama) picLamaTerdaftar = true; }
      
      if (picLamaTerdaftar && busyPics.indexOf(picLama) === -1) {
          return { nama: picLama, wa: getWaPIC(picLama) }; 
      }
  }

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
    
    clearCache();
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function getDataByNIM(nim) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var data = sheet.getDataRange().getValues();
 
  for (var i = 1; i < data.length; i++) {
    if (data[i][KOLOM.NIM] == nim) {
      var tglFormat = (data[i][KOLOM.TANGGAL] instanceof Date) ? Utilities.formatDate(data[i][KOLOM.TANGGAL], TIMEZONE_JKT, "yyyy-MM-dd") : fromIndoDateString(data[i][KOLOM.TANGGAL].toString());
      var jamFormat = (data[i][KOLOM.JAM] instanceof Date) ? Utilities.formatDate(data[i][KOLOM.JAM], TIMEZONE_JKT, "HH:mm") : data[i][KOLOM.JAM];

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
    
    // VALIDASI ROW & NIM (Mencegah manipulasi Inspect Element / Overwriting)
    if (sheet.getRange(row, KOLOM.NIM + 1).getValue() != info.nim) {
      return { status: "error", message: "Akses Ditolak: Data tidak sinkron. Gagal melakukan otentikasi identitas." };
    }

    var isWChecked = sheet.getRange(row, KOLOM.CENTANG_W + 1).getValue() === true || String(sheet.getRange(row, KOLOM.CENTANG_W + 1).getValue()).toUpperCase() === "TRUE"; 
    var isXChecked = sheet.getRange(row, KOLOM.CENTANG_X + 1).getValue() === true || String(sheet.getRange(row, KOLOM.CENTANG_X + 1).getValue()).toUpperCase() === "TRUE"; 
    if (isWChecked && isXChecked) return { status: "error", message: "Responsi telah selesai." };

    var tglIndoStr = toIndoDateString(info.tanggal);

    var mhsBentrok = cekBentrokDosen(sheet, KOLOM, info.nim, info.dosbing, tglIndoStr, info.jam);
    if (mhsBentrok) {
      return { 
        status: "error", 
        message: "Jadwal Bentrok! Dosen pembimbing ini sudah dijadwalkan untuk menguji mahasiswa bernama '" + mhsBentrok + "' pada hari dan jam tersebut. Silakan pilih waktu yang lain." 
      };
    }

    var picSaatIni = sheet.getRange(row, KOLOM.PIC + 1).getValue() ? sheet.getRange(row, KOLOM.PIC + 1).getValue().toString().trim() : "";
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

    sheet.getRange(row, KOLOM.PIC + 1).setValue(assignedPIC.nama);
    sheet.getRange(row, KOLOM.WA_PIC + 1).setValue(formatWA(assignedPIC.wa));
    sheet.getRange(row, KOLOM.STATUS_NOTIF + 1).setValue("BARU");
    
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e0f2fe');
    
    clearCache();
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
    var maxCol = Math.max(31, sheet.getLastColumn()); 
    var rowData = new Array(maxCol).fill(""); 

    var tglIndoStr = toIndoDateString(info.tanggal);

    var mhsBentrok = cekBentrokDosen(sheet, KOLOM, info.nim, info.dosbing, tglIndoStr, info.jam);
    if (mhsBentrok) {
      return { 
        status: "error", 
        message: "Jadwal Bentrok! Dosen pembimbing ini sudah dijadwalkan untuk menguji mahasiswa bernama '" + mhsBentrok + "' pada hari dan jam tersebut. Silakan pilih waktu yang lain." 
      };
    }

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

    rowData[KOLOM.PIC] = assignedPIC.nama;
    rowData[KOLOM.WA_PIC] = formatWA(assignedPIC.wa);
    rowData[KOLOM.STATUS_NOTIF] = "BARU";
    
    sheet.getRange(newRow, 1, 1, maxCol).setValues([rowData]); 
    sheet.getRange(newRow, 1, 1, maxCol).setBackground('#ecfdf5');
    
    clearCache();
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
    
    // VALIDASI ROW & NIM
    if (sheet.getRange(row, KOLOM.NIM + 1).getValue() != info.nim) {
        return { status: "error", message: "Akses Ditolak: Data tidak sinkron." };
    }
    
    sheet.getRange(row, KOLOM.WA_MHS + 1).setValue(formatWA(info.waMhs)); 
    sheet.getRange(row, KOLOM.PERUSAHAAN + 1).setValue(info.perusahaan);      
    sheet.getRange(row, KOLOM.NAMA_SPV + 1).setValue(info.spv); 
    sheet.getRange(row, KOLOM.WA_SPV + 1).setValue(formatWA(info.waSpv));              
    sheet.getRange(row, KOLOM.EMAIL_SPV + 1).setValue(info.emailSpv); 
    sheet.getRange(row, KOLOM.EMAIL_SITA + 1).setValue(info.emailSita);      
    sheet.getRange(row, KOLOM.ALAMAT + 1).setValue(info.alamat); 
    sheet.getRange(row, KOLOM.STATUS_SURAT + 1).setValue("Sudah Konfirmasi");  
    
    clearCache();
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
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
        jamDisplay = (rawJam instanceof Date) ? Utilities.formatDate(rawJam, TIMEZONE_JKT, "HH:mm") : rawJam.toString();
      }

      if (rawTanggal && rawTanggal !== "") {
        if (rawTanggal instanceof Date) {
          var ymd = Utilities.formatDate(rawTanggal, TIMEZONE_JKT, "yyyy-MM-dd");
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
      
      if (!isWChecked) {
        return { status: "error", message: "Anda belum melaksanakan responsi atau status belum diupdate oleh PIC." };
      }

      return {
        status: "success",
        row: i + 1, 
        nim: nim, 
        nama: data[i][KOLOM.NAMA_MHS], 
        perusahaan: data[i][KOLOM.PERUSAHAAN], 
        spv: data[i][KOLOM.NAMA_SPV], 
        alamat: data[i][KOLOM.ALAMAT] || "",
        
        isEditEvaluasi: isXChecked, 
        
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
  if (!info || !info.row || !info.nim) return { status: "error", message: "Data tidak valid atau kosong." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
    var KOLOM = getKamusKolom(sheet);
    var row = info.row;

    // VALIDASI ROW & NIM (Mencegah manipulasi Inspect Element / Overwriting)
    if (sheet.getRange(row, KOLOM.NIM + 1).getValue() != info.nim) {
      return { status: "error", message: "Akses Ditolak: Data tidak sinkron. Gagal melakukan otentikasi identitas." };
    }

    // SIMPAN DATA ALAMAT BARU
    if (info.alamatEvaluasi !== undefined) {
      sheet.getRange(row, KOLOM.ALAMAT + 1).setValue(info.alamatEvaluasi);
    }

    sheet.getRange(row, KOLOM.KEHADIRAN_SPV + 1).setValue(info.kehadiranSpvEvaluasi); 
    
    if(info.kehadiranSpvEvaluasi === "Tidak Hadir") {
      sheet.getRange(row, KOLOM.NAMA_PENGGANTI + 1).setValue(info.namaPenggantiEvaluasi); 
      sheet.getRange(row, KOLOM.JABATAN_PENGGANTI + 1).setValue(info.jabatanPenggantiEvaluasi); 
    } else {
      sheet.getRange(row, KOLOM.NAMA_PENGGANTI + 1).setValue(""); 
      sheet.getRange(row, KOLOM.JABATAN_PENGGANTI + 1).setValue(""); 
    }

    sheet.getRange(row, KOLOM.SPV_BERDAMPAK + 1).setValue(info.spvBerdampak); 
    if(info.spvBerdampak === "Lainnya") {
      sheet.getRange(row, KOLOM.SPV_BERDAMPAK + 1).setValue(info.namaSpvBerdampakManual); 
      sheet.getRange(row, KOLOM.JABATAN_SPV_BERDAMPAK + 1).setValue(info.jabatanSpvBerdampakManual); 
    } else {
      sheet.getRange(row, KOLOM.JABATAN_SPV_BERDAMPAK + 1).setValue(""); 
    }

    sheet.getRange(row, KOLOM.CENTANG_X + 1).setValue(true); 
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#dcfce3'); 

    clearCache();
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

// ==========================================
// 9. FUNGSI AUTO-GENERATE SURAT (ADMIN PANEL)
// ==========================================

// A. Fungsi untuk menarik daftar mahasiswa yang butuh surat
function getAntreanSurat() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  var rawData = sheet.getDataRange().getValues();
  var antrean = [];
 
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][KOLOM.NAMA_MHS] !== "") { 
      var statusSrt = rawData[i][KOLOM.STATUS_SURAT] ? rawData[i][KOLOM.STATUS_SURAT].toString().trim() : "";
      
      antrean.push({ 
        nim: rawData[i][KOLOM.NIM].toString(),
        nama: rawData[i][KOLOM.NAMA_MHS].toString(),
        perusahaan: rawData[i][KOLOM.PERUSAHAAN] ? rawData[i][KOLOM.PERUSAHAAN].toString() : "-",
        spv: rawData[i][KOLOM.NAMA_SPV] ? rawData[i][KOLOM.NAMA_SPV].toString() : "-",
        dosbing: rawData[i][KOLOM.DOSBING] ? rawData[i][KOLOM.DOSBING].toString() : "-", // PASTIKAN BARIS INI ADA
        status: (statusSrt === "Sudah Konfirmasi") ? "Sudah Konfirmasi" : "Belum Konfirmasi"
      });
    }
  }
  return antrean;
}


// B. Fungsi Mesin Cetak PDF
function generateSuratPengantar(nim) {
  // ID yang sudah Anda siapkan
  var TEMPLATE_ID = "1gV8AJI-VrQ9rSwIhm_BuTQ_M9DHEvL0GafWMyZvi5Cg"; 
  var FOLDER_ID = "1q9q1rLXBAeUHkxL_WcBRIYc73UeWIoP1";
  var PORTAL_URL = "https://management.uii.ac.id/portal-responsi/?"; 

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var KOLOM = getKamusKolom(sheet);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][KOLOM.NIM] == nim) {
        
        var namaSpv = data[i][KOLOM.NAMA_SPV] || "-";
        var perusahaan = data[i][KOLOM.PERUSAHAAN] || "-";
        // Ambil ID SPV (yang ?spv1122 dll)
        var idSpv = data[i][KOLOM.ID_SPV] ? data[i][KOLOM.ID_SPV].toString().trim() : ""; 
        var mhsName = data[i][KOLOM.NAMA_MHS] || nim;

        // Cegat jika ID SPV belum ada di Spreadsheet
        if(!idSpv || idSpv === "") {
            return {status: "error", message: "ID SPV belum tersedia di Spreadsheet (Kolom ID SPV). Pastikan ID SPV sudah di-generate agar link Dashboard valid."};
        }

        var linkDashboard = PORTAL_URL + idSpv;
        var namaFilePdf = "Surat_Undangan_SPV_" + perusahaan.replace(/[^a-zA-Z0-9]/g, "") + "_" + mhsName.replace(/[^a-zA-Z0-9]/g, "");

        // 1. Duplikasi Template Docs
        var fileTemplate = DriveApp.getFileById(TEMPLATE_ID);
        var folderTujuan = DriveApp.getFolderById(FOLDER_ID);
        var docCopy = fileTemplate.makeCopy(namaFilePdf, folderTujuan);
        
        // 2. Buka file duplikat dan lakukan Replace Text (Mail Merge)
        var doc = DocumentApp.openById(docCopy.getId());
        var body = doc.getBody();
        
        body.replaceText("<<Nama SPV>>", namaSpv);
        body.replaceText("<<Perusahaan>>", perusahaan);
        body.replaceText("<<Link>>", linkDashboard);
        
        doc.saveAndClose();

        // 3. Konversi menjadi PDF
        var blobPdf = docCopy.getAs('application/pdf');
        var filePdf = folderTujuan.createFile(blobPdf);
        filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        // 4. Hapus file Docs Temporary agar drive bersih
        docCopy.setTrashed(true);
        
        // (Opsional) Jika ingin otomatis menandai di sheet bahwa surat sudah dibuat, Anda bisa uncomment kode di bawah:
        // sheet.getRange(i + 1, KOLOM.STATUS_SURAT + 1).setValue("Surat Telah Di-Generate");

        return {status: "success", url: filePdf.getUrl(), message: "Surat berhasil dibuat!"};
      }
    }
    return {status: "error", message: "NIM Mahasiswa tidak ditemukan."};
  } catch (error) {
    return {status: "error", message: error.message};
  }
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

// ==========================================
// 10. FUNGSI AUTO-GENERATE SURAT BULK (GROUPING SPV)
// ==========================================
function generateSuratGrupSPV(groupsToProcess) {
  var TEMPLATE_ID = "1gV8AJI-VrQ9rSwIhm_BuTQ_M9DHEvL0GafWMyZvi5Cg"; 
  var FOLDER_ID = "1q9q1rLXBAeUHkxL_WcBRIYc73UeWIoP1";
  var PORTAL_URL = "https://management.uii.ac.id/portal-responsi/?"; 

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var KOLOM = getKamusKolom(sheet);
    var data = sheet.getDataRange().getValues();

    var hasil = { sukses: 0, gagal: 0, pesan: [] };

    // Looping sebanyak jumlah SPV yang dicentang
    for (var g = 0; g < groupsToProcess.length; g++) {
      var group = groupsToProcess[g];
      var namaSpv = group.spv || "-";
      var perusahaan = group.perusahaan || "-";

      // Karena 1 grup memiliki SPV yang sama, kita ambil ID SPV dari mhs pertama saja
      var firstNim = group.mahasiswa[0].nim;
      var idSpv = "";

      // Cari ID SPV di Spreadsheet
      for (var i = 1; i < data.length; i++) {
        if (data[i][KOLOM.NIM] == firstNim) {
           idSpv = data[i][KOLOM.ID_SPV] ? data[i][KOLOM.ID_SPV].toString().trim() : "";
           break;
        }
      }

      if(!idSpv || idSpv === "") {
         hasil.gagal++;
         hasil.pesan.push("Gagal: ID SPV untuk " + namaSpv + " belum di-generate di Spreadsheet.");
         continue; // Lanjut ke SPV berikutnya (skip yang ini)
      }

      var linkDashboard = PORTAL_URL + idSpv;
      var namaFilePdf = "Surat_Undangan_SPV_" + perusahaan.replace(/[^a-zA-Z0-9]/g, "") + "_" + namaSpv.replace(/[^a-zA-Z0-9]/g, "");

      // Rangkai Teks Daftar Mahasiswa (misal: "1. Budi (20311...)\n2. Andi (20311...)")
      var daftarMahasiswaTeks = "";
      for(var m = 0; m < group.mahasiswa.length; m++) {
         var mhs = group.mahasiswa[m];
         daftarMahasiswaTeks += (m + 1) + ". " + mhs.nama + " (" + mhs.nim + ")\n";
      }

      // 1. Duplikasi Template
      var fileTemplate = DriveApp.getFileById(TEMPLATE_ID);
      var folderTujuan = DriveApp.getFolderById(FOLDER_ID);
      var docCopy = fileTemplate.makeCopy(namaFilePdf, folderTujuan);
      
      // 2. Buka duplikat dan lakukan Replace Text (Mail Merge)
      var doc = DocumentApp.openById(docCopy.getId());
      var body = doc.getBody();
      
      body.replaceText("<<Nama SPV>>", namaSpv);
      body.replaceText("<<Perusahaan>>", perusahaan);
      body.replaceText("<<Link>>", linkDashboard);
      
      // MENGGANTI PLACEHOLDER DAFTAR MAHASISWA
      body.replaceText("<<Daftar Mahasiswa>>", daftarMahasiswaTeks.trim());
      
      doc.saveAndClose();

      // 3. Konversi ke PDF & Buka Akses
      var blobPdf = docCopy.getAs('application/pdf');
      var filePdf = folderTujuan.createFile(blobPdf);
      filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // 4. Bersihkan file Docs
      docCopy.setTrashed(true);

      // (Opsi) Tandai baris mahasiswa ini bahwa suratnya sudah selesai
      // Bisa dilooping lagi untuk menulis ke kolom STATUS_SURAT

      hasil.sukses++;
    }

    return { status: "success", summary: hasil };

  } catch (error) {
    return { status: "error", message: error.message };
  }
}
