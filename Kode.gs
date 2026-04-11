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
// Sistem ini akan mencari posisi kolom berdasarkan judulnya di Baris 1.
// Jika teks judul tidak ditemukan, sistem akan memakai indeks default bawaan lama.
function getKamusKolom(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  function cari(namaJudul, indeksDefault) {
    var indeks = headers.indexOf(namaJudul);
    return indeks !== -1 ? indeks : indeksDefault;
  }

  // SILAKAN UBAH TEKS DI DALAM KUTIP (contoh: "Tanggal") AGAR SAMA PERSIS DENGAN HEADER EXCEL-MU
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
    CENTANG_W: cari("Kolom W", 22),       // Ganti kutip ini jika nama headernya berbeda
    CENTANG_X: cari("Kolom X", 23),       // Ganti kutip ini jika nama headernya berbeda
    ALAMAT: cari("Alamat Perusahaan", 24),
    ID_SPV: cari("ID SPV", 25),           // SPV (Z)
    ID_DOSEN: cari("ID Dosen", 26),       // Dosen (AA)
    STATUS_SURAT: cari("Status Surat", 27)
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
// 3. FUNGSI MENGAMBIL DATA (FASE 2 & FASE 1)
// ==========================================
function getPreviewData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet); // <-- Panggil Kamus
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();
  var previewData = [];
 
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][KOLOM.NAMA_MHS] !== "") { 
      var dateValue = rawData[i][KOLOM.TANGGAL]; 
      var jamValue = displayData[i][KOLOM.JAM]; 
      var timestamp = 0;
      var rawTanggal = "";
     
      if (dateValue && dateValue !== "") {
        var dateObj = (dateValue instanceof Date) ? new Date(dateValue.getTime()) : new Date(fromIndoDateString(dateValue.toString()));
        if (!isNaN(dateObj.getTime())) {
          rawTanggal = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
          if (jamValue && jamValue.includes(':')) {
            var timeParts = jamValue.split(':');
            dateObj.setHours(parseInt(timeParts[0], 10), parseInt(timeParts[1], 10), 0, 0);
          }
          timestamp = dateObj.getTime();
        }
      }

      var isWChecked = rawData[i][KOLOM.CENTANG_W] === true || String(rawData[i][KOLOM.CENTANG_W]).toUpperCase() === "TRUE"; 
      var isXChecked = rawData[i][KOLOM.CENTANG_X] === true || String(rawData[i][KOLOM.CENTANG_X]).toUpperCase() === "TRUE"; 
      var finalStatus = "Belum Daftar Ulang"; 

      if (timestamp > 0) {
        if (isWChecked && isXChecked) finalStatus = "Responsi Selesai";
        else if (isWChecked && !isXChecked) finalStatus = "Belum Isi Form Pasca Responsi";
        else finalStatus = "Belum Responsi";
      }

      var picTeknis = displayData[i][KOLOM.PIC] || "-";
      var waPicRaw = rawData[i][KOLOM.WA_PIC] || "";
      var waMhsRaw = rawData[i][KOLOM.WA_MHS] || "";
      var waSpvRaw = rawData[i][KOLOM.WA_SPV] || "";
      var waDosenRaw = rawData[i][KOLOM.WA_DOSEN] || "";
      
      previewData.push({ 
        row: i + 1,
        timestamp: timestamp, 
        tanggal: displayData[i][KOLOM.TANGGAL] || "", 
        tanggalInput: rawTanggal,
        jam: displayData[i][KOLOM.JAM] || "", 
        nim: displayData[i][KOLOM.NIM] || "-",
        nama: displayData[i][KOLOM.NAMA_MHS], 
        perusahaan: displayData[i][KOLOM.PERUSAHAAN] || "-",
        spv: displayData[i][KOLOM.NAMA_SPV] || "-",
        dosbing: displayData[i][KOLOM.DOSBING] || "-",
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
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();
  var phase1Data = [];
 
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][KOLOM.NAMA_MHS] !== "") { 
      var statusKolomAB = rawData[i][KOLOM.STATUS_SURAT];
      phase1Data.push({ 
        nama: displayData[i][KOLOM.NAMA_MHS] || "-", 
        dosbing: displayData[i][KOLOM.DOSBING] ? displayData[i][KOLOM.DOSBING] : "-",
        perusahaan: displayData[i][KOLOM.PERUSAHAAN] || "-", 
        spv: displayData[i][KOLOM.NAMA_SPV] || "-", 
        status: (statusKolomAB && statusKolomAB.toString().trim() !== "") ? "Sudah Konfirmasi" : "Belum Konfirmasi" 
      });
    }
  }
  phase1Data.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return phase1Data;
}

// ==========================================
// 4. FUNGSI CARI & SIMPAN DATA (MAHASISWA & PIC)
// ==========================================
function updateByPIC(info) {
  if (!info || !info.row) return { status: "error", message: "Gagal: Data baris tidak valid." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; 
    var KOLOM = getKamusKolom(sheet);
    var row = info.row;

    // Ingat: fungsi getRange(baris, lajur) dimulai dari indeks 1, jadi kita tambahkan +1 dari Kamus.
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

    sheet.getRange(row, KOLOM.TANGGAL + 1).setValue(toIndoDateString(info.tanggal)); 
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
    
    // Sesuaikan lebar array dengan jumlah kolom maksimum di spreadsheet saat ini
    var maxCol = Math.max(28, sheet.getLastColumn());
    var rowData = new Array(maxCol).fill(""); 
    
    rowData[KOLOM.NO] = newRow - 1; 
    rowData[KOLOM.TANGGAL] = toIndoDateString(info.tanggal); 
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
// 5. FUNGSI PORTAL PERSONALISASI (DOSEN & SPV)
// ==========================================
function getJadwalPersonalisasi(targetId) {
  if (!targetId) return { isFound: false };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var KOLOM = getKamusKolom(sheet);
  
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();

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
    var currentId = (displayData[i][idColumnIndex] || "").toString().trim().toLowerCase();

    if (currentId === targetId.toLowerCase()) {
      
      if (!result.isFound) {
        result.isFound = true;
        if (isDosen) {
          result.name = displayData[i][KOLOM.DOSBING] || "-"; 
        } else {
          result.name = displayData[i][KOLOM.NAMA_SPV] || "-";  
          result.company = displayData[i][KOLOM.PERUSAHAAN] || "-"; 
          result.email = displayData[i][KOLOM.EMAIL_SPV] || "-";  
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

      var waPicRaw = rawData[i][KOLOM.WA_PIC] || "";

      result.data.push({
        tanggal: displayData[i][KOLOM.TANGGAL] || "Belum ditentukan",
        jam: displayData[i][KOLOM.JAM] || "-",
        nim: displayData[i][KOLOM.NIM] || "-",
        nama: displayData[i][KOLOM.NAMA_MHS] || "-",
        mitra: isDosen ? (displayData[i][KOLOM.NAMA_SPV] || "-") : (displayData[i][KOLOM.DOSBING] || "-"), 
        perusahaan: displayData[i][KOLOM.PERUSAHAAN] || "-",
        spv: displayData[i][KOLOM.NAMA_SPV] || "-",
        dosbing: displayData[i][KOLOM.DOSBING] || "-",
        status: finalStatus,
        pic: displayData[i][KOLOM.PIC] || "-",
        waPic: waPicRaw !== "" ? formatWA(waPicRaw) : ""
      });
    }
  }

  return result;
}
