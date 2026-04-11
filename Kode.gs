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
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();
  var previewData = [];
 
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][4] !== "") { 
      var dateValue = rawData[i][1]; 
      var jamValue = displayData[i][2]; 
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

      var isWChecked = rawData[i][22] === true || String(rawData[i][22]).toUpperCase() === "TRUE"; 
      var isXChecked = rawData[i][23] === true || String(rawData[i][23]).toUpperCase() === "TRUE"; 
      var finalStatus = "Belum Daftar Ulang"; 

      if (timestamp > 0) {
        if (isWChecked && isXChecked) finalStatus = "Responsi Selesai";
        else if (isWChecked && !isXChecked) finalStatus = "Belum Isi Form Pasca Responsi";
        else finalStatus = "Belum Responsi";
      }

      var picTeknis = displayData[i][13] || "-";
      var waPicRaw = rawData[i][14] || "";
      var waMhsRaw = rawData[i][5] || "";
      var waSpvRaw = rawData[i][8] || "";
      var waDosenRaw = rawData[i][11] || ""; // Kolom L
      
      previewData.push({ 
        row: i + 1,
        timestamp: timestamp, 
        tanggal: displayData[i][1] || "", 
        tanggalInput: rawTanggal,
        jam: displayData[i][2] || "", 
        nim: displayData[i][3] || "-",
        nama: displayData[i][4], 
        perusahaan: displayData[i][6] || "-",
        spv: displayData[i][7] || "-",
        dosbing: displayData[i][10] || "-",
        waMhs: waMhsRaw !== "" ? formatWA(waMhsRaw) : "",
        waSpv: waSpvRaw !== "" ? formatWA(waSpvRaw) : "",
        waDosen: waDosenRaw !== "" ? formatWA(waDosenRaw) : "",
        pic: picTeknis, 
        waPic: waPicRaw !== "" ? formatWA(waPicRaw) : "", 
        status: finalStatus,
        isWChecked: isWChecked // Menambahkan info centang W
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
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();
  var phase1Data = [];
 
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][4] !== "") { 
      var statusKolomAB = rawData[i][27];
      phase1Data.push({ 
        nama: displayData[i][4] || "-", dosbing: displayData[i][10] ? displayData[i][10] : "-",
        perusahaan: displayData[i][6] || "-", spv: displayData[i][7] || "-", 
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
    var row = info.row;

    // Eksekusi Pembaruan Data
    if (info.tanggal) sheet.getRange(row, 2).setValue(toIndoDateString(info.tanggal)); 
    if (info.jam) sheet.getRange(row, 3).setValue(info.jam);        
    if (info.spv) sheet.getRange(row, 8).setValue(info.spv);
    if (info.waMhs) sheet.getRange(row, 6).setValue(formatWA(info.waMhs)); 
    if (info.waSpv) sheet.getRange(row, 9).setValue(formatWA(info.waSpv));      
    if (info.waDosen !== undefined) sheet.getRange(row, 12).setValue(formatWA(info.waDosen));

    // UPDATE STATUS CENTANG KOLOM W (23) DARI PIC
    if (info.isWChecked !== undefined) {
        sheet.getRange(row, 23).setValue(info.isWChecked ? true : false);
    }

    // Beri highlight warna kuning muda (tanda diedit oleh PIC)
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#fef9c3');
    
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function getDataByNIM(nim) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
 
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] == nim) {
      var tglFormat = (data[i][1] instanceof Date) ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd") : fromIndoDateString(data[i][1].toString());
      var jamFormat = (data[i][2] instanceof Date) ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), "HH:mm") : data[i][2];

      var isWChecked = data[i][22] === true || String(data[i][22]).toUpperCase() === "TRUE"; 
      var isXChecked = data[i][23] === true || String(data[i][23]).toUpperCase() === "TRUE"; 
      
      var rawWaMhs = data[i][5] ? data[i][5].toString() : ""; if(rawWaMhs.startsWith("62")) rawWaMhs = "0" + rawWaMhs.substring(2); else if(rawWaMhs.startsWith("8")) rawWaMhs = "0" + rawWaMhs;
      var rawWaSpv = data[i][8] ? data[i][8].toString() : ""; if(rawWaSpv.startsWith("62")) rawWaSpv = "0" + rawWaSpv.substring(2); else if(rawWaSpv.startsWith("8")) rawWaSpv = "0" + rawWaSpv;
      var rawWaPengganti = data[i][20] ? data[i][20].toString() : ""; if(rawWaPengganti.startsWith("62")) rawWaPengganti = "0" + rawWaPengganti.substring(2); else if(rawWaPengganti.startsWith("8")) rawWaPengganti = "0" + rawWaPengganti;
      var rawWaPic = data[i][14] ? data[i][14].toString() : ""; if(rawWaPic.startsWith("62")) rawWaPic = "0" + rawWaPic.substring(2); else if(rawWaPic.startsWith("8")) rawWaPic = "0" + rawWaPic;

      return {
        isNew: false, isFinished: (isWChecked && isXChecked), row: i + 1, nim: nim, nama: data[i][4], waMhs: rawWaMhs,    
        tanggal: tglFormat, jam: jamFormat, perusahaan: data[i][6], spv: data[i][7], waSpv: rawWaSpv, dosbing: data[i][10],        
        emailSpv: data[i][16] || "", emailSita: data[i][17] || "", kehadiranSpv: data[i][18] || "", namaPengganti: data[i][19] || "",
        waPengganti: rawWaPengganti, emailPengganti: data[i][21] || "", alamatPerusahaan: data[i][24] || "", pic: data[i][13] || "-", waPic: rawWaPic
      };
    }
  }
  return { isNew: true, isFinished: false, nim: nim };
}

function simpanUpdate(info) {
  if (!info || !info.nim) return { status: "error", message: "Gagal: Data kosong." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; var row = info.row;
    var isWChecked = sheet.getRange(row, 23).getValue() === true || String(sheet.getRange(row, 23).getValue()).toUpperCase() === "TRUE"; 
    var isXChecked = sheet.getRange(row, 24).getValue() === true || String(sheet.getRange(row, 24).getValue()).toUpperCase() === "TRUE"; 
    if (isWChecked && isXChecked) return { status: "error", message: "Responsi telah selesai." };

    sheet.getRange(row, 2).setValue(toIndoDateString(info.tanggal)); sheet.getRange(row, 3).setValue(info.jam);        
    sheet.getRange(row, 6).setValue(formatWA(info.waMhs)); sheet.getRange(row, 7).setValue(info.perusahaan); 
    sheet.getRange(row, 8).setValue(info.spv); sheet.getRange(row, 9).setValue(formatWA(info.waSpv));      
    sheet.getRange(row, 16).setValue("Belum Responsi"); sheet.getRange(row, 17).setValue(info.emailSpv);  
    sheet.getRange(row, 18).setValue(info.emailSita); sheet.getRange(row, 19).setValue(info.kehadiranSpv);
    sheet.getRange(row, 20).setValue(info.namaPengganti); sheet.getRange(row, 21).setValue(formatWA(info.waPengganti));
    sheet.getRange(row, 22).setValue(info.emailPengganti);
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#e0f2fe');
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function simpanBaru(info) {
  if (!info || !info.nim) return { status: "error", message: "Gagal: Data kosong." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; var newRow = sheet.getLastRow() + 1;
    var rowData = new Array(28).fill(""); 
    rowData[0] = newRow - 1; rowData[1] = toIndoDateString(info.tanggal); rowData[2] = info.jam; rowData[3] = info.nim;                       
    rowData[4] = info.nama; rowData[5] = formatWA(info.waMhs); rowData[6] = info.perusahaan; rowData[7] = info.spv;                       
    rowData[8] = formatWA(info.waSpv); rowData[10] = info.dosbing; rowData[15] = "Belum Responsi"; rowData[16] = info.emailSpv;                 
    rowData[17] = info.emailSita; rowData[18] = info.kehadiranSpv; rowData[19] = info.namaPengganti; rowData[20] = formatWA(info.waPengganti);
    rowData[21] = info.emailPengganti;
    sheet.getRange(newRow, 1, 1, 28).setValues([rowData]); sheet.getRange(newRow, 1, 1, 28).setBackground('#ecfdf5');
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function simpanSurat(info) {
  if (!info || !info.nim) return { status: "error", message: "Gagal: NIM kosong." };
  var lock = LockService.getScriptLock(); lock.waitLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; var row = info.row;
    if (sheet.getRange(row, 4).getValue() != info.nim) return { status: "error", message: "Akses Ditolak: Data tidak sinkron." };
    sheet.getRange(row, 6).setValue(formatWA(info.waMhs)); sheet.getRange(row, 7).setValue(info.perusahaan);      
    sheet.getRange(row, 8).setValue(info.spv); sheet.getRange(row, 9).setValue(formatWA(info.waSpv));              
    sheet.getRange(row, 17).setValue(info.emailSpv); sheet.getRange(row, 18).setValue(info.emailSita);      
    sheet.getRange(row, 25).setValue(info.alamat); sheet.getRange(row, 28).setValue("Sudah Konfirmasi");  
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; } finally { lock.releaseLock(); }
}

function getDropdownOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; var rawData = sheet.getDataRange().getValues();
  var companies = []; var spvs = [];
  for (var i = 1; i < rawData.length; i++) {
    if (rawData[i][6] && rawData[i][6] !== "" && companies.indexOf(rawData[i][6]) === -1) companies.push(rawData[i][6]);
    if (rawData[i][7] && rawData[i][7] !== "" && spvs.indexOf(rawData[i][7]) === -1) spvs.push(rawData[i][7]);
  }
  companies.sort(); spvs.sort();
  return { companies: companies, spvs: spvs };
}


// ==========================================
// 5. FUNGSI PORTAL PERSONALISASI (DOSEN & SPV)
// ==========================================
function getJadwalPersonalisasi(targetId) {
  if (!targetId) return { isFound: false };

  // ⚠️ PENTING: KONFIGURASI KOLOM ID ⚠️
  // Sesuaikan angka ini dengan posisi kolom di Google Spreadsheet Anda yang menyimpan ID "dos..." dan "spv...".
  // Ingat: Indeks array di Google Apps Script dimulai dari 0.
  // (Kolom A=0, B=1, ... Y=24, Z=25, AA=26, AB=27, dst.)
  var INDEX_KOLOM_ID_DOSEN = 26; // Dosen di Kolom AA (diubah dari 25 ke 26)
  var INDEX_KOLOM_ID_SPV = 25;   // SPV di Kolom Z (diubah dari 26 ke 25)

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();

  var isDosen = targetId.toLowerCase().startsWith('dos');
  var idColumnIndex = isDosen ? INDEX_KOLOM_ID_DOSEN : INDEX_KOLOM_ID_SPV;

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

    // Jika ID di baris ini cocok dengan targetId dari URL (?dos2023)
    if (currentId === targetId.toLowerCase()) {
      
      // Ambil Profil (Hanya dipanggil sekali saat pencocokan pertama)
      if (!result.isFound) {
        result.isFound = true;
        if (isDosen) {
          result.name = displayData[i][10] || "-"; // Nama Dosen Pembimbing (Kolom K)
        } else {
          result.name = displayData[i][7] || "-";  // Nama Supervisor (Kolom H)
          result.company = displayData[i][6] || "-"; // Perusahaan (Kolom G)
          result.email = displayData[i][16] || "-";  // Email SPV (Kolom Q)
        }
      }

      // Format Data Baris Jadwal
      var isWChecked = rawData[i][22] === true || String(rawData[i][22]).toUpperCase() === "TRUE";
      var isXChecked = rawData[i][23] === true || String(rawData[i][23]).toUpperCase() === "TRUE";
      var finalStatus = "Belum Responsi";

      // Logika status yang sama dengan Dashboard PIC/Mahasiswa
      if (isWChecked && isXChecked) {
        finalStatus = "Responsi Selesai";
      } else if (isWChecked && !isXChecked) {
        finalStatus = "Belum Isi Form"; // Teks persis yang dicek oleh Portal.html
      }

      var waPicRaw = rawData[i][14] || "";

      result.data.push({
        tanggal: displayData[i][1] || "Belum ditentukan",
        jam: displayData[i][2] || "-",
        nim: displayData[i][3] || "-",
        nama: displayData[i][4] || "-",
        mitra: isDosen ? (displayData[i][7] || "-") : (displayData[i][10] || "-"), // Info silang
        perusahaan: displayData[i][6] || "-",
        spv: displayData[i][7] || "-",
        dosbing: displayData[i][10] || "-",
        status: finalStatus,
        pic: displayData[i][13] || "-",
        waPic: waPicRaw !== "" ? formatWA(waPicRaw) : ""
      });
    }
  }

  return result;
}
