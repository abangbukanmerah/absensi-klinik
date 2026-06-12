/**
 * SISTEM ABSENSI MANDIRI - KLINIK VIDYA MEDIKA
 * Backend Script (Code.gs) - Versi 5.3 (Clean Compilation & Auto-Sync)
 * * Logika anti-double absen dalam 1 hari yang sama dengan pencatatan unit kerja,
 * shift yang fleksibel, dan sistem pendeteksi spreadsheet otomatis yang bebas bug.
 */

// ====================================================================================
// SPREADSHEET ID KLINIK VIDYA MEDIKA
// ------------------------------------------------------------------------------------
// TIPS: Kosongkan saja variabel di bawah ini ( const SPREADSHEET_ID = ""; ) jika Anda 
// membuat script ini langsung melalui menu "Ekstensi > Apps Script" di Google Sheet.
// Sistem akan mendeteksi spreadsheet Anda secara otomatis!
// ====================================================================================
const SPREADSHEET_ID = "1hhLgnvsVud27C4ugsiK_CYTvUIzBsQc5btE81hL23JY"; 

/**
 * Helper untuk mengambil Spreadsheet secara aman (Bound & Standalone Script)
 * Dibuat fail-safe agar tidak crash jika ID salah atau kosong.
 */
function getSpreadsheet() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID.trim().length > 10) {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e) {
      // Jika gagal (masalah hak akses/ID salah), lanjut ke fallback active spreadsheet
    }
  }
  
  // Fallback: Gunakan spreadsheet aktif tempat script ini terikat (Bound Script)
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  
  throw new Error("Spreadsheet tidak ditemukan! Pastikan SPREADSHEET_ID di baris paling atas Code.gs sudah diisi dengan ID Spreadsheet Anda yang benar, atau pastikan script ini dibuat langsung dari menu 'Ekstensi > Apps Script' di dalam Google Sheet Anda.");
}

// Fungsi untuk melayani halaman HTML utama
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Absensi Mandiri - Klinik Vidya Medika')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * SETUP DATABASE OTOMATIS
 * Mengupgrade database yang ada agar memiliki struktur kolom Unit dan Shift baru
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  
  // 1. SETUP SHEET PENGATURAN
  let sheetSetting = ss.getSheetByName('Pengaturan');
  if (!sheetSetting) {
    sheetSetting = ss.insertSheet('Pengaturan');
    sheetSetting.appendRow(['Nama Parameter', 'Nilai', 'Keterangan']);
    sheetSetting.appendRow(['ADMIN_PIN', '1234', 'PIN untuk mengakses dashboard admin (default: 1234)']);
    sheetSetting.appendRow(['NAMA_KLINIK', 'Klinik Vidya Medika', 'Nama Instansi']);
    sheetSetting.getRange('A1:C1').setFontWeight('bold').setBackground('#009639').setFontColor('#FFFFFF');
    sheetSetting.autoResizeColumns(1, 3);
  }
  
  // 2. SETUP SHEET KARYAWAN (Dengan kolom Unit)
  let sheetKaryawan = ss.getSheetByName('Karyawan');
  if (!sheetKaryawan) {
    sheetKaryawan = ss.insertSheet('Karyawan');
    sheetKaryawan.appendRow(['ID Pegawai', 'Nama Lengkap', 'Jabatan', 'Status', 'Unit']);
    sheetKaryawan.appendRow(['VM001', 'dr. Ahmad Subarjo', 'Dokter Umum', 'Aktif', 'Pelayanan Medis']);
    sheetKaryawan.appendRow(['VM002', 'Siti Rahma, S.Kep', 'Perawat Kepala', 'Aktif', 'Keperawatan']);
    sheetKaryawan.appendRow(['VM003', 'Budi Santoso, Amd.Far', 'Apoteker', 'Aktif', 'Farmasi']);
    sheetKaryawan.appendRow(['VM004', 'Lani Fitriani', 'Administrasi', 'Aktif', 'Manajemen']);
    
    sheetKaryawan.getRange('A1:E1').setFontWeight('bold').setBackground('#009639').setFontColor('#FFFFFF');
    sheetKaryawan.autoResizeColumns(1, 5);
  } else {
    // Jika sheet karyawan sudah ada, pastikan kolom Unit terdaftar
    const headerRow = sheetKaryawan.getRange(1, 1, 1, sheetKaryawan.getLastColumn()).getValues()[0];
    const headersLower = headerRow.map(function(h) { return h.toString().toLowerCase().trim(); });
    if (headersLower.indexOf('unit') === -1) {
      sheetKaryawan.getRange(1, sheetKaryawan.getLastColumn() + 1).setValue('Unit').setFontWeight('bold');
    }
  }
  
  // 3. SETUP SHEET KEHADIRAN (Dengan format baru)
  let sheetKehadiran = ss.getSheetByName('Kehadiran');
  const newHeaders = ['Timestamp', 'Tanggal', 'Jam', 'ID Pegawai', 'Nama Pegawai', 'Jabatan', 'Unit', 'Tipe Absen', 'Shift', 'Latitude', 'Longitude', 'Link Lokasi'];
  
  if (!sheetKehadiran) {
    sheetKehadiran = ss.insertSheet('Kehadiran');
    sheetKehadiran.appendRow(newHeaders);
    sheetKehadiran.getRange('A1:L1').setFontWeight('bold').setBackground('#009639').setFontColor('#FFFFFF');
    sheetKehadiran.autoResizeColumns(1, 12);
  } else {
    // Upgrade baris header sheet kehadiran yang sudah ada tanpa menghapus data absensi lama
    sheetKehadiran.getRange(1, 1, 1, 12).setValues([newHeaders]).setFontWeight('bold').setBackground('#009639').setFontColor('#FFFFFF');
    sheetKehadiran.autoResizeColumns(1, 12);
  }
  
  SpreadsheetApp.flush();
  return "Database Berhasil Diinisialisasi!";
}

/**
 * Mengambil daftar seluruh karyawan aktif dengan kolom Unit (Fail-Safe dengan Auto-Inisialisasi)
 */
function getActiveEmployees() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Karyawan');
    
    // Proteksi Auto-Inisialisasi jika sheet hilang/belum siap
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Karyawan');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    const idIdx = headers.indexOf('id pegawai');
    const namaIdx = headers.indexOf('nama lengkap');
    const jabatanIdx = headers.indexOf('jabatan');
    const statusIdx = headers.indexOf('status');
    const unitIdx = headers.indexOf('unit');
    
    const employees = [];
    
    for (let i = 1; i < data.length; i++) {
      const id = data[i][idIdx !== -1 ? idIdx : 0];
      const nama = data[i][namaIdx !== -1 ? namaIdx : 1];
      const jabatan = data[i][jabatanIdx !== -1 ? jabatanIdx : 2];
      const status = data[i][statusIdx !== -1 ? statusIdx : 3];
      const unit = unitIdx !== -1 ? data[i][unitIdx] : "-";
      
      if (status && status.toString().trim().toLowerCase() === 'aktif' && id) {
        employees.push({ 
          id: id.toString(), 
          nama: nama.toString(), 
          jabatan: jabatan.toString(),
          unit: unit ? unit.toString() : "-"
        });
      }
    }
    return employees;
  } catch (e) {
    throw new Error("Gagal mengambil data karyawan: " + e.message);
  }
}

/**
 * Mencatat kehadiran ke dalam Sheet Kehadiran secara real-time dengan proteksi ganda, Unit, dan Shift
 */
function recordAttendance(employeeId, type, latitude, longitude, selectedShift) {
  try {
    const ss = getSpreadsheet();
    let sheetKaryawan = ss.getSheetByName('Karyawan');
    let sheetKehadiran = ss.getSheetByName('Kehadiran');
    
    // Proteksi Auto-Inisialisasi jika sheet hilang/belum siap
    if (!sheetKaryawan || !sheetKehadiran) {
      setupDatabase();
      sheetKaryawan = ss.getSheetByName('Karyawan');
      sheetKehadiran = ss.getSheetByName('Kehadiran');
    }
    
    // Cari data karyawan berdasarkan ID menggunakan index tajuk dinamis
    const karyawanData = sheetKaryawan.getDataRange().getValues();
    const karyawanHeaders = karyawanData[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    const idIdx = karyawanHeaders.indexOf('id pegawai');
    const namaIdx = karyawanHeaders.indexOf('nama lengkap');
    const jabatanIdx = karyawanHeaders.indexOf('jabatan');
    const unitIdx = karyawanHeaders.indexOf('unit');
    
    let karyawan = null;
    for (let i = 1; i < karyawanData.length; i++) {
      const rowId = karyawanData[i][idIdx !== -1 ? idIdx : 0].toString().trim();
      if (rowId === employeeId.toString().trim()) {
        karyawan = {
          id: karyawanData[i][idIdx !== -1 ? idIdx : 0].toString(),
          nama: karyawanData[i][namaIdx !== -1 ? namaIdx : 1].toString(),
          jabatan: karyawanData[i][jabatanIdx !== -1 ? jabatanIdx : 2].toString(),
          unit: (unitIdx !== -1 && karyawanData[i][unitIdx]) ? karyawanData[i][unitIdx].toString() : "-"
        };
        break;
      }
    }
    
    if (!karyawan) {
      return { success: false, message: "ID Pegawai tidak ditemukan." };
    }
    
    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone();
    
    const formattedDate = Utilities.formatDate(now, tz, "dd-MM-yyyy");
    const formattedTime = Utilities.formatDate(now, tz, "HH:mm:ss");
    
    // Normalisasi input tipe absen untuk komparasi duplikasi
    let inputDisplayTipe = "Datang";
    if (type.toLowerCase().includes("manual")) {
      inputDisplayTipe = type.split(" (Manual: ")[0];
    } else if (type.toLowerCase() === 'pulang') {
      inputDisplayTipe = "Pulang";
    } else if (type.toLowerCase() === 'datang') {
      inputDisplayTipe = "Datang";
    } else {
      inputDisplayTipe = type;
    }

    // Proteksi: Blokir absensi ganda tipe dan shift yang sama pada hari yang sama
    const dataKehadiran = sheetKehadiran.getDataRange().getValues();
    const kehadiranHeaders = dataKehadiran[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    const khTanggalIdx = kehadiranHeaders.indexOf('tanggal');
    const khIdIdx = kehadiranHeaders.indexOf('id pegawai');
    const khTipeIdx = kehadiranHeaders.indexOf('tipe absen');
    
    for (let i = 1; i < dataKehadiran.length; i++) {
      const rowDate = cleanAndFormatDate(dataKehadiran[i][khTanggalIdx !== -1 ? khTanggalIdx : 1], tz);
      const rowEmpId = dataKehadiran[i][khIdIdx !== -1 ? khIdIdx : 3] ? dataKehadiran[i][khIdIdx !== -1 ? khIdIdx : 3].toString().trim() : "";
      const rowTipeFull = dataKehadiran[i][khTipeIdx !== -1 ? khTipeIdx : 7] ? dataKehadiran[i][khTipeIdx !== -1 ? khTipeIdx : 7].toString().trim() : "";
      
      let rowDisplayTipe = "Datang";
      if (rowTipeFull.toLowerCase().includes("manual")) {
        rowDisplayTipe = rowTipeFull.split(" (Manual: ")[0];
      } else if (rowTipeFull.toLowerCase() === 'pulang') {
        rowDisplayTipe = "Pulang";
      } else if (rowTipeFull.toLowerCase() === 'datang') {
        rowDisplayTipe = "Datang";
      } else {
        rowDisplayTipe = rowTipeFull;
      }

      if (rowEmpId === employeeId.toString().trim() && rowDate === formattedDate && rowDisplayTipe.toLowerCase() === inputDisplayTipe.toLowerCase()) {
        return { 
          success: false, 
          message: "Akses Ditolak! Anda sudah tercatat melakukan Absen " + inputDisplayTipe + " hari ini (" + formattedDate + ")." 
        };
      }
    }
    
    let mapLink = "-";
    if (latitude && longitude && latitude !== 0 && longitude !== 0) {
      mapLink = "https://www.google.com/maps?q=" + latitude + "," + longitude;
    }
    
    // Tulis baris data absensi lengkap dengan kolom Unit dan Shift baru
    sheetKehadiran.appendRow([
      Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss"), // Timestamp
      "'" + formattedDate, // Tanggal
      "'" + formattedTime, // Jam
      karyawan.id,            
      karyawan.nama,          
      karyawan.jabatan,
      karyawan.unit, // Unit
      type,                   
      selectedShift || "-", // Shift
      latitude || "-", 
      longitude || "-", 
      mapLink                 
    ]);
    
    SpreadsheetApp.flush();
    
    return {
      success: true,
      message: "Absensi " + type + " berhasil dicatat!",
      data: {
        nama: karyawan.nama,
        waktu: formattedTime,
        tanggal: formattedDate
      }
    };
    
  } catch (e) {
    return { success: false, message: "Terjadi kesalahan sistem: " + e.message };
  }
}

/**
 * Memvalidasi PIN Admin
 */
function verifyAdminPin(pin) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Pengaturan');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('Pengaturan');
    }
    
    const data = sheet.getDataRange().getValues();
    let savedPin = "1234";
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'ADMIN_PIN') {
        savedPin = data[i][1].toString().trim();
        break;
      }
    }
    return pin.toString().trim() === savedPin;
  } catch (e) {
    return false;
  }
}

/**
 * Helper Pembersih & Penormalisasi Format Tanggal yang Tangguh
 */
function cleanAndFormatDate(rawDate, timezone) {
  if (!rawDate) return "";
  
  if (rawDate instanceof Date) {
    return Utilities.formatDate(rawDate, timezone, "dd-MM-yyyy");
  }
  
  let dateStr = rawDate.toString().trim().replace(/^'/, "");
  if (dateStr === "" || dateStr.toLowerCase() === "tanggal") return "";
  
  if (/^\d{2}-\d{2}-\d{4}$/.test(dateStr)) {
    return dateStr;
  }
  
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    const parts = dateStr.split('-');
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  
  if (dateStr.includes('/')) {
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      if (parts[0].length === 4) { // yyyy/MM/dd -> dd-MM-yyyy
        return `${parts[2].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[0]}`;
      } else if (parts[2].length === 4) { // dd/MM/yyyy -> dd-MM-yyyy
        return `${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[2]}`;
      }
    }
  }
  
  try {
    const parsed = new Date(dateStr);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timezone, "dd-MM-yyyy");
    }
  } catch(e) {}
  
  return dateStr;
}

/**
 * Helper Pembersih Jam
 */
function cleanAndFormatTime(rawTime, timezone) {
  if (!rawTime) return "";
  if (rawTime instanceof Date) {
    return Utilities.formatDate(rawTime, timezone, "HH:mm:ss");
  }
  let timeStr = rawTime.toString().trim();
  if (timeStr === "" || timeStr.toLowerCase() === "jam") return "";
  return timeStr;
}

/**
 * Mengambil rekapitulasi data absensi secara real-time dengan index tajuk dinamis (Fail-Safe dengan Auto-Inisialisasi)
 */
function getAttendanceData() {
  try {
    SpreadsheetApp.flush();
    const ss = getSpreadsheet();
    let sheetKehadiran = ss.getSheetByName('Kehadiran');
    
    if (!sheetKehadiran) {
      setupDatabase();
      sheetKehadiran = ss.getSheetByName('Kehadiran');
    }
    
    const tz = ss.getSpreadsheetTimeZone();
    const data = sheetKehadiran.getDataRange().getValues();
    const records = [];
    
    const now = new Date();
    const todayFormatted = Utilities.formatDate(now, tz, "dd-MM-yyyy");
    
    // Temukan index tajuk kolom secara dinamis
    const headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    const timestampIdx = headers.indexOf('timestamp');
    const tanggalIdx = headers.indexOf('tanggal');
    const jamIdx = headers.indexOf('jam');
    const idIdx = headers.indexOf('id pegawai');
    const namaIdx = headers.indexOf('nama pegawai');
    const jbtIdx = headers.indexOf('jabatan');
    const unitIdx = headers.indexOf('unit');
    const tipeIdx = headers.indexOf('tipe absen');
    const shiftIdx = headers.indexOf('shift');
    const latIdx = headers.indexOf('latitude');
    const lngIdx = headers.indexOf('longitude');
    const linkIdx = headers.indexOf('link lokasi');
    
    for (let i = data.length - 1; i >= 1; i--) {
      const idPegawai = idIdx !== -1 ? data[i][idIdx].toString().trim() : "";
      const namaPegawai = namaIdx !== -1 ? data[i][namaIdx].toString().trim() : "";
      
      if (idPegawai === "" && namaPegawai === "") continue;
      if (idPegawai === "ID Pegawai" || namaPegawai === "Nama Pegawai" || namaPegawai === "Nama Lengkap") continue; 
      
      const formattedDate = cleanAndFormatDate(tanggalIdx !== -1 ? data[i][tanggalIdx] : data[i][1], tz);
      const formattedTime = cleanAndFormatTime(jamIdx !== -1 ? data[i][jamIdx] : data[i][2], tz);
      const rowTimestamp = timestampIdx !== -1 ? data[i][timestampIdx] : data[i][0];

      let formattedTimestamp = "";
      if (rowTimestamp instanceof Date) {
        formattedTimestamp = Utilities.formatDate(rowTimestamp, tz, "yyyy-MM-dd HH:mm:ss");
      } else {
        formattedTimestamp = rowTimestamp ? rowTimestamp.toString() : "";
      }
      
      records.push({
        timestamp: formattedTimestamp,
        tanggal: formattedDate,         
        jam: formattedTime,             
        idPegawai: idPegawai,
        nama: namaPegawai,
        jabatan: jbtIdx !== -1 ? data[i][jbtIdx].toString().trim() : "-",
        unit: unitIdx !== -1 ? data[i][unitIdx].toString().trim() : "-",
        tipe: tipeIdx !== -1 ? data[i][tipeIdx].toString().trim() : "-",
        shift: shiftIdx !== -1 ? data[i][shiftIdx].toString().trim() : "-",
        lat: latIdx !== -1 ? data[i][latIdx].toString().trim() : "-",
        lng: lngIdx !== -1 ? data[i][lngIdx].toString().trim() : "-",
        mapLink: linkIdx !== -1 ? data[i][linkIdx].toString().trim() : "-"
      });
    }
    
    return {
      records: records,
      today: todayFormatted
    };
  } catch (e) {
    throw new Error("Gagal mengambil data rekap: " + e.message);
  }
}

/**
 * MENGHAPUS DATA ABSENSI BERDASARKAN TIMESTAMP UNIK DAN ID PEGAWAI (KOREKSI MANUAL)
 */
function deleteAttendanceRecord(timestampStr, employeeId) {
  try {
    SpreadsheetApp.flush();
    const ss = getSpreadsheet();
    const sheetKehadiran = ss.getSheetByName('Kehadiran');
    if (!sheetKehadiran) return { success: false, message: "Sheet Kehadiran tidak ditemukan." };
    
    const tz = ss.getSpreadsheetTimeZone();
    const data = sheetKehadiran.getDataRange().getValues();
    const headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    const timestampIdx = headers.indexOf('timestamp');
    const idIdx = headers.indexOf('id pegawai');
    
    let targetRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      const rowTimestamp = timestampIdx !== -1 ? data[i][timestampIdx] : data[i][0];
      let formattedTimestamp = "";
      if (rowTimestamp instanceof Date) {
        formattedTimestamp = Utilities.formatDate(rowTimestamp, tz, "yyyy-MM-dd HH:mm:ss");
      } else {
        formattedTimestamp = rowTimestamp ? rowTimestamp.toString().trim() : "";
      }
      
      const rowEmpId = idIdx !== -1 ? data[i][idIdx].toString().trim() : "";
      
      if (formattedTimestamp === timestampStr && rowEmpId === employeeId.toString().trim()) {
        targetRow = i + 1; // Apps Script baris menggunakan index berbasis 1 (1-based index)
        break;
      }
    }
    
    if (targetRow !== -1) {
      sheetKehadiran.deleteRow(targetRow);
      SpreadsheetApp.flush(); // Paksa simpan database
      return { success: true, message: "Data absensi berhasil dihapus/dikoreksi dari database!" };
    } else {
      return { success: false, message: "Data absensi tidak ditemukan atau sudah terhapus." };
    }
  } catch (e) {
    return { success: false, message: "Gagal menghapus data: " + e.message };
  }
}
