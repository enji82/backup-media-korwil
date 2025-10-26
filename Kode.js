/**
 * ===================================================================
 * ======================= 1. KONFIGURASI PUSAT ======================
 * ===================================================================
 * Semua ID Spreadsheet dan Folder disimpan di sini agar mudah dikelola.
 */

const SPREADSHEET_CONFIG = {
  // --- Modul SK Pembagian Tugas ---
  SK_BAGI_TUGAS: { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "SK Tabel Kirim" },
  SK_FORM_RESPONSES: { id: "1AmvOJAhOfdx09eT54x62flWzBZ1xNQ8Sy5lzvT9zJA4", sheet: "Form Responses 1" },
  
  // --- Modul Laporan Bulanan ---
  LAPBUL_FORM_RESPONSES_PAUD: { id: "1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs", sheet: "Form Responses 1" },
  LAPBUL_FORM_RESPONSES_SD: { id: "1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s", sheet: "Input" },
  LAPBUL_RIWAYAT: { id: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU", sheet: "Riwayat" },
  LAPBUL_STATUS: { id: "1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU", sheet: "Status" },

  // --- Data Pendukung ---
  DATA_SEKOLAH: { id: "1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs", sheet: "Data Sekolah" },
  DATA_SEKOLAH_PAUD: { id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA" }, 
  DROPDOWN_DATA: { id: "1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA" },

  // --- Data SIABA ---
  SIABA_REKAP: { id: "1x3b-yzZbiqP2XfJNRC3XTbMmRTHLd8eEdUqAlKY3v9U", sheet: "Rekap Script" },
  SIABA_TIDAK_PRESENSI: { id: "1mjXz5l_cqBiiR3x9qJ7BU4yQ3f0ghERT9ph8CC608Zc", sheet: "Rekap Script" },
};

const FOLDER_CONFIG = {
  // --- Modul SK Pembagian Tugas ---
  MAIN_SK: "1GwIow8B4O1OWoq3nhpzDbMO53LXJJUKs",

  // --- Modul Laporan Bulanan ---
  LAPBUL_KB: "18CxRT-eledBGRtHW1lFd2AZ8Bub6q5ra",
  LAPBUL_TK: "1WUNz_BSFmcwRVlrG67D2afm9oJ-bVI9H",
  LAPBUL_SD: "1I8DRQYpBbTt1mJwtD1WXVD6UK51TC8El",
};
/**
 * ===================================================================
 * ===================== 2. FUNGSI INTI APLIKASI =====================
 * ===================================================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * ===================================================================
 * ===================== 3. FUNGSI UTILITAS UMUM =====================
 * ===================================================================
 */

function handleError(functionName, error) {
  Logger.log(`Error di ${functionName}: ${error.message}\nStack: ${error.stack}`);
  throw new Error(error.message);
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function getDataFromSheet(configKey) {
  try {
    const config = SPREADSHEET_CONFIG[configKey];
    if (!config) throw new Error(`Konfigurasi untuk '${configKey}' tidak ditemukan.`);
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error(`Sheet '${config.sheet}' tidak ditemukan.`);
    return sheet.getDataRange().getValues();
    // Selalu gunakan getValues() untuk konsistensi data tanggal
  } catch (e) {
    handleError(`getDataFromSheet: ${configKey}`, e);
  }
}

function getCachedData(key, fetchFunction) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached != null) {
    Logger.log("Mengambil data dari Cache: " + key);
    return JSON.parse(cached);
  }
  Logger.log("Mengambil data dari Spreadsheet dan menyimpan ke Cache: " + key);
  const freshData = fetchFunction();
  cache.put(key, JSON.stringify(freshData), 21600); // Simpan selama 6 jam
  return freshData;
}


/**
 * ===================================================================
 * ================= 4. MODUL SK PEMBAGIAN TUGAS =====================
 * ===================================================================
 */

function getMasterSkOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
    const getValuesFromSheet = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      return sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
                  .flat()
                  .filter(value => String(value).trim() !== '');
    };

    return {
      'Nama SD': getValuesFromSheet('Nama SD').sort(),
      'Tahun Ajaran': getValuesFromSheet('Tahun Ajaran').sort().reverse(),
      'Semester': getValuesFromSheet('Semester').sort(),
      'Kriteria SK': getValuesFromSheet('Kriteria SK').sort()
    };
  } catch (e) {
    return handleError('getMasterSkOptions', e);
  }
}

function processManualForm(formData) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);  
    
    const tahunAjaranFolderName = formData.tahunAjaran.replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);

    const semesterFolderName = formData.semester;
    const targetFolder = getOrCreateFolder(tahunAjaranFolder, semesterFolderName);

    const newFilename = `${formData.namaSD} - ${tahunAjaranFolderName} - ${formData.semester} - ${formData.kriteriaSK}.pdf`;
    
    const decodedData = Utilities.base64Decode(formData.fileData.data);
    const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
    const newFile = targetFolder.createFile(blob);
    const fileUrl = newFile.getUrl();
    const newRow = [ new Date(), formData.namaSD, formData.tahunAjaran, formData.semester, formData.nomorSK, new Date(formData.tanggalSK), formData.kriteriaSK, fileUrl ];
    
    sheet.appendRow(newRow);
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 6).setNumberFormat("dd-MM-yyyy");

    return "Dokumen SK berhasil diunggah.";
  } catch (e) {
    return handleError('processManualForm', e);
  }
}

function getRiwayatPengirimanSKData() {
  try {
    const data = getDataFromSheet('SK_FORM_RESPONSES');
    if (data.length < 2) return data;
    let headers = data[0].map(h => String(h).trim());
    let dataRows = data.slice(1);

    // Pindahkan kolom "Tanggal Unggah" ke paling akhir
    const tglUnggahIndex = headers.indexOf('Tanggal Unggah');
    if (tglUnggahIndex > -1) {
      const [tglUnggahHeader] = headers.splice(tglUnggahIndex, 1);
      headers.push(tglUnggahHeader);
      dataRows = dataRows.map(row => {
        const [tglUnggahData] = row.splice(tglUnggahIndex, 1);
        row.push(tglUnggahData);
        return row;
      });
    }

    // --- PERUBAHAN PENGURUTAN DI SINI ---
    // Dapatkan indeks baru dari "Tanggal Unggah" setelah dipindahkan
    const tglUnggahSortIndex = headers.indexOf('Tanggal Unggah');
    const tglSKIndex = headers.indexOf('Tanggal SK');
    
    const parseDate = (value) => {
        if (!value) return new Date(0);
        if (value instanceof Date && !isNaN(value)) return value;
        const date = new Date(value);
        return isNaN(date) ? new Date(0) : date;
    };
    
    // Urutkan dataRows berdasarkan kolom "Tanggal Unggah" (terbaru di atas)
    dataRows.sort((a, b) => {
      const dateA = parseDate(a[tglUnggahSortIndex]);
      const dateB = parseDate(b[tglUnggahSortIndex]);
      return dateB - dateA; // Mengurutkan descending
    });

    const formattedDataRows = dataRows.map(row => {
        return row.map((cell, index) => {
            if (cell instanceof Date) {
                if (index === tglSKIndex) {
                    return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
                } else {
      
                  return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
                }
            }
            return cell;
        });
    });

    return [headers].concat(formattedDataRows);

  } catch(e) {
    return handleError('getRiwayatPengirimanSKData', e);
  }
}

function getStatusPengirimanSKData() { 
  const data = getDataFromSheet('SK_BAGI_TUGAS');
  return data.map(row => row.map(cell => String(cell))); // Pastikan semua data adalah string
}

function getSKDataForManagement() {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet) throw new Error(`Sheet '${config.sheet}' tidak ditemukan.`);
    
    const originalData = sheet.getDataRange().getValues();
    if (originalData.length < 2) {
      return { headers: originalData.length > 0 ? originalData[0] : [], rows: [] };
    }
    
    // [PERUBAHAN 1] Mendefinisikan urutan header baru yang diinginkan
    const desiredHeaders = ["Nama SD", "Tahun Ajaran", "Semester", "Nomor SK", "Tanggal SK", "Kriteria SK", "Dokumen", "Tanggal Unggah", "Update"];
    const dataRows = originalData.slice(1);
    const originalHeadersMap = originalData[0].map(h => String(h).trim());

    const parseDate = (value) => {
        if (!value) return new Date(0);
        if (value instanceof Date && !isNaN(value)) return value;
        const date = new Date(value);
        return isNaN(date) ? new Date(0) : date;
    };
    
    const indexedData = dataRows.map((row, index) => ({ row: row, originalIndex: index + 2 }));

    // Logika pengurutan tidak berubah
    indexedData.sort((a, b) => {
      const updateIndex = originalHeadersMap.indexOf('Update');
      const timestampIndex = originalHeadersMap.indexOf('Tanggal Unggah');
      const dateB_update = (updateIndex > -1) ? parseDate(b.row[updateIndex]) : new Date(0);
      const dateA_update = (updateIndex > -1) ? parseDate(a.row[updateIndex]) : new Date(0);
      if (dateB_update.getTime() !== dateA_update.getTime()) { return dateB_update - dateA_update; }
      const dateB_timestamp = parseDate(b.row[timestampIndex]);
      const dateA_timestamp = parseDate(a.row[timestampIndex]);
     
      return dateB_timestamp - dateA_timestamp;
    });

    const formattedRows = indexedData.map(item => {
      const rowData = {};
      originalHeadersMap.forEach((header, i) => {
        let cell = item.row[i];
        if ((header === 'Tanggal Unggah' || header === 'Update' || header === 'Tanggal SK') && cell) {
          const dateObject = parseDate(cell);
          if (dateObject.getTime() !== 0) {
      
            const format = (header === 'Tanggal SK') ? "dd/MM/yyyy" : "dd/MM/yyyy HH:mm:ss";
            rowData[header] = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), format);
          } else {
            rowData[header] = '';
          }
        } else {
          rowData[header] = String(cell);
        }
 
      });
      return { rowIndex: item.originalIndex, data: rowData };
    });

    // [PERUBAHAN 2] Mengirimkan data dengan urutan header yang baru
    return { headers: desiredHeaders, rows: formattedRows };

  } catch (e) {
    return handleError("getSKDataForManagement", e);
  }
}

function getSKDataByRow(rowIndex) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header] = rowValues[i];
    });
    return rowData;
  } catch (e) {
    return handleError("getSKDataByRow", e);
  }
}

function updateSKData(formData) {
  try {
    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const range = sheet.getRange(formData.rowIndex, 1, 1, headers.length);
    const existingRowValues = range.getDisplayValues()[0]; // getDisplayValues untuk perbandingan string
    const existingRowObject = {};
    headers.forEach((header, i) => { existingRowObject[header] = existingRowValues[i]; });

    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.MAIN_SK);
    const tahunAjaranFolderName = existingRowObject['Tahun Ajaran'].replace(/\//g, '-');
    const tahunAjaranFolder = getOrCreateFolder(mainFolder, tahunAjaranFolderName);
    
    let fileUrl = existingRowObject['Dokumen'];
    const fileUrlIndex = headers.indexOf('Dokumen');

    // Tentukan nama file dan folder tujuan BARU berdasarkan data form
    const newSemesterFolderName = formData['Semester'];
    const newTargetFolder = getOrCreateFolder(tahunAjaranFolder, newSemesterFolderName);
    const newFilename = `${existingRowObject['Nama SD']} - ${tahunAjaranFolderName} - ${newSemesterFolderName} - ${formData['Kriteria SK']}.pdf`;

    // Skenario 1: File BARU diunggah
    if (formData.fileData && formData.fileData.data) {
      // Hapus file lama jika ada
      if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        try {
          const fileId = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
          if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch (e) {
          Logger.log(`Gagal menghapus file lama saat upload baru: ${e.message}`);
        }
      }
      
      // Unggah file baru
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFilename);
      const newFile = newTargetFolder.createFile(blob);
      fileUrl = newFile.getUrl();

    // Skenario 2: TIDAK ada file baru, tapi data (Semester/Kriteria) berubah
    } else if (fileUrlIndex > -1 && existingRowObject['Dokumen']) {
        const fileIdMatch = existingRowObject['Dokumen'].match(/[-\w]{25,}/);
        if (fileIdMatch) {
            const fileId = fileIdMatch[0];
            const file = DriveApp.getFileById(fileId);
            const currentFileName = file.getName();
            const currentParentFolder = file.getParents().next();

            // Cek jika nama file atau folder perlu diubah
            if (currentFileName !== newFilename || currentParentFolder.getName() !== newSemesterFolderName) {
                file.moveTo(newTargetFolder);
                file.setName(newFilename);
                fileUrl = file.getUrl(); // Perbarui URL setelah dipindah/diubah namanya
                Logger.log(`File dipindahkan ke folder '${newSemesterFolderName}' dan diubah namanya menjadi '${newFilename}'`);
            }
        }
    }
    
    formData['Dokumen'] = fileUrl; // Pastikan URL file diperbarui di formData
    formData['Update'] = new Date();

    const newRowValuesForSheet = headers.map(header => {
      return formData.hasOwnProperty(header) ? formData[header] : existingRowObject[header];
    });

    // Gunakan getRange(row, col, numRows, numCols) dan setValues()
    sheet.getRange(formData.rowIndex, 1, 1, headers.length).setValues([newRowValuesForSheet]);
    
    const tanggalSKIndex = headers.indexOf('Tanggal SK');
    if (tanggalSKIndex !== -1) {
      sheet.getRange(formData.rowIndex, tanggalSKIndex + 1).setNumberFormat("dd-MM-yyyy");
    }
    
    return "Data berhasil diperbarui!";
  } catch (e) {
    return handleError('updateSKData', e);
  }
}


function deleteSKData(rowIndex, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    const config = SPREADSHEET_CONFIG.SK_FORM_RESPONSES;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fileUrlIndex = headers.findIndex(h => h.trim().toLowerCase() === 'dokumen');
    
    if (fileUrlIndex !== -1) {
        const fileUrl = sheet.getRange(rowIndex, fileUrlIndex + 1).getValue();
        if (fileUrl && typeof fileUrl === 'string') {
            const fileId = fileUrl.match(/[-\w]{25,}/);
            if (fileId) {
                try {
                    DriveApp.getFileById(fileId[0]).setTrashed(true);
                } catch (err) {
                    Logger.log(`Gagal menghapus file dengan ID ${fileId[0]}: ${err.message}`);
                }
            }
        }
    }
    
    sheet.deleteRow(rowIndex);
    return "Data dan file terkait berhasil dihapus.";
  } catch (e) {
    return handleError("deleteSKData", e);
  }
}


/**
 * ===================================================================
 * ==================== 5. MODUL LAPORAN BULAN =======================
 * ===================================================================
 */

function getPaudSchoolLists() {
  const cacheKey = 'paud_school_lists';
  return getCachedData(cacheKey, function() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DATA_SEKOLAH_PAUD.id);
    const getValues = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().filter(Boolean).sort();
    };
    return { KB: getValues('Nama KB'), TK: getValues('Nama TK') };
  });
}

function getSdSchoolLists() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DATA_SEKOLAH_PAUD.id);
    const getValues = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().filter(Boolean).sort();
    };
    return {
      SDN: getValues('Nama SDN'),
      SDS: getValues('Nama SDS')
    };
  } catch (e) {
    return handleError('getSdSchoolLists', e);
  }
}

function processLapbulFormPaud(formData) {
  try {
    const jenjang = formData.jenjang;
    let FOLDER_ID_LAPBUL;
    if (jenjang === 'KB') {
      FOLDER_ID_LAPBUL = FOLDER_CONFIG.LAPBUL_KB;
    } else if (jenjang === 'TK') {
      FOLDER_ID_LAPBUL = FOLDER_CONFIG.LAPBUL_TK;
    } else {
      throw new Error("Jenjang tidak valid: " + jenjang);
    }
    
    const mainFolder = DriveApp.getFolderById(FOLDER_ID_LAPBUL);
    const tahunFolder = getOrCreateFolder(mainFolder, formData.tahun);
    const bulanFolder = getOrCreateFolder(tahunFolder, formData.laporanBulan);

    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    const newFileName = `${formData.namaSekolah} - Lapbul ${formData.laporanBulan} ${formData.tahun}.pdf`;
    const newFile = bulanFolder.createFile(blob).setName(newFileName);
    const fileUrl = newFile.getUrl();

    const config = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);

    const newRow = [
      new Date(),
      formData.laporanBulan, formData.tahun, formData.npsn, formData.statusSekolah, formData.jumlahRombel, formData.jenjang, formData.namaSekolah,
      formData.murid_0_1_L, formData.murid_0_1_P, formData.murid_1_2_L, formData.murid_1_2_P,
      formData.murid_2_3_L, formData.murid_2_3_P, formData.murid_3_4_L, formData.murid_3_4_P,
      formData.murid_4_5_L, formData.murid_4_5_P, formData.murid_5_6_L, formData.murid_5_6_P,
      formData.murid_6_up_L, formData.murid_6_up_P, formData.kelompok_A_L, formData.kelompok_A_P,
      formData.kelompok_B_L, formData.kelompok_B_P, formData.kepsek_ASN, formData.kepsek_Non_ASN,
      formData.guru_PNS, formData.guru_PPPK, formData.guru_GTY, formData.guru_GTT,
      formData.tendik_Penjaga, formData.tendik_TAS, formData.tendik_Pustakawan, formData.tendik_Lainnya,
      fileUrl
    ];
    sheet.appendRow(newRow);

    return "Sukses! Laporan Bulan PAUD berhasil dikirim.";
  } catch (e) {  
    return handleError('processLapbulFormPaud', e);
  }
}

function processLapbulFormSd(formData) {
  try {
    // 1. Proses File dan Simpan ke Google Drive
    const mainFolder = DriveApp.getFolderById(FOLDER_CONFIG.LAPBUL_SD);
    const tahunFolder = getOrCreateFolder(mainFolder, formData.tahun);
    const bulanFolder = getOrCreateFolder(tahunFolder, formData.laporanBulan);

    const fileData = formData.fileData;
    const decodedData = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
    
    const newFileName = `${formData.namaSekolah} - Lapbul ${formData.laporanBulan} ${formData.tahun}.pdf`;
    const newFile = bulanFolder.createFile(blob).setName(newFileName);
    const fileUrl = newFile.getUrl();

    // 2. Siapkan Data untuk Disimpan ke Google Sheet
    const config = SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    const getValue = (key) => formData[key] || 0;

    // Susun baris baru sesuai urutan kolom A s/d GH
    const newRow = [
      new Date(), // A
      formData.laporanBulan, // B
      formData.tahun, // C
      formData.statusSekolah, // D
      formData.namaSekolah, // E
      formData.npsn, // F
      formData.jumlahRombel, // G
      fileUrl, // H
      
      // Kelas 1 (I-AC)
      getValue('k1_jumlah_rombel'), getValue('k1_rombel_tunggal_L'), getValue('k1_rombel_tunggal_P'),
      getValue('k1_rombel_a_L'), getValue('k1_rombel_a_P'), getValue('k1_rombel_b_L'), getValue('k1_rombel_b_P'), getValue('k1_rombel_c_L'), getValue('k1_rombel_c_P'),
      getValue('k1_agama_islam_L'), getValue('k1_agama_islam_P'), getValue('k1_agama_kristen_L'), getValue('k1_agama_kristen_P'), getValue('k1_agama_katolik_L'), getValue('k1_agama_katolik_P'), 
      getValue('k1_agama_hindu_L'), getValue('k1_agama_hindu_P'), getValue('k1_agama_buddha_L'), getValue('k1_agama_buddha_P'), getValue('k1_agama_konghucu_L'), getValue('k1_agama_konghucu_P'),
      
      // Kelas 2 (AD-AX)
      getValue('k2_jumlah_rombel'), getValue('k2_rombel_tunggal_L'), getValue('k2_rombel_tunggal_P'),
      getValue('k2_rombel_a_L'), getValue('k2_rombel_a_P'), getValue('k2_rombel_b_L'), getValue('k2_rombel_b_P'), getValue('k2_rombel_c_L'), getValue('k2_rombel_c_P'),
      getValue('k2_agama_islam_L'), getValue('k2_agama_islam_P'), getValue('k2_agama_kristen_L'), getValue('k2_agama_kristen_P'), getValue('k2_agama_katolik_L'), getValue('k2_agama_katolik_P'), 
      getValue('k2_agama_hindu_L'), getValue('k2_agama_hindu_P'), getValue('k2_agama_buddha_L'), getValue('k2_agama_buddha_P'), getValue('k2_agama_konghucu_L'), getValue('k2_agama_konghucu_P'),

    
      // Kelas 3 (AY-BS)
      getValue('k3_jumlah_rombel'), getValue('k3_rombel_tunggal_L'), getValue('k3_rombel_tunggal_P'),
      getValue('k3_rombel_a_L'), getValue('k3_rombel_a_P'), getValue('k3_rombel_b_L'), getValue('k3_rombel_b_P'), getValue('k3_rombel_c_L'), getValue('k3_rombel_c_P'),
      getValue('k3_agama_islam_L'), getValue('k3_agama_islam_P'), getValue('k3_agama_kristen_L'), getValue('k3_agama_kristen_P'), getValue('k3_agama_katolik_L'), getValue('k3_agama_katolik_P'), 
      getValue('k3_agama_hindu_L'), getValue('k3_agama_hindu_P'), getValue('k3_agama_buddha_L'), getValue('k3_agama_buddha_P'), getValue('k3_agama_konghucu_L'), getValue('k3_agama_konghucu_P'),
      
      // Kelas 4 (BT-CN)
      getValue('k4_jumlah_rombel'), getValue('k4_rombel_tunggal_L'), getValue('k4_rombel_tunggal_P'),
      getValue('k4_rombel_a_L'), getValue('k4_rombel_a_P'), getValue('k4_rombel_b_L'), getValue('k4_rombel_b_P'), getValue('k4_rombel_c_L'), getValue('k4_rombel_c_P'),
      getValue('k4_agama_islam_L'), getValue('k4_agama_islam_P'), getValue('k4_agama_kristen_L'), getValue('k4_agama_kristen_P'), getValue('k4_agama_katolik_L'), getValue('k4_agama_katolik_P'), 
      getValue('k4_agama_hindu_L'), 
      getValue('k4_agama_hindu_P'), getValue('k4_agama_buddha_L'), getValue('k4_agama_buddha_P'), getValue('k4_agama_konghucu_L'), getValue('k4_agama_konghucu_P'),

      // Kelas 5 (CO-DI)
      getValue('k5_jumlah_rombel'), getValue('k5_rombel_tunggal_L'), getValue('k5_rombel_tunggal_P'),
      getValue('k5_rombel_a_L'), getValue('k5_rombel_a_P'), getValue('k5_rombel_b_L'), getValue('k5_rombel_b_P'), getValue('k5_rombel_c_L'), getValue('k5_rombel_c_P'),
      getValue('k5_agama_islam_L'), getValue('k5_agama_islam_P'), getValue('k5_agama_kristen_L'), getValue('k5_agama_kristen_P'), getValue('k5_agama_katolik_L'), getValue('k5_agama_katolik_P'), 
      getValue('k5_agama_hindu_L'), getValue('k5_agama_hindu_P'), getValue('k5_agama_buddha_L'), getValue('k5_agama_buddha_P'), getValue('k5_agama_konghucu_L'), getValue('k5_agama_konghucu_P'),

      // Kelas 6 (DJ-ED)
      getValue('k6_jumlah_rombel'), getValue('k6_rombel_tunggal_L'), getValue('k6_rombel_tunggal_P'),
      getValue('k6_rombel_a_L'), getValue('k6_rombel_a_P'), getValue('k6_rombel_b_L'), getValue('k6_rombel_b_P'), getValue('k6_rombel_c_L'), getValue('k6_rombel_c_P'),
      getValue('k6_agama_islam_L'), getValue('k6_agama_islam_P'), getValue('k6_agama_kristen_L'), getValue('k6_agama_kristen_P'), getValue('k6_agama_katolik_L'), getValue('k6_agama_katolik_P'), 
     
      getValue('k6_agama_hindu_L'), getValue('k6_agama_hindu_P'), getValue('k6_agama_buddha_L'), getValue('k6_agama_buddha_P'), getValue('k6_agama_konghucu_L'), getValue('k6_agama_konghucu_P'),

      // PTK Negeri (EE-FH)
      getValue('ptk_negeri_ks_pns'), getValue('ptk_negeri_ks_pppk'),
      getValue('ptk_negeri_guru_kelas_pns'), getValue('ptk_negeri_guru_kelas_pppk'), getValue('ptk_negeri_guru_kelas_pppkpw'), getValue('ptk_negeri_guru_kelas_gtt'),
      getValue('ptk_negeri_guru_pai_pns'), getValue('ptk_negeri_guru_pai_pppk'), getValue('ptk_negeri_guru_pai_pppkpw'), getValue('ptk_negeri_guru_pai_gtt'),
      getValue('ptk_negeri_guru_pjok_pns'), getValue('ptk_negeri_guru_pjok_pppk'), getValue('ptk_negeri_guru_pjok_pppkpw'), getValue('ptk_negeri_guru_pjok_gtt'),
      getValue('ptk_negeri_guru_kristen_pns'), getValue('ptk_negeri_guru_kristen_pppk'), getValue('ptk_negeri_guru_kristen_pppkpw'), getValue('ptk_negeri_guru_kristen_gtt'),
      getValue('ptk_negeri_guru_inggris_gtt'), getValue('ptk_negeri_guru_lainnya_gtt'),
      getValue('ptk_negeri_tendik_operator_pppk'), getValue('ptk_negeri_tendik_operator_pppkpw'), getValue('ptk_negeri_tendik_operator_ptt'),
      getValue('ptk_negeri_tendik_pengelola_pppk'), getValue('ptk_negeri_tendik_pengelola_pppkpw'), getValue('ptk_negeri_tendik_pengelola_ptt'),
      getValue('ptk_negeri_tendik_penjaga_ptt'), getValue('ptk_negeri_tendik_tas_ptt'), getValue('ptk_negeri_tendik_pustakawan_ptt'), getValue('ptk_negeri_tendik_lainnya_ptt'),

      // PTK Swasta (FI-GH)
 
      getValue('ptk_swasta_ks_gty'), getValue('ptk_swasta_ks_gtt'),
      getValue('ptk_swasta_guru_kelas_gty'), getValue('ptk_swasta_guru_kelas_gtt'),
      getValue('ptk_swasta_guru_pai_gty'), getValue('ptk_swasta_guru_pai_gtt'),
      getValue('ptk_swasta_guru_pjok_gty'), getValue('ptk_swasta_guru_pjok_gtt'),
      getValue('ptk_swasta_guru_kristen_gty'), getValue('ptk_swasta_guru_kristen_gtt'),
      getValue('ptk_swasta_guru_inggris_gty'), getValue('ptk_swasta_guru_inggris_gtt'),
      getValue('ptk_swasta_guru_lainnya_gty'), getValue('ptk_swasta_guru_lainnya_gtt'),
      getValue('ptk_swasta_tendik_operator_pty'), getValue('ptk_swasta_tendik_operator_ptt'),
      getValue('ptk_swasta_tendik_pengelola_pty'), getValue('ptk_swasta_tendik_pengelola_ptt'),
      getValue('ptk_swasta_tendik_penjaga_pty'), getValue('ptk_swasta_tendik_penjaga_ptt'),
      getValue('ptk_swasta_tendik_tas_pty'), getValue('ptk_swasta_tendik_tas_ptt'),
      getValue('ptk_swasta_tendik_pustakawan_pty'), getValue('ptk_swasta_tendik_pustakawan_ptt'),
      getValue('ptk_swasta_tendik_lainnya_pty'), getValue('ptk_swasta_tendik_lainnya_ptt')
    ];
    sheet.appendRow(newRow);

    return "Sukses! Laporan Bulan SD berhasil dikirim.";
  } catch (e) {
    return handleError('processLapbulFormSd', e);
  }
}

function getLapbulRiwayatData() {
  try {
    const sheet = SpreadsheetApp.openById("1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU").getSheetByName("Riwayat");

    if (!sheet) {
      throw new Error("Sheet 'Riwayat' di spreadsheet gabungan tidak ditemukan.");
    }

    const allData = sheet.getDataRange().getDisplayValues();

    // KUNCI PERUBAHAN 1: Tentukan urutan kolom final yang Anda inginkan di sini.
    // "Jenjang" tetap ada di akhir agar logika penyembunyian kolom di JavaScript tetap berfungsi.
    const desiredHeaders = ["Nama Sekolah", "Status", "Bulan", "Tahun", "Rombel", "Dokumen", "Tanggal Unggah", "Jenjang"];

    if (allData.length < 2) {
      return [desiredHeaders]; // Jika data kosong, kirim header dengan urutan yang benar.
    }

    const sourceHeaders = allData[0].map(h => h.trim());
    const dataRows = allData.slice(1);

    // KUNCI PERBAIKAN 2: Buat peta untuk menemukan posisi kolom asli di spreadsheet.
    const headerMap = {};
    sourceHeaders.forEach((header, index) => {
      headerMap[header] = index;
    });

    // KUNCI PERBAIKAN 3: Susun ulang setiap baris data sesuai urutan 'desiredHeaders'.
    const reorderedDataRows = dataRows.map(row => {
      const newRow = [];
      desiredHeaders.forEach(header => {
        const sourceIndex = headerMap[header];
        // Jika kolom ditemukan di sumber, ambil datanya. Jika tidak, beri placeholder '-'.
        newRow.push(sourceIndex !== undefined ? row[sourceIndex] : '-');
      });
      return newRow;
    });

    // Kirim header baru dan data yang sudah disusun ulang.
    // Proses pengurutan berdasarkan tanggal akan tetap dilakukan di sisi browser seperti sebelumnya.
    return [desiredHeaders].concat(reorderedDataRows);

  } catch(e) {
    Logger.log(`Error in getLapbulRiwayatData: ${e.message}\nStack: ${e.stack}`);
    throw new Error(`Terjadi error di server: ${e.message}`);
  }
}

function getLapbulStatusData() {
  try {
    // Membuka spreadsheet dan sheet "Status" yang benar.
    const sheet = SpreadsheetApp.openById("1aKEIkhKApmONrCg-QQbMhXyeGDJBjCZrhR-fvXZFtJU").getSheetByName("Status");
    
    if (!sheet) {
      throw new Error("Sheet 'Status' tidak ditemukan.");
    }
    
    // Langsung ambil dan kirim semua data yang terlihat (getDisplayValues).
    // Proses pemilihan kolom dan filter akan dilakukan di javascript.html
    return sheet.getDataRange().getDisplayValues();

  } catch (e) {
    Logger.log(`Error in getLapbulStatusData: ${e.message}`);
    throw new Error(`Terjadi error di server: ${e.message}`);
  }
}

function getLapbulKelolaData() {
  try {
    const paudSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName("Form Responses 1");
    const sdSheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName("Input");
    const finalHeaders = ["Nama Sekolah", "Status", "Jenjang", "Jumlah Rombel", "Bulan", "Tahun", "Dokumen", "Tanggal Unggah", "Update"];
    let combinedData = [];

    const parseDateForSort = (dateStr) => {
        if (!dateStr || !(typeof dateStr === 'string' || dateStr instanceof Date)) return new Date(0);
        if (dateStr instanceof Date) return dateStr;
        const parts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})\s(\d{2}):(\d{2}):(\d{2})/);
        if (parts) {
            return new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5], parts[6]);
        }
        const dateOnlyParts = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
        if (dateOnlyParts) {
            return new Date(dateOnlyParts[3], dateOnlyParts[2] - 1, dateOnlyParts[1]);
        }
        return new Date(0);
    };

    const processSheetData = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim());
      const rows = data.slice(1);

      const formatDate = (cell) => {
          if (cell instanceof Date) {
              return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          }
          return cell;
      };

      rows.forEach((row, index) => {
        const timestampCell = row[headers.indexOf("Tanggal Unggah")];
        if (!timestampCell) return;

        const rowData = {};
        headers.forEach((header, i) => {
            rowData[header] = (header === "Tanggal Unggah" || header === "Update") ? formatDate(row[i]) : row[i];
        });

        if (sourceName === 'SD') {
          rowData['Jenjang'] = 'SD';
        }
        
        combinedData.push({
          rowIndex: index + 2,
          source: sourceName,
          data: rowData
        });
      });
    };

    processSheetData(paudSheet, 'PAUD');
    processSheetData(sdSheet, 'SD');
    
    // KUNCI PERBAIKAN: Logika pengurutan diubah agar HANYA menggunakan "Tanggal Unggah"
    combinedData.sort((a, b) => {
        const dateB = parseDateForSort(b.data['Tanggal Unggah']);
        const dateA = parseDateForSort(a.data['Tanggal Unggah']);
        return dateB - dateA; // Mengurutkan dari yang terbaru ke terlama
    });

    return { headers: finalHeaders, rows: combinedData };

  } catch (e) {
    Logger.log(`Error in getLapbulKelolaData: ${e.message}\nStack: ${e.stack}`);
    return { error: `Terjadi error di server: ${e.message}` };
  }
}

/**
 * Mengambil satu baris data Laporan Bulan berdasarkan sumber dan nomor baris.
 * @param {number} rowIndex Nomor baris di spreadsheet.
 * @param {string} source Sumber data ('PAUD' atau 'SD').
 * @returns {Object} Objek yang berisi data dari baris tersebut.
 */
function getLapbulDataByRow(rowIndex, source) {
  try {
    let sheet;
    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const rowData = {};
    headers.forEach((header, i) => {
      rowData[header.trim()] = values[i];
    });
    return rowData;
  } catch (e) {
    return handleError('getLapbulDataByRow', e);
  }
}


/**
 * Memperbarui satu baris data Laporan Bulan.
 * @param {Object} formData Objek data dari form edit di client.
 * @returns {string} Pesan sukses.
 */
function updateLapbulData(formData) {
  try {
    let sheet, FOLDER_ID;
    const source = formData.source;
    const rowIndex = formData.rowIndex;
    if (!source || !rowIndex) throw new Error("Informasi 'source' atau 'rowIndex' tidak ditemukan.");

    // KUNCI PERBAIKAN: Blok ini dipindahkan ke atas agar variabel 'sheet'
    // terdefinisi sebelum digunakan.
    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
      const jenjang = formData.jenjang || sheet.getRange(rowIndex, headers.indexOf('Jenjang') + 1).getValue();
      FOLDER_ID = jenjang === 'KB' ? FOLDER_CONFIG.LAPBUL_KB : FOLDER_CONFIG.LAPBUL_TK;
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
      FOLDER_ID = FOLDER_CONFIG.LAPBUL_SD;
    } else {
      throw new Error("Sumber data tidak valid: " + source);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());

    if (!formData.laporanBulan || !formData.tahun) {
        throw new Error("Informasi 'laporanBulan' atau 'tahun' kosong. Tidak dapat memproses file.");
    }

    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const existingValues = range.getValues()[0];

    if (formData.fileData && formData.fileData.data) {
      const docHeaderName = 'Dokumen';
      const docIndex = headers.indexOf(docHeaderName);
      if (docIndex !== -1 && existingValues[docIndex]) {
        try {
          const fileId = String(existingValues[docIndex]).match(/[-\w]{25,}/);
          if (fileId) DriveApp.getFileById(fileId[0]).setTrashed(true);
        } catch (e) { Logger.log(`Gagal menghapus file lama: ${e.message}`); }
      }
      
      const mainFolder = DriveApp.getFolderById(FOLDER_ID);
      const tahunFolder = getOrCreateFolder(mainFolder, formData.tahun);
      const bulanFolder = getOrCreateFolder(tahunFolder, formData.laporanBulan);
      
      const namaSekolah = formData.namaSekolah || existingValues[headers.indexOf('Nama Sekolah')];
      const newFileName = `${namaSekolah} - Lapbul ${formData.laporanBulan} ${formData.tahun}.pdf`;
      const decodedData = Utilities.base64Decode(formData.fileData.data);
      const blob = Utilities.newBlob(decodedData, formData.fileData.mimeType, newFileName);
      const newFile = bulanFolder.createFile(blob);
      formData[docHeaderName] = newFile.getUrl();
    }

    formData['Update'] = new Date();
    const newRowValues = headers.map((header, index) => {
      if (formData.hasOwnProperty(header)) {
        return formData[header];
      }
      return existingValues[index];
    });
    range.setValues([newRowValues]);

    const updateColIndex = headers.indexOf('Update');
    if (updateColIndex !== -1) {
      sheet.getRange(rowIndex, updateColIndex + 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
    
    return "Data berhasil diperbarui.";
  } catch (e) {
    Logger.log(`Error in updateLapbulData: ${e.message}\nStack: ${e.stack}`);
    return { error: `Terjadi error di server: ${e.message}` };
  }
}

/**
 * Menghapus data Laporan Bulan dari sheet dan file terkait dari Drive.
 * @param {number} rowIndex Nomor baris di spreadsheet.
 * @param {string} source Sumber data ('PAUD' atau 'SD').
 * @param {string} deleteCode Kode hapus (yyyyMMdd).
 * @returns {string} Pesan sukses.
 */
function deleteLapbulData(rowIndex, source, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    let sheet;
    if (source === 'PAUD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_PAUD.sheet);
    } else if (source === 'SD') {
      sheet = SpreadsheetApp.openById(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.id).getSheetByName(SPREADSHEET_CONFIG.LAPBUL_FORM_RESPONSES_SD.sheet);
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const docIndex = headers.findIndex(h => h.trim().toLowerCase() === 'dokumen');
    if (docIndex !== -1) {
        const fileUrl = sheet.getRange(rowIndex, docIndex + 1).getValue();
        if (fileUrl && typeof fileUrl === 'string') {
            const fileId = fileUrl.match(/[-\w]{25,}/);
            if (fileId) {
                try {
                    DriveApp.getFileById(fileId[0]).setTrashed(true);
                } catch (err) {
                    Logger.log(`Gagal menghapus file Laporan Bulan dengan ID ${fileId[0]}: ${err.message}`);
                }
            }
        }
    }
    
    sheet.deleteRow(rowIndex);
    return "Data Laporan Bulan dan file terkait berhasil dihapus.";
  } catch (e) {
    return handleError("deleteLapbulData", e);
  }
}

/**
 * ===================================================================
 * =================== 6. MODUL GOOGLE DRIVE (ARSIP) =================
 * ===================================================================
 */

function getFolders(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const subFolders = parentFolder.getFolders();
    const folderList = [];
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      folderList.push({
        id: folder.getId(),
        name: folder.getName()
      });
    }
    folderList.sort((a, b) => b.name.localeCompare(a.name));
    return folderList;
  } catch (e) {
    return handleError("getFolders", e);
  }
}

function getFiles(folderId) {
  try {
    const parentFolder = DriveApp.getFolderById(folderId);
    const files = parentFolder.getFiles();
    const fileList = [];
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        name: file.getName(),
        url: file.getUrl()
      });
    }
    fileList.sort((a, b) => a.name.localeCompare(b.name));
    return fileList;
  } catch (e) {
    return handleError("getFiles", e);
  }
}

function getLapbulInfo() {
  const cacheKey = 'lapbul_info_v1'; // Kunci unik untuk cache
  return getCachedData(cacheKey, function() {
    try {
      // Menggunakan kembali konfigurasi DROPDOWN_DATA karena ID spreadsheet sama
      const ss = SpreadsheetApp.openById(SPREADSHEET_CONFIG.DROPDOWN_DATA.id);
      const sheet = ss.getSheetByName('Informasi');
      
      if (!sheet) {
        throw new Error("Sheet 'Informasi' tidak ditemukan di spreadsheet referensi.");
      }

      const lastRow = sheet.getLastRow();
 
      if (lastRow < 2) return []; // Kembalikan array kosong jika tidak ada data

      // Ambil data dari A2 sampai baris terakhir yang berisi konten
      const range = sheet.getRange('A2:A' + lastRow);
      const values = range.getValues()
                          .flat() // Mengubah array 2D menjadi 1D
              
              .filter(item => String(item).trim() !== ''); // Menghapus baris kosong
      return values;
    } catch (e) {
      return handleError('getLapbulInfo', e);
    }
  });
}

function getUnduhFormatInfo() {
  const cacheKey = 'unduh_format_info_v1';
  return getCachedData(cacheKey, function() {
    try {
      const ss = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
      const sheet = ss.getSheetByName('Informasi');
      
      if (!sheet) {
        throw new Error("Sheet 'Informasi' tidak ditemukan.");
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];

      // Ambil data dari B2 sampai baris 
      const range = sheet.getRange('B2:B' + lastRow);
      const values = range.getDisplayValues()
                          .flat()
                          .filter(item => String(item).trim() !== '');
      return values;
    } catch (e) {
      return handleError('getUnduhFormatInfo', e);
  
    }
  });
}

function getKeadaanPtkPaudData() {
  try {
    // ID Spreadsheet dari URL yang Anda berikan
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Keadaan PTK PAUD");
    
    if (!sheet) {
      throw new Error("Sheet 'Keadaan PTK PAUD' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getKeadaanPtkPaudData', e);
  }
}

function getKeadaanPtkSdData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("Keadaan PTK SD");
    if (!sheet) {
      throw new Error("Sheet 'Keadaan PTK SD' tidak ditemukan.");
    }
    
    return sheet.getDataRange().getDisplayValues();

  } catch (e) {
    return handleError('getKeadaanPtkSdData', e);
  }
}

/**
 * Mengambil data jumlah PTK bulanan untuk PAUD dari spreadsheet.
 */
function getJumlahPtkPaudBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Data PAUD Bulanan");
    if (!sheet) {
      throw new Error("Sheet 'Data PAUD Bulanan' tidak ditemukan.");
    }
    
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    // [PERBAIKAN] Menambahkan Kolom D (Bulan) dan E (Tahun) untuk filter
    // Indeks kolom yang diinginkan (A=0, B=1, C=2, D=3, E=4, F=5, dst.)
    const colIndices = [0, 1, 2, 3, 4, 5, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 35];
    // Filter header sesuai kolom yang diinginkan
    const finalHeaders = colIndices.map(index => headers[index]);

    // Filter setiap baris untuk hanya menyertakan data dari kolom yang diinginkan
    const finalData = dataRows.map(row => {
      return colIndices.map(index => row[index]);
    });

    return [finalHeaders].concat(finalData);

  } catch (e) {
    return handleError('getJumlahPtkPaudBulananData', e);
  }
}

/**
 * Mengambil data daftar PTK PAUD dari spreadsheet.
 */
function getDaftarPtkPaudData() {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) {
      throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");
    }
    
    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    let dataRows = allData.slice(1);

    // [PERBAIKAN] Mengurutkan baris data berdasarkan abjad pada kolom "Nama" (indeks 2)
    dataRows.sort((a, b) => {
      // localeCompare digunakan untuk pengurutan abjad yang benar
      return a[2].localeCompare(b[2]);
    });

    const colIndices = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    const finalHeaders = colIndices.map(index => headers[index]);
    
    const finalData = dataRows.map(row => {
      return colIndices.map(index => row[index]);
    });

    return [finalHeaders].concat(finalData);

  } catch (e) {
    return handleError('getDaftarPtkPaudData', e);
  }
}

function getKelolaPtkPaudData() {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const allData = sheet.getDataRange().getValues();
    if (allData.length < 2) {
      return { headers: [], rows: [] };
    }

    const headers = allData[0].map(h => String(h).trim());
    const dataRows = allData.slice(1);

    // [PERBAIKAN] Langkah 1: Simpan nomor baris asli SEBELUM diurutkan
    const indexedData = dataRows.map((row, index) => ({
      originalRowIndex: index + 2, // Baris fisik di spreadsheet (mulai dari 2)
      rowData: row
    }));

    // Fungsi bantu untuk mengubah tanggal menjadi objek yang bisa dibandingkan
    const parseDate = (value) => {
      if (value instanceof Date && !isNaN(value)) {
        return value.getTime();
      }
      return 0; 
    };

    // Langkah 2: Lakukan pengurutan 2 tingkat pada data yang sudah diindeks
    const updateIndex = headers.indexOf('Update');
    const dateInputIndex = headers.indexOf('Tanggal Input');
    
    indexedData.sort((a, b) => {
      // Urutan 1: Berdasarkan kolom "Update" (terbaru di atas)
      const updateA = parseDate(a.rowData[updateIndex]);
      const updateB = parseDate(b.rowData[updateIndex]);
      if (updateB !== updateA) {
        return updateB - updateA;
      }

      // Urutan 2: Jika "Update" sama, urutkan berdasarkan "Tanggal Input" (terbaru di atas)
      const dateInputA = parseDate(a.rowData[dateInputIndex]);
      const dateInputB = parseDate(b.rowData[dateInputIndex]);
      return dateInputB - dateInputA;
    });

    // Langkah 3: Buat data final untuk dikirim ke frontend
    const finalData = indexedData.map(item => {
      const rowDataObject = {};
      headers.forEach((header, i) => {
        let cell = item.rowData[i];
        if (cell instanceof Date) {
          rowDataObject[header] = Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        } else {
          rowDataObject[header] = cell || "";
        }
      });
      
      // [KUNCI PERBAIKAN] Kirim nomor baris asli yang sudah disimpan sebelumnya
      return { rowIndex: item.originalRowIndex, data: rowDataObject }; 
    });

    return { headers, rows: finalData };
  } catch (e) {
    Logger.log(`Error getKelolaPtkPaudData: ${e.message}`);
    return { error: e.message, headers: [], rows: [] };
  }
}

/**
 * Mengambil data satu baris PTK PAUD berdasarkan nomor barisnya.
 */
function getPtkPaudDataByRow(rowIndex) {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      // Untuk tanggal, ambil nilai mentahnya agar bisa di-format di form
      if (header === 'TMT' && values[i] instanceof Date) {
        rowData[header] = Utilities.formatDate(values[i], "UTC", "yyyy-MM-dd");
      } else {
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    Logger.log(`Error in getPtkPaudDataByRow: ${e.message}`);
    return { error: e.message };
  }
}

function updatePtkPaudData(formData) {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const rowIndex = formData.rowIndex;
    if (!rowIndex) throw new Error("Nomor baris (rowIndex) tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const oldValues = range.getValues()[0]; // Ambil nilai yang sudah ada

    // Buat objek dari nilai lama untuk memudahkan akses
    const oldRowData = {};
    headers.forEach((header, i) => {
      oldRowData[header] = oldValues[i];
    });

    // Tambahkan timestamp update
    formData['Update'] = new Date();

    // Buat baris baru dengan mempertahankan nilai lama jika tidak ada di form
    const newRowValues = headers.map(header => {
      // [KUNCI PERBAIKAN] Jika header ada di form, gunakan nilai baru.
      // Jika tidak (seperti 'Tanggal Input'), gunakan nilai lama.
      if (formData.hasOwnProperty(header)) {
        return formData[header];
      } else {
        return oldRowData[header];
      }
    });

    range.setValues([newRowValues]);

    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    Logger.log(`Error in updatePtkPaudData: ${e.message}`);
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

function getNewPtkPaudOptions() {
  try {
    // ID Spreadsheet baru untuk sumber data form
    const ss = SpreadsheetApp.openById("1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU");
    const sheet = ss.getSheetByName("Form PAUD");
    if (!sheet) throw new Error("Sheet 'Form PAUD' tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // Ambil kolom A sampai D

    const jenjangOptions = [...new Set(data.map(row => row[0]).filter(Boolean))].sort();
    const statusOptions = [...new Set(data.map(row => row[2]).filter(Boolean))].sort();
    const jabatanOptions = [...new Set(data.map(row => row[3]).filter(Boolean))].sort();

    // Buat pemetaan Nama Lembaga berdasarkan Jenjang
    const lembagaMap = {};
    data.forEach(row => {
      const jenjang = row[0];
      const lembaga = row[1];
      if (jenjang && lembaga) {
        if (!lembagaMap[jenjang]) {
          lembagaMap[jenjang] = [];
        }
        if (!lembagaMap[jenjang].includes(lembaga)) {
          lembagaMap[jenjang].push(lembaga);
        }
      }
    });
    // Urutkan setiap daftar lembaga
    for (const jenjang in lembagaMap) {
      lembagaMap[jenjang].sort();
    }

    return {
      'Jenjang': jenjangOptions,
      'Nama Lembaga': lembagaMap,
      'Status': statusOptions,
      'Jabatan': jabatanOptions
    };
  } catch (e) {
    Logger.log(`Error in getNewPtkPaudOptions: ${e.message}`);
    return { error: e.message };
  }
}

/**
 * Menambahkan data PTK PAUD baru ke spreadsheet.
 */
function addNewPtkPaud(formData) {
  try {
    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(header => {
      if (header === 'Tanggal Input') {
        return new Date();
      }
      return formData[header] || ""; // Gunakan string kosong jika tidak ada data
    });

    // appendRow akan selalu menambahkan data di baris paling bawah
    sheet.appendRow(newRow);
    
    // [PERBAIKAN] Dua baris kode untuk mengurutkan data telah dihapus dari sini

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    Logger.log(`Error in addNewPtkPaud: ${e.message}`);
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

/**
 * Menghapus data PTK PAUD dari spreadsheet.
 */
function deletePtkPaudData(rowIndex, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    const ss = SpreadsheetApp.openById("1iZO2VYIqKAn_ykJEzVAWtYS9dd23F_Y7TjeGN1nDSAk");
    const sheet = ss.getSheetByName("PTK PAUD");
    if (!sheet) throw new Error("Sheet 'PTK PAUD' tidak ditemukan.");
    
    // Validasi rowIndex untuk memastikan itu angka dan dalam rentang yang valid
    const maxRows = sheet.getLastRow();
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > maxRows) {
      throw new Error("Nomor baris tidak valid atau di luar jangkauan.");
    }

    sheet.deleteRow(rowIndex);
    
    return "Data PTK berhasil dihapus.";
  } catch (e) {
    Logger.log(`Error in deletePtkPaudData: ${e.message}`);
    throw new Error(`Gagal menghapus data: ${e.message}`);
  }
}

function getJumlahPtkSdBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("PTK Bulanan SD");
    if (!sheet) {
      throw new Error("Sheet 'PTK Bulanan SD' tidak ditemukan.");
    }
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getJumlahPtkSdBulananData', e);
  }
}

/**
 * Mengambil data daftar PTK SD Negeri dari spreadsheet.
 */
function getDaftarPtkSdnData() {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    const sheet = ss.getSheetByName("PTK SDN");
    if (!sheet) {
      throw new Error("Sheet 'PTK SDN' tidak ditemukan.");
    }
    // Mengirim seluruh data mentah, pemrosesan dilakukan di client-side
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdnData', e);
  }
}

/**
 * Mengambil data daftar PTK SD Swasta dari spreadsheet.
 */
function getDaftarPtkSdsData() {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    const sheet = ss.getSheetByName("PTK SDS");
    if (!sheet) {
      throw new Error("Sheet 'PTK SDS' tidak ditemukan.");
    }
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getDaftarPtkSdsData', e);
  }
}

/**
 * Mengambil data gabungan PTK SDN dan SDS untuk halaman Kelola.
 */
function getKelolaPtkSdData() {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    const sdnSheet = ss.getSheetByName("PTK SDN");
    const sdsSheet = ss.getSheetByName("PTK SDS");

    let combinedData = [];

    const processSheet = (sheet, sourceName) => {
      if (!sheet || sheet.getLastRow() < 2) return;
      const data = sheet.getDataRange().getDisplayValues(); // Menggunakan DisplayValues lebih aman
      const rows = data.slice(1);

      const idx = {
          unitKerja: 0, nama: 1, status: 3, nipNiy: 4, jabatan: 9, tglInput: 12, update: 13
      };

      rows.forEach((row, index) => {
        if (!row[idx.nama]) return; // Lewati baris jika nama kosong
        
        combinedData.push({
          rowIndex: index + 2,
          source: sourceName,
          data: {
            'Unit Kerja': row[idx.unitKerja],
            'Nama': row[idx.nama],
            'Status': row[idx.status],
            'NIP/NIY': row[idx.nipNiy],
            'Jabatan': row[idx.jabatan],
            'Tanggal Input': row[idx.tglInput],
            'Update': row[idx.update]
          }
        });
      });
    };

    processSheet(sdnSheet, 'SDN');
    processSheet(sdsSheet, 'SDS');
    
    // TIDAK ADA PENGURUTAN DI SINI, HANYA MENGIRIM DATA
    return { rows: combinedData };

  } catch (e) {
    Logger.log(`Error in getKelolaPtkSdData: ${e.message}`);
    return { error: e.message };
  }
}

/**
 * Mengambil opsi-opsi untuk form Tambah PTK SD.
 */
function getNewPtkSdOptions() {
  try {
    const dataSekolahSS = SpreadsheetApp.openById("1qeOYVfqFQdoTpysy55UIdKwAJv3VHo4df3g6u6m72Bs");
    const getValues = (sheetName) => {
      const sheet = dataSekolahSS.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return [];
      return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().flat().filter(Boolean).sort();
    };

    return {
      'Nama SDN': getValues('Nama SDN'),
      'Nama SDS': getValues('Nama SDS'),
      'Status Kepegawaian': getValues('Status Kepegawaian SD'),
      'Pangkat': getValues('Pangkat'),
      'Jabatan': getValues('Jabatan')
    };
  } catch (e) {
    return handleError('getNewPtkSdOptions', e);
  }
}

/**
 * Menambahkan data PTK SD baru ke spreadsheet yang sesuai.
 */
function addNewPtkSd(formData) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    
    if (formData.statusSekolah === 'Negeri') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (formData.statusSekolah === 'Swasta') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Status Sekolah tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet untuk status '${formData.statusSekolah}' tidak ditemukan.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    // Tambahkan timestamp
    formData['Tanggal Input'] = new Date();

    const newRow = headers.map(header => formData[header] || ""); // Gunakan string kosong jika tidak ada data

    sheet.appendRow(newRow);

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    Logger.log(`Error in addNewPtkSd: ${e.message}`);
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

/**
 * Mengambil semua opsi dropdown untuk form Tambah/Edit PTK SD.
 */
function getNewPtkSdOptions() {
  try {
    const ss = SpreadsheetApp.openById("1prqqKQBYzkCNFmuzblNAZE41ag9rZTCiY2a0WvZCTvU");
    const sheet = ss.getSheetByName("Form SD");
    if (!sheet) throw new Error("Sheet 'Form SD' tidak ditemukan.");

    const getUniqueValues = (col) => {
      const data = sheet.getRange(`${col}2:${col}${sheet.getLastRow()}`).getDisplayValues().flat().filter(Boolean);
      return [...new Set(data)].sort();
    };

    return {
      'Unit Kerja Negeri': getUniqueValues('B'),
      'Unit Kerja Swasta': getUniqueValues('A'),
      'Status Negeri': getUniqueValues('C'),
      'Pangkat PNS': getUniqueValues('D'),
      'Pangkat PPPK': getUniqueValues('E'),
      'Pangkat PPPK PW': getUniqueValues('F'),
      'Jabatan': getUniqueValues('G'),
      'Tugas Tambahan': getUniqueValues('H'),
      'Status Swasta': getUniqueValues('I'),
    };
  } catch (e) {
    return handleError('getNewPtkSdOptions', e);
  }
}

/**
 * Menambahkan data PTK SD baru ke sheet yang sesuai.
 */
function addNewPtkSd(formData) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    
    if (formData.statusSekolah === 'Negeri') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (formData.statusSekolah === 'Swasta') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Status Sekolah tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet untuk status '${formData.statusSekolah}' tidak ditemukan.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    formData['Input'] = new Date();

    const newRow = headers.map(header => {
      let value = formData[header] || "";
      // KUNCI PERBAIKAN 2: Jika header adalah NUPTK dan ada nilainya, tambahkan apostrof
      if (header === 'NUPTK' && value) {
        return "'" + value;
      }
      return value;
    });

    sheet.appendRow(newRow);

    return "Data PTK baru berhasil disimpan.";
  } catch (e) {
    Logger.log(`Error in addNewPtkSd: ${e.message}`);
    throw new Error(`Gagal menyimpan data: ${e.message}`);
  }
}

/**
 * Mengambil data satu baris PTK SD berdasarkan nomor baris dan sumbernya.
 */
function getPtkSdDataByRow(rowIndex, source) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid.");
    }
    if (!sheet) throw new Error(`Sheet '${source}' tidak ditemukan.`);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const displayValues = sheet.getRange(rowIndex, 1, 1, headers.length).getDisplayValues()[0];
    
    const rowData = {};
    headers.forEach((header, i) => {
      // Untuk tanggal, ambil nilai mentahnya agar bisa di-format di form
      if ((header === 'TMT' || header === 'TMT CPNS' || header === 'TMT PNS') && values[i] instanceof Date) {
        rowData[header] = Utilities.formatDate(values[i], "UTC", "yyyy-MM-dd");
      } else {
        rowData[header] = displayValues[i];
      }
    });
    return rowData;
  } catch (e) {
    return handleError('getPtkSdDataByRow', e);
  }
}

/**
 * Memperbarui data PTK SD yang ada di spreadsheet.
 */
function updatePtkSdData(formData) {
  try {
    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    const source = formData.source;
    
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid.");
    }

    if (!sheet) throw new Error(`Sheet '${source}' tidak ditemukan.`);
    const rowIndex = formData.rowIndex;
    if (!rowIndex) throw new Error("Nomor baris (rowIndex) tidak ditemukan.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    const oldValues = range.getValues()[0];

    // Tambahkan timestamp update
    formData['Update'] = new Date();
    
    // NUPTK sebagai teks
    if (formData['NUPTK']) {
      formData['NUPTK'] = "'" + formData['NUPTK'];
    }

    const newRowValues = headers.map((header, index) => {
      // Jika header ada di form, gunakan nilai baru. Jika tidak, gunakan nilai lama.
      return formData.hasOwnProperty(header) ? formData[header] : oldValues[index];
    });

    range.setValues([newRowValues]);

    return "Data PTK berhasil diperbarui.";
  } catch (e) {
    Logger.log(`Error in updatePtkSdData: ${e.message}`);
    throw new Error(`Gagal memperbarui data: ${e.message}`);
  }
}

/**
 * Mengambil data Kebutuhan PTK SD Negeri.
 */
function getKebutuhanPtkSdnData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("Kebutuhan Guru");
    if (!sheet) {
      throw new Error("Sheet 'Kebutuhan Guru' tidak ditemukan.");
    }
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getKebutuhanPtkSdnData', e);
  }
}

/**
 * Menghapus data PTK SD (Negeri atau Swasta) dari spreadsheet.
 */
function deletePtkSdData(rowIndex, source, deleteCode) {
  try {
    const today = new Date();
    const todayCode = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    
    if (String(deleteCode).trim() !== todayCode) {
      throw new Error("Kode Hapus salah.");
    }

    const ss = SpreadsheetApp.openById("1HlyLv3Ai3_vKFJu3EKznqI9v8g0tfqiNg0UbIojNMQ0");
    let sheet;
    if (source === 'SDN') {
      sheet = ss.getSheetByName("PTK SDN");
    } else if (source === 'SDS') {
      sheet = ss.getSheetByName("PTK SDS");
    } else {
      throw new Error("Sumber data tidak valid: " + source);
    }

    if (!sheet) throw new Error("Sheet sumber '" + source + "' tidak ditemukan.");
    
    const maxRows = sheet.getLastRow();
    if (isNaN(rowIndex) || rowIndex < 2 || rowIndex > maxRows) {
      throw new Error("Nomor baris tidak valid atau di luar jangkauan.");
    }

    sheet.deleteRow(rowIndex);
    
    return "Data PTK berhasil dihapus.";
  } catch (e) {
    Logger.log(`Error in deletePtkSdData: ${e.message}`);
    throw new Error(`Gagal menghapus data: ${e.message}`);
  }
}

/**
 * Mengambil data keadaan murid PAUD menurut kelas dan usia.
 */
function getMuridPaudKelasUsiaData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Murid Kelas");
    
    if (!sheet) {
      throw new Error("Sheet 'Murid Kelas' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridPaudKelasUsiaData', e);
  }
}

/**
 * Mengambil data keadaan murid PAUD menurut jenis kelamin.
 */
function getMuridPaudJenisKelaminData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Murid JK");
    
    if (!sheet) {
      throw new Error("Sheet 'Murid JK' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridPaudJenisKelaminData', e);
  }
}

/**
 * Mengambil data jumlah murid PAUD bulanan.
 */
function getMuridPaudJumlahBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1an0oQQPdMh6wrUJIAzTGYk3DKFvYprK5SU7RmRXjIgs");
    const sheet = ss.getSheetByName("Murid Bulanan");
    
    if (!sheet) {
      throw new Error("Sheet 'Murid Bulanan' tidak ditemukan.");
    }
    
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridPaudJumlahBulananData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut kelas.
 */
function getMuridSdKelasData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Kelas");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Kelas' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdKelasData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut rombel.
 */
function getMuridSdRombelData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Rombel");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Rombel' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdRombelData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut jenis kelamin.
 */
function getMuridSdJenisKelaminData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel JK");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel JK' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdJenisKelaminData', e);
  }
}

/**
 * Mengambil data keadaan murid SD menurut agama.
 */
function getMuridSdAgamaData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Agama");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Agama' tidak ditemukan.");
    }
    
    // Mengambil semua data dari sheet
    return sheet.getDataRange().getDisplayValues();
  } catch (e) {
    return handleError('getMuridSdAgamaData', e);
  }
}

/**
 * Mengambil data jumlah murid SD bulanan.
 */
function getMuridSdJumlahBulananData() {
  try {
    const ss = SpreadsheetApp.openById("1u4tNL3uqt5xHITXYwHnytK6Kul9Siam-vNYuzmdZB4s");
    const sheet = ss.getSheetByName("SD Tabel Bulanan");
    
    if (!sheet) {
      throw new Error("Sheet 'SD Tabel Bulanan' tidak ditemukan.");
    }
    
    // PERBAIKAN: Menggunakan getValues() untuk mendapatkan angka murni, bukan teks.
    return sheet.getDataRange().getValues();
  } catch (e) {
    return handleError('getMuridSdJumlahBulananData', e);
  }
}

/**
 * ===================================================================
 * ======================= 7. MODUL SIABA ============================
 * ===================================================================
 */

/**
 * Mengambil opsi filter (Tahun, Bulan, Unit Kerja) untuk Daftar Presensi SIABA.
 */
function getSiabaFilterOptions() {
  try {
    // Mengambil data Unit Kerja dari sheet "Unit Siaba"
    const ssDropdown = SpreadsheetApp.openById("1wiDKez4rL5UYnpP2-OZjYowvmt1nRx-fIMy9trJlhBA");
    const sheetUnitKerja = ssDropdown.getSheetByName("Unit Siaba");
    let unitKerjaOptions = [];
    if (sheetUnitKerja && sheetUnitKerja.getLastRow() > 1) {
      unitKerjaOptions = sheetUnitKerja.getRange(2, 1, sheetUnitKerja.getLastRow() - 1, 1)
                                      .getDisplayValues()
                                      .flat()
                                      .filter(Boolean)
                                      .sort();
    }

    // Mengambil data Tahun dan Bulan dari Spreadsheet SIABA
    const ssSiaba = SpreadsheetApp.openById(SPREADSHEET_CONFIG.SIABA_REKAP.id);
    const sheetSiaba = ssSiaba.getSheetByName(SPREADSHEET_CONFIG.SIABA_REKAP.sheet);
    if (!sheetSiaba || sheetSiaba.getLastRow() < 2) {
         throw new Error("Sheet Rekap SIABA tidak ditemukan atau kosong.");
    }

    const tahunBulanData = sheetSiaba.getRange(2, 1, sheetSiaba.getLastRow() - 1, 2).getDisplayValues();
    const uniqueTahun = [...new Set(tahunBulanData.map(row => row[0]))].filter(Boolean).sort().reverse();
    const uniqueBulan = [...new Set(tahunBulanData.map(row => row[1]))].filter(Boolean);
    
    const monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    uniqueBulan.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    return {
      'Tahun': uniqueTahun,
      'Bulan': uniqueBulan,
      'Unit Kerja': unitKerjaOptions
    };
  } catch (e) {
    return handleError('getSiabaFilterOptions', e);
  }
}


/**
 * Mengambil data presensi SIABA berdasarkan filter TAHUN dan BULAN.
 * FINAL: Mengurutkan berdasarkan 6 level prioritas.
 * Prioritas: TP -> TA -> PLA -> TAp -> TU -> Nama.
 */
function getSiabaPresensiData(filters) {
  try {
    const { tahun, bulan } = filters;
    if (!tahun || !bulan) {
      throw new Error("Filter Tahun dan Bulan wajib diisi.");
    }

    const config = SPREADSHEET_CONFIG.SIABA_REKAP;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
        return { headers: [], rows: [] };
    }

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => {
      const tahunMatch = String(row[0]) === String(tahun); // Kolom A
      const bulanMatch = String(row[1]) === String(bulan); // Kolom B
      return tahunMatch && bulanMatch;
    });

    const startIndex = 2;
    const endIndex = 86;

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => {
       return row.slice(startIndex, endIndex + 1);
    });
    
    // --- BLOK PENGURUTAN MULTI-LEVEL (FINAL) ---
    
    const tpIndex = displayHeaders.indexOf('TP');
    const taIndex = displayHeaders.indexOf('TA');
    const plaIndex = displayHeaders.indexOf('PLA');
    const tapIndex = displayHeaders.indexOf('TAp');
    const tuIndex = displayHeaders.indexOf('TU'); // <-- Tambahkan Prioritas 5
    const namaIndex = displayHeaders.indexOf('Nama');

    // Lakukan pengurutan hanya jika kolom prioritas utama (TP) ditemukan
    if (tpIndex !== -1) {
      displayRows.sort((a, b) => {
        const compareDesc = (index) => {
            if (index === -1) return 0; // Jika kolom tidak ditemukan, jangan urutkan
            const valB = parseInt(b[index], 10) || 0;
            const valA = parseInt(a[index], 10) || 0;
            return valB - valA;
        };
        
        // Prioritas 1: Urutkan berdasarkan TP
        let diff = compareDesc(tpIndex);
        if (diff !== 0) return diff;

        // Prioritas 2: Jika TP sama, urutkan berdasarkan TA
        diff = compareDesc(taIndex);
        if (diff !== 0) return diff;

        // Prioritas 3: Jika TA sama, urutkan berdasarkan PLA
        diff = compareDesc(plaIndex);
        if (diff !== 0) return diff;

        // Prioritas 4: Jika PLA sama, urutkan berdasarkan TAp
        diff = compareDesc(tapIndex);
        if (diff !== 0) return diff;
        
        // Prioritas 5: Jika TAp sama, urutkan berdasarkan TU
        diff = compareDesc(tuIndex);
        if (diff !== 0) return diff;

        // Prioritas 6 (Fallback): Jika semua sama, urutkan berdasarkan Nama
        if (namaIndex !== -1) {
            const namaA = a[namaIndex] || "";
            const namaB = b[namaIndex] || "";
            return namaA.localeCompare(namaB);
        }

        return 0;
      });
    }
    // --- AKHIR BLOK PENGURUTAN ---

    return { headers: displayHeaders, rows: displayRows };

  } catch (e) {
    return handleError('getSiabaPresensiData', e);
  }
}

/**
 * Mengambil opsi filter (Tahun, Bulan, Unit Kerja) dari sheet "ASN Tidak Presensi".
 */
function getSiabaTidakPresensiFilterOptions() {
  try {
    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
         throw new Error("Sheet 'Rekap Script' untuk data Tidak Presensi tidak ditemukan atau kosong.");
    }

    const filterData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues(); // Kolom A, B, C
    const uniqueTahun = [...new Set(filterData.map(row => row[0]))].filter(Boolean).sort().reverse();
    const uniqueBulan = [...new Set(filterData.map(row => row[1]))].filter(Boolean);
    const uniqueUnitKerja = [...new Set(filterData.map(row => row[2]))].filter(Boolean).sort();
    
    const monthOrder = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    uniqueBulan.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));

    return {
      'Tahun': uniqueTahun,
      'Bulan': uniqueBulan,
      'Unit Kerja': uniqueUnitKerja
    };
  } catch (e) {
    return handleError('getSiabaTidakPresensiFilterOptions', e);
  }
}

/**
 * Mengambil data ASN Tidak Presensi berdasarkan filter.
 * Diurutkan berdasarkan 'Jumlah' (terbanyak) lalu 'Nama' (abjad).
 */
function getSiabaTidakPresensiData(filters) {
  try {
    const { tahun, bulan, unitKerja } = filters;
    if (!tahun || !bulan) {
      throw new Error("Filter Tahun dan Bulan wajib diisi.");
    }

    const config = SPREADSHEET_CONFIG.SIABA_TIDAK_PRESENSI;
    const sheet = SpreadsheetApp.openById(config.id).getSheetByName(config.sheet);
    if (!sheet || sheet.getLastRow() < 2) {
        return { headers: [], rows: [] };
    }

    const allData = sheet.getDataRange().getDisplayValues();
    const headers = allData[0];
    const dataRows = allData.slice(1);

    const filteredRows = dataRows.filter(row => {
      const tahunMatch = String(row[0]) === String(tahun);
      const bulanMatch = String(row[1]) === String(bulan);
      const unitKerjaMatch = (unitKerja === "Semua") || (String(row[2]) === String(unitKerja));
      return tahunMatch && bulanMatch && unitKerjaMatch;
    });

    const startIndex = 3; // Kolom D
    const endIndex = 8;   // Kolom I

    const displayHeaders = headers.slice(startIndex, endIndex + 1);
    const displayRows = filteredRows.map(row => {
       return row.slice(startIndex, endIndex + 1);
    });
    
    // --- BLOK PENGURUTAN BARU ---
    // Cari posisi (indeks) kolom 'Jumlah' dan 'Nama' secara dinamis
    const jumlahIndex = displayHeaders.indexOf('Jumlah'); // Seharusnya di posisi ke-4
    const namaIndex = displayHeaders.indexOf('Nama');     // Seharusnya di posisi ke-0

    // Lakukan pengurutan
    displayRows.sort((a, b) => {
      // Prioritas 1: Urutkan berdasarkan Jumlah (menurun)
      const valB_jumlah = (jumlahIndex !== -1) ? (parseInt(b[jumlahIndex], 10) || 0) : 0;
      const valA_jumlah = (jumlahIndex !== -1) ? (parseInt(a[jumlahIndex], 10) || 0) : 0;
      if (valB_jumlah !== valA_jumlah) {
        return valB_jumlah - valA_jumlah;
      }

      // Prioritas 2 (Fallback): Jika Jumlah sama, urutkan berdasarkan Nama (abjad A-Z)
      if (namaIndex !== -1) {
          const namaA = a[namaIndex] || "";
          const namaB = b[namaIndex] || "";
          return namaA.localeCompare(namaB);
      }
      return 0;
    });
    // --- AKHIR BLOK PENGURUTAN ---

    return { headers: displayHeaders, rows: displayRows };
  } catch (e) {
    return handleError('getSiabaTidakPresensiData', e);
  }
}