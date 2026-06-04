/**
 * KuyDesain-Manager Backend - Google Apps Script
 * Mengelola interaksi antara Google Spreadsheet sebagai Database dan Frontend Web App.
 * Menyediakan setup database otomatis dan fungsi CRUD.
 */

// Global App ID untuk melacak instance
const APP_TITLE = "KuyDesain-Manager";

/**
 * Menghandle request GET untuk menampilkan halaman web utama
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Setup Database: Membuat sheet yang dibutuhkan jika belum ada
 * Dipanggil otomatis saat aplikasi mendeteksi sheet kosong atau bisa dipanggil manual.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Daftar sheet yang dibutuhkan beserta header kolomnya
  const sheetsConfig = {
    "Pemasukan": ["ID", "Tanggal", "Nama Pelanggan", "Jenis Transaksi", "Nominal", "Catatan"],
    "Pengeluaran": ["ID", "Tanggal", "Jenis Pengeluaran", "Nominal", "Catatan"],
    "Stok_Barang": ["ID", "Nama Barang", "Jenis Barang", "Supplier", "Satuan", "Limit Minimum", "Harga Beli Awal", "Stok Sekarang"],
    "Barang_Masuk": ["ID", "Tanggal Masuk", "Nama Barang", "Supplier", "Jumlah Masuk", "Harga Satuan"],
    "Barang_Keluar": ["ID", "Tanggal", "Nama Barang", "Jumlah", "Keperluan"],
    "Penjualan": ["ID", "Tanggal", "Nama Pelanggan", "Jenis Layanan", "Qty", "Harga Satuan", "Total"],
    "Hutang": ["ID", "Tanggal", "Nama Supplier", "Nominal", "Status"],
    "Piutang": ["ID", "Tanggal", "Nama Pelanggan", "Nominal", "Status"]
  };

  for (let sheetName in sheetsConfig) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheetsConfig[sheetName]);
      // Format header menjadi bold
      sheet.getRange(1, 1, 1, sheetsConfig[sheetName].length)
           .setFontWeight("bold")
           .setBackground("#1e3a8a") // Biru tua
           .setFontColor("#ffffff");
    }
  }
  return "Database berhasil diverifikasi dan disiapkan!";
}

/**
 * Membantu mendapatkan sheet atau melakukan auto-setup jika sheet tidak ditemukan
 */
function getSheetSafe(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

/**
 * Generate ID Unik acak
 */
function generateId() {
  return "KD-" + Math.floor(100000 + Math.random() * 900000);
}

/**
 * Mengambil semua data gabungan untuk Dashboard & Laporan Keuangan
 */
function getAllData() {
  try {
    const data = {
      pemasukan: getSheetData("Pemasukan"),
      pengeluaran: getSheetData("Pengeluaran"),
      stok: getSheetData("Stok_Barang"),
      barangMasuk: getSheetData("Barang_Masuk"),
      barangKeluar: getSheetData("Barang_Keluar"),
      penjualan: getSheetData("Penjualan"),
      hutang: getSheetData("Hutang"),
      piutang: getSheetData("Piutang")
    };
    return { success: true, data: data };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Helper untuk mengambil data dari sheet dan mengubahnya menjadi array of objects
 */
function getSheetData(sheetName) {
  const sheet = getSheetSafe(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  return values.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      let val = row[i];
      // Format tanggal agar aman dikirim ke JSON
      if (val instanceof Date) {
        val = val.toISOString().split('T')[0];
      }
      obj[header.replace(/\s+/g, '')] = val; // menghilangkan spasi di key objek
    });
    return obj;
  });
}

/**
 * 1. Simpan Pemasukan Baru
 */
function savePemasukan(data) {
  try {
    const sheet = getSheetSafe("Pemasukan");
    const id = generateId();
    sheet.appendRow([
      id,
      new Date(data.tanggal),
      data.namaPelanggan,
      data.jenisTransaksi,
      parseFloat(data.nominal),
      data.catatan || ""
    ]);
    return { success: true, message: "Pemasukan berhasil disimpan" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 2. Simpan Pengeluaran Baru
 */
function savePengeluaran(data) {
  try {
    const sheet = getSheetSafe("Pengeluaran");
    const id = generateId();
    sheet.appendRow([
      id,
      new Date(data.tanggal),
      data.jenisPengeluaran,
      parseFloat(data.nominal),
      data.catatan || ""
    ]);
    return { success: true, message: "Pengeluaran berhasil disimpan" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 3. Tambah Barang Baru (Stok Awal)
 */
function saveBarangBaru(data) {
  try {
    const sheet = getSheetSafe("Stok_Barang");
    const id = generateId();
    
    // Validasi apakah barang dengan nama yang sama sudah ada
    const existing = getSheetData("Stok_Barang");
    const isDuplicate = existing.some(b => b.NamaBarang.toLowerCase() === data.namaBarang.toLowerCase());
    if (isDuplicate) {
      return { success: false, error: "Barang dengan nama tersebut sudah terdaftar!" };
    }

    sheet.appendRow([
      id,
      data.namaBarang,
      data.jenisBarang,
      data.supplier,
      data.satuan,
      parseInt(data.limitMinimum),
      parseFloat(data.hargaBeliAwal),
      0 // Stok Sekarang dimulai dari 0, bertambah jika ada barang masuk
    ]);
    return { success: true, message: "Barang baru berhasil ditambahkan" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 4. Catat Barang Masuk (Menambah Stok)
 * Jika barang belum pernah didaftarkan di modul Stok_Barang,
 * sistem akan mendaftarkannya secara otomatis dengan parameter default.
 */
function saveBarangMasuk(data) {
  try {
    const id = generateId();
    const sheetMasuk = getSheetSafe("Barang_Masuk");
    const sheetStok = getSheetSafe("Stok_Barang");
    
    // Tambah log barang masuk
    sheetMasuk.appendRow([
      id,
      new Date(data.tanggalMasuk),
      data.namaBarang,
      data.supplier,
      parseInt(data.jumlahMasuk),
      parseFloat(data.hargaSatuan)
    ]);
    
    // Update stok di sheet Stok_Barang
    const rows = sheetStok.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][1].toString().toLowerCase() === data.namaBarang.toString().toLowerCase()) { // Match case-insensitive
        const currentStok = parseInt(rows[i][7] || 0);
        sheetStok.getRange(i + 1, 8).setValue(currentStok + parseInt(data.jumlahMasuk));
        updated = true;
        break;
      }
    }
    
    // JIKA BARANG BARU (Belum terdaftar di stok)
    // Daftarkan barang baru tersebut secara otomatis agar konsistensi tabel persediaan terjaga
    if (!updated) {
      const newStokId = generateId();
      sheetStok.appendRow([
        newStokId,
        data.namaBarang,
        "Barang Masuk", // Jenis / Kategori default untuk penginputan langsung
        data.supplier,
        "Pcs", // Satuan default
        5, // Limit minimum default
        parseFloat(data.hargaSatuan), // Harga beli dari transaksi saat ini
        parseInt(data.jumlahMasuk) // Stok awal diisi langsung dari jumlah masuk
      ]);
    }
    
    return { 
      success: true, 
      message: updated 
        ? "Barang masuk berhasil dicatat & stok diperbarui!" 
        : "Barang baru didaftarkan secara otomatis ke dalam data Stok Barang!" 
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 5. Kurangi Stok Otomatis Berdasarkan Jenis Layanan Penjualan (Rules)
 */
function kurangiStokOtomatis(jenisLayanan, qty) {
  const sheetStok = getSheetSafe("Stok_Barang");
  const sheetKeluar = getSheetSafe("Barang_Keluar");
  const rows = sheetStok.getDataRange().getValues();
  
  let targetBarang = "";
  let faktorPengurangan = 1; // 1 item stok per Qty layanan default
  
  // Rule Sistem Otomatis Berdasarkan PRD
  if (jenisLayanan === "Fotocopy") {
    targetBarang = "Kertas HVS";
    faktorPengurangan = 1;
  } else if (jenisLayanan === "Cetak Foto") {
    targetBarang = "Kertas Foto";
    faktorPengurangan = 1;
  } else if (jenisLayanan === "Cetak Stiker") {
    targetBarang = "Stiker";
    faktorPengurangan = 1;
  } else if (jenisLayanan === "Print") {
    // Print mengurangi kertas dan tinta (kita buat simulasi mengurangi Kertas HVS)
    targetBarang = "Kertas HVS";
    faktorPengurangan = 1;
  }
  
  if (targetBarang !== "") {
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][1].toLowerCase().includes(targetBarang.toLowerCase())) {
        const currentStok = parseInt(rows[i][7] || 0);
        const jumlahKurang = qty * faktorPengurangan;
        sheetStok.getRange(i + 1, 8).setValue(Math.max(0, currentStok - jumlahKurang));
        
        // Catat riwayat barang keluar
        sheetKeluar.appendRow([
          generateId(),
          new Date(),
          rows[i][1], // Nama barang asli di database
          jumlahKurang,
          `Penjualan Layanan: ${jenisLayanan}`
        ]);
        
        updated = true;
        break;
      }
    }
  }
}

/**
 * 6. Simpan Penjualan / Kasir Baru
 */
function savePenjualan(data) {
  try {
    const sheet = getSheetSafe("Penjualan");
    const id = generateId();
    const total = parseFloat(data.qty) * parseFloat(data.hargaSatuan);
    const tgl = new Date(data.tanggal);
    
    // Simpan ke penjualan
    sheet.appendRow([
      id,
      tgl,
      data.namaPelanggan,
      data.jenisLayanan,
      parseInt(data.qty),
      parseFloat(data.hargaSatuan),
      total
    ]);
    
    // Otomatis kurangi stok sesuai aturan di PRD
    kurangiStokOtomatis(data.jenisLayanan, parseInt(data.qty));
    
    // Otomatis masukkan ke Pemasukan
    const sheetPemasukan = getSheetSafe("Pemasukan");
    sheetPemasukan.appendRow([
      generateId(),
      tgl,
      data.namaPelanggan,
      data.jenisLayanan,
      total,
      `Kasir: Transaksi Penjualan ID ${id}`
    ]);
    
    return { success: true, message: "Transaksi Berhasil dan Nota siap dibuat!", idPenjualan: id, total: total };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 7. Simpan Hutang Supplier Baru
 */
function saveHutang(data) {
  try {
    const sheet = getSheetSafe("Hutang");
    const id = generateId();
    sheet.appendRow([
      id,
      new Date(data.tanggal),
      data.namaSupplier,
      parseFloat(data.nominal),
      data.status || "Belum Lunas"
    ]);
    return { success: true, message: "Catatan hutang berhasil disimpan" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 8. Simpan Piutang Pelanggan Baru
 */
function savePiutang(data) {
  try {
    const sheet = getSheetSafe("Piutang");
    const id = generateId();
    sheet.appendRow([
      id,
      new Date(data.tanggal),
      data.namaPelanggan,
      parseFloat(data.nominal),
      data.status || "Belum Bayar"
    ]);
    return { success: true, message: "Catatan piutang berhasil disimpan" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Update Status Hutang/Piutang (Lunas / Belum Lunas)
 */
function updateStatus(tipe, id, statusBaru) {
  try {
    const sheetName = tipe === "hutang" ? "Hutang" : "Piutang";
    const sheet = getSheetSafe(sheetName);
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.getRange(i + 1, 5).setValue(statusBaru);
        return { success: true, message: `Status ${tipe} berhasil diperbarui menjadi ${statusBaru}` };
      }
    }
    return { success: false, error: "ID tidak ditemukan" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Export Laporan Excel Berwarna & Profesional via Google Drive Link / Download Link
 * Membuat spreadsheet baru sementara di Drive, memformatnya, lalu memberikan URL download.
 */
function generateProfessionalExcel() {
  try {
    const originalSs = SpreadsheetApp.getActiveSpreadsheet();
    const tempFileName = "Laporan_KuyDesain_Manager_" + new Date().toISOString().split('T')[0];
    
    // Buat spreadsheet baru di Drive
    const tempSs = SpreadsheetApp.create(tempFileName);
    const tempSheets = tempSs.getSheets();
    
    // Copy sheet dan struktur data dari database asli
    const sourceSheetNames = ["Pemasukan", "Pengeluaran", "Stok_Barang", "Barang_Masuk", "Barang_Keluar", "Penjualan", "Hutang", "Piutang"];
    
    sourceSheetNames.forEach((name, idx) => {
      const srcSheet = originalSs.getSheetByName(name);
      if (srcSheet) {
        let destSheet = tempSs.getSheetByName(name);
        if (!destSheet) {
          destSheet = tempSs.insertSheet(name);
        }
        
        // Copy seluruh data & format dasar
        srcSheet.getDataRange().copyTo(destSheet.getRange(1, 1));
        
        // Mempercantik Tampilan Table Excel secara Profesional
        const lastRow = destSheet.getLastRow();
        const lastCol = destSheet.getLastColumn();
        
        if (lastRow > 0 && lastCol > 0) {
          // 1. Header Row (Biru Tua sesuai Logo, Teks Putih, Bold)
          const headerRange = destSheet.getRange(1, 1, 1, lastCol);
          headerRange.setBackground("#1e3a8a") 
                     .setFontColor("#ffffff")
                     .setFontWeight("bold")
                     .setHorizontalAlignment("center");
                     
          // 2. Set Font Family ke Arial / Roboto agar modern
          const bodyRange = destSheet.getRange(1, 1, lastRow, lastCol);
          bodyRange.setFontFamily("Arial")
                   .setFontSize(10);
                   
          // 3. Gridlines & Border
          bodyRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
          
          // 4. Auto-fit column widths
          for (let col = 1; col <= lastCol; col++) {
            destSheet.autoResizeColumn(col);
          }
          
          // 5. Format nominal rupiah di kolom yang relevan
          formatCurrencyColumns(destSheet, name, lastRow);
        }
      }
    });
    
    // Hapus "Sheet1" bawaan dari file temp
    const defaultSheet = tempSs.getSheetByName("Sheet1");
    if (defaultSheet) tempSs.deleteSheet(defaultSheet);
    
    // Simpan perubahan
    SpreadsheetApp.flush();
    
    // Dapatkan URL Download langsung sebagai XLSX
    const fileId = tempSs.getId();
    const file = DriveApp.getFileById(fileId);
    
    // Set file agar bisa diakses sementara oleh siapapun yang memiliki link (agar bisa didownload langsung)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const downloadUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
    
    return { success: true, downloadUrl: downloadUrl, fileId: fileId };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Format kolom berisi harga menjadi mata uang Rupiah
 */
function formatCurrencyColumns(sheet, name, lastRow) {
  if (lastRow <= 1) return;
  // Menentukan indeks kolom nominal di tiap sheet (1-based index)
  let colsToFormat = [];
  if (name === "Pemasukan") colsToFormat = [5]; // Nominal
  else if (name === "Pengeluaran") colsToFormat = [4]; // Nominal
  else if (name === "Stok_Barang") colsToFormat = [7]; // Harga Beli Awal
  else if (name === "Barang_Masuk") colsToFormat = [6]; // Harga Satuan
  else if (name === "Penjualan") colsToFormat = [6, 7]; // Harga Satuan, Total
  else if (name === "Hutang") colsToFormat = [4]; // Nominal
  else if (name === "Piutang") colsToFormat = [4]; // Nominal
  
  colsToFormat.forEach(col => {
    sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('_-* Rp" " #,##0_-;-(Rp" " #,##0);_-* "Rp -"_-;_-@_-');
  });
}
