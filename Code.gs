/**
 * KuyDesain-Manager Backend
 * Menggunakan Google Spreadsheet sebagai Database utama
 */

function doGet() {
  // Mencari file HTML dengan nama "Index"
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('KuyDesain-Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Setup Database otomatis di Google Spreadsheet aktif
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetsDef = {
    "Pemasukan": ["Tanggal", "Nama Pelanggan", "Jenis Transaksi", "Nominal", "Catatan"],
    "Pengeluaran": ["Tanggal", "Jenis Pengeluaran", "Nominal", "Catatan"],
    "Stok": ["ID Barang", "Nama Barang", "Jenis Barang", "Supplier", "Satuan", "Stok", "Limit Minimum", "Harga Beli Awal"],
    "BarangMasuk": ["Tanggal", "ID Barang", "Nama Barang", "Supplier", "Jumlah Masuk", "Harga Satuan"],
    "Penjualan": ["Tanggal", "ID Transaksi", "Nama Pelanggan", "Jenis Layanan", "Qty", "Harga Satuan", "Total"],
    "Hutang": ["ID Hutang", "Tanggal", "Nama Supplier", "Nominal", "Status"],
    "Piutang": ["ID Piutang", "Tanggal", "Nama Pelanggan", "Nominal", "Status"]
  };

  for (let sheetName in sheetsDef) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheetsDef[sheetName]);
      
      // Styling header kolom spreadsheet
      const headerRange = sheet.getRange(1, 1, 1, sheetsDef[sheetName].length);
      headerRange.setBackground("#120C18")
                 .setFontColor("#FFFFFF")
                 .setFontWeight("bold")
                 .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }
  }
  
  // Mengisi stok barang awal bawaan jika kosong
  const stokSheet = ss.getSheetByName("Stok");
  if (stokSheet.getLastRow() === 1) {
    const defaultItems = [
      ["BRG-001", "Kertas HVS A4", "Kertas", "Sinar Dunia", "Rim", 100, 10, 45000],
      ["BRG-002", "Kertas Foto Glossy", "Kertas", "Spectra", "Pack", 50, 5, 35000],
      ["BRG-003", "Kertas Stiker", "Kertas", "Glossy Paper", "Pack", 40, 5, 40000],
      ["BRG-004", "Tinta Hitam", "Tinta", "Epson Ink", "Botol", 10, 2, 85000],
      ["BRG-005", "Tinta Warna", "Tinta", "Epson Ink", "Botol", 10, 2, 95000]
    ];
    defaultItems.forEach(row => stokSheet.appendRow(row));
  }

  return {
    success: true,
    message: "Database KuyDesain-Manager berhasil disiapkan!",
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * Mengambil seluruh data tabel dari Google Spreadsheet
 */
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = {};
  const sheets = ["Pemasukan", "Pengeluaran", "Stok", "BarangMasuk", "Penjualan", "Hutang", "Piutang"];
  
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        data[sheetName] = values.map(row => {
          let obj = {};
          headers.forEach((header, index) => {
            obj[header.replace(/\s+/g, '')] = row[index];
          });
          return obj;
        });
      } else {
        data[sheetName] = [];
      }
    } else {
      data[sheetName] = [];
    }
  });
  
  return data;
}

function addPemasukan(tanggal, namaPelanggan, jenisTransaksi, nominal, catatan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pemasukan");
  sheet.appendRow([tanggal, namaPelanggan, jenisTransaksi, nominal, catatan]);
  return { success: true };
}

function addPengeluaran(tanggal, jenisPengeluaran, nominal, catatan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pengeluaran");
  sheet.appendRow([tanggal, jenisPengeluaran, nominal, catatan]);
  return { success: true };
}

function addStokBarang(idBarang, namaBarang, jenisBarang, supplier, satuan, limitMin, hargaBeli) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Stok");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idBarang) {
      return { success: false, message: "ID Barang sudah terdaftar!" };
    }
  }
  sheet.appendRow([idBarang, namaBarang, jenisBarang, supplier, satuan, 0, limitMin, hargaBeli]);
  return { success: true };
}

function addBarangMasuk(tanggal, idBarang, namaBarang, supplier, jumlahMasuk, hargaSatuan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masukSheet = ss.getSheetByName("BarangMasuk");
  masukSheet.appendRow([tanggal, idBarang, namaBarang, supplier, jumlahMasuk, hargaSatuan]);
  
  const stokSheet = ss.getSheetByName("Stok");
  const data = stokSheet.getDataRange().getValues();
  let updated = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idBarang) {
      const currentStok = Number(data[i][5]);
      stokSheet.getRange(i + 1, 6).setValue(currentStok + Number(jumlahMasuk));
      updated = true;
      break;
    }
  }
  return { success: true, updatedStok: updated };
}

function savePenjualan(tanggal, idTransaksi, namaPelanggan, jenisLayanan, qty, hargaSatuan, total) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Catat ke penjualan
  const penjualanSheet = ss.getSheetByName("Penjualan");
  penjualanSheet.appendRow([tanggal, idTransaksi, namaPelanggan, jenisLayanan, qty, hargaSatuan, total]);
  
  // 2. Integrasikan otomatis ke kas masuk (Pemasukan)
  const pemasukanSheet = ss.getSheetByName("Pemasukan");
  pemasukanSheet.appendRow([tanggal, namaPelanggan, jenisLayanan, total, "Transaksi Kasir ID: " + idTransaksi]);
  
  // 3. Pengurangan Stok Otomatis berdasarkan aturan bisnis PRD
  const stokSheet = ss.getSheetByName("Stok");
  const stokData = stokSheet.getDataRange().getValues();
  const qtyDeduct = Number(qty);
  
  const deductStock = (itemName, count) => {
    for (let i = 1; i < stokData.length; i++) {
      if (stokData[i][1].toLowerCase().indexOf(itemName.toLowerCase()) !== -1) {
        const current = Number(stokData[i][5]);
        stokSheet.getRange(i + 1, 6).setValue(Math.max(0, current - count));
        break;
      }
    }
  };

  if (jenisLayanan === "Fotocopy") {
    deductStock("Kertas HVS", qtyDeduct);
  } else if (jenisLayanan === "Cetak Foto") {
    deductStock("Kertas Foto", qtyDeduct);
  } else if (jenisLayanan === "Cetak Stiker") {
    deductStock("Kertas Stiker", qtyDeduct);
  } else if (jenisLayanan === "Print") {
    deductStock("Kertas HVS", qtyDeduct);
    deductStock("Tinta Hitam", Math.ceil(qtyDeduct * 0.05));
  }
  
  return { success: true };
}

function addHutang(idHutang, tanggal, namaSupplier, nominal, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Hutang");
  sheet.appendRow([idHutang, tanggal, namaSupplier, nominal, status]);
  return { success: true };
}

function addPiutang(idPiutang, tanggal, namaPelanggan, nominal, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Piutang");
  sheet.appendRow([idPiutang, tanggal, namaPelanggan, nominal, status]);
  return { success: true };
}

function toggleHutang(idHutang, statusBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Hutang");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idHutang) {
      sheet.getRange(i + 1, 5).setValue(statusBaru);
      return { success: true };
    }
  }
  return { success: false };
}

function togglePiutang(idPiutang, statusBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Piutang");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idPiutang) {
      sheet.getRange(i + 1, 5).setValue(statusBaru);
      return { success: true };
    }
  }
  return { success: false };
}
