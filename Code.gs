const SPREADSHEET_ID = '1eHH-MF7_OptTJyLxSTZRJDm0Hl4QvwGsHgnQW60cPvk';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('krankuindonesia - ERP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Setup function to initialize database schema (Sheets & Headers)
 * Run this function once from the Apps Script Editor.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const schema = {
    'Master_Produk': ['ID_Produk', 'Nama_Produk', 'Harga_Jual_Default'],
    'Master_Supplier': ['ID_Supplier', 'Nama', 'Alamat', 'No_Tlp'],
    'Master_Settings': ['Kunci', 'Nilai', 'Tipe_Persen_atau_Nominal'],
    'Master_Toko': ['ID_Toko', 'Marketplace', 'Nama_Toko', 'Fee_Ecommerce_Pct', 'Fee_GratisOngkir_Pct', 'Max_GratisOngkir', 'Fee_Promo_Pct', 'Max_Promo'],
    'Inbound_Stok': ['ID_Batch', 'Tanggal', 'ID_Produk', 'Nama_Supplier', 'No_Invoice_Supplier', 'Qty_Masuk', 'Qty_Sisa', 'Harga_Beli_Kotor', 'Diskon_Supplier', 'Harga_Beli_Bersih'],
    'Invoice_Header': ['No_Invoice', 'Tanggal', 'Shift', 'Tipe', 'Marketplace', 'Nama_Toko', 'Total_Omset', 'Total_Fee_Admin', 'Biaya_Pelayanan', 'Biaya_Lain', 'Total_HPP', 'Laba_Bersih', 'Status_Dana', 'Tanggal_Withdraw'],
    'Invoice_Detail': ['No_Invoice', 'ID_Produk', 'Qty', 'Harga_Jual_Aktual', 'Total_HPP_FIFO'],
    'Pengeluaran': ['Tanggal', 'Kategori', 'Nominal', 'Keterangan'],
    'Retur_Log': ['ID_Retur', 'Tanggal', 'No_Invoice_Asli', 'ID_Produk', 'Qty', 'Kondisi', 'Kerugian_Manual', 'Beban_Dipotong_Dari_Invoice_No'],
    'Master_User': ['Username', 'Password', 'Role'],
    'Buku_Kas': ['Tanggal', 'Tipe_Transaksi', 'Keterangan', 'Masuk_Debit', 'Keluar_Kredit', 'Saldo_Berjalan', 'Rekening']
  };

  for (const sheetName in schema) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // Set headers if the sheet is empty or we want to ensure headers exist
    const headers = schema[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
    
    // Basic formatting
    sheet.setFrozenRows(1);
  }
  
  // Optional: add initial settings if empty
  const settingsSheet = ss.getSheetByName('Master_Settings');
  if (settingsSheet.getLastRow() <= 1) {
    settingsSheet.getRange(2, 1, 5, 3).setValues([
      ['Fee_Ecommerce', 0.05, 'Persen'],
      ['Fee_GratisOngkir', 0.04, 'Persen'],
      ['Max_Ongkir', 10000, 'Nominal'],
      ['Fee_Promo_Extra', 0.01, 'Persen'],
      ['Tarif_Pelayanan_Per_Resi', 1250, 'Nominal']
    ]);
  }

  // Seed default admin user if Master_User is empty
  const userSheet = ss.getSheetByName('Master_User');
  if (userSheet && userSheet.getLastRow() <= 1) {
    userSheet.appendRow(['admin', 'admin123', 'Admin']);
  }
}

// --- SERIALIZATION HELPER ---
function serializeRow(headers, row) {
  let obj = {};
  headers.forEach((header, index) => {
    if (row[index] instanceof Date) {
      let d = row[index];
      obj[header] = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0') + "-" + String(d.getDate()).padStart(2,'0');
    } else {
      obj[header] = row[index];
    }
  });
  return obj;
}

// --- MASTER PRODUK CRUD ---
function getMasterProduk() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_Produk');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function saveMasterProduk(produk) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_Produk');
  const data = sheet.getDataRange().getValues();
  
  let exists = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === produk.ID_Produk) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[produk.Nama_Produk, produk.Harga_Jual_Default]]);
      exists = true;
      break;
    }
  }
  
  if (!exists) {
    sheet.appendRow([produk.ID_Produk, produk.Nama_Produk, produk.Harga_Jual_Default]);
  }
  return true;
}

function deleteMasterProduk(idProduk) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_Produk');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idProduk) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// --- MASTER SETTINGS ---
function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_Settings');
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for(let i=1; i<data.length; i++) {
    settings[data[i][0]] = { nilai: data[i][1], tipe: data[i][2] };
  }
  return settings;
}

// --- MASTER SUPPLIER CRUD ---
function getMasterSupplier() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_Supplier');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function saveMasterSupplier(supplier) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Master_Supplier');
  if(!sheet) {
      sheet = ss.insertSheet('Master_Supplier');
      sheet.appendRow(['ID_Supplier', 'Nama', 'Alamat', 'No_Tlp']);
  }
  
  if (!supplier.ID_Supplier) {
    supplier.ID_Supplier = 'SUP-' + new Date().getTime();
  }
  
  sheet.appendRow([supplier.ID_Supplier, supplier.Nama, supplier.Alamat, supplier.No_Tlp]);
  return true;
}

// --- MASTER TOKO CRUD ---
function getMasterToko() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_Toko');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function saveMasterToko(toko) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Master_Toko');
  if(!sheet) {
      sheet = ss.insertSheet('Master_Toko');
      sheet.appendRow(['ID_Toko', 'Marketplace', 'Nama_Toko', 'Fee_Ecommerce_Pct', 'Fee_GratisOngkir_Pct', 'Max_GratisOngkir', 'Fee_Promo_Pct', 'Max_Promo']);
  }
  
  if (!toko.ID_Toko) {
    toko.ID_Toko = 'TK-' + new Date().getTime();
  }
  
  const data = sheet.getDataRange().getValues();
  let exists = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === toko.ID_Toko) {
      sheet.getRange(i + 1, 2, 1, 7).setValues([[toko.Marketplace, toko.Nama_Toko, toko.Fee_Ecommerce_Pct, toko.Fee_GratisOngkir_Pct, toko.Max_GratisOngkir, toko.Fee_Promo_Pct, toko.Max_Promo]]);
      exists = true;
      break;
    }
  }
  
  if (!exists) {
    sheet.appendRow([toko.ID_Toko, toko.Marketplace, toko.Nama_Toko, toko.Fee_Ecommerce_Pct, toko.Fee_GratisOngkir_Pct, toko.Max_GratisOngkir, toko.Fee_Promo_Pct, toko.Max_Promo]);
  }
  return true;
}

function deleteMasterToko(idToko) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_Toko');
  if(!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idToko) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// --- INVENTORY (INBOUND STOK & FIFO) ---
function getInboundStok() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Inbound_Stok');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => serializeRow(headers, row));
}

function saveInboundStok(inbound) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Inbound_Stok');
  
  const tanggal = new Date(inbound.Tanggal);
  const qtyMasuk = parseInt(inbound.Qty_Masuk, 10);
  const hargaKotor = parseFloat(inbound.Harga_Beli_Kotor);
  const diskon = parseFloat(inbound.Diskon_Supplier);
  const hargaBersih = parseFloat(inbound.Harga_Beli_Bersih);
  
  if (!inbound.ID_Batch) {
    // Create new
    const idBatch = 'BATCH-' + new Date().getTime();
    const qtySisa = qtyMasuk;
    sheet.appendRow([
      idBatch, tanggal, inbound.ID_Produk, inbound.Nama_Supplier, inbound.No_Invoice_Supplier,
      qtyMasuk, qtySisa, hargaKotor, diskon, hargaBersih
    ]);
  } else {
    // Edit existing
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === inbound.ID_Batch) {
        // Update row (Columns 2 to 10)
        sheet.getRange(i + 1, 2, 1, 9).setValues([[
          tanggal, inbound.ID_Produk, inbound.Nama_Supplier, inbound.No_Invoice_Supplier,
          qtyMasuk, qtyMasuk, hargaKotor, diskon, hargaBersih
        ]]);
        break;
      }
    }
  }
  return true;
}

function deleteInboundStok(idBatch) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Inbound_Stok');
  const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idBatch) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
  return false;
}

function getInventorySummary() {
  const produkList = getMasterProduk();
  const stokList = getInboundStok();
  
  let summary = produkList.map(p => {
    let totalQty = 0;
    let totalAssetValue = 0;
    
    stokList.forEach(s => {
      if (s.ID_Produk === p.ID_Produk) {
        totalQty += s.Qty_Sisa;
        totalAssetValue += (s.Qty_Sisa * (s.Harga_Beli_Bersih || 0));
      }
    });
    
    return {
      ID_Produk: p.ID_Produk,
      Nama_Produk: p.Nama_Produk,
      Total_Stok: totalQty,
      Nilai_Aset: totalAssetValue
    };
  });
  
  return summary;
}

function getInitialData() {
  return {
    produk: getMasterProduk(),
    supplier: getMasterSupplier(),
    toko: getMasterToko(),
    stok: getInboundStok(),
    inventorySummary: getInventorySummary(),
    settings: getSettings(),
    invoices: getInvoiceHeaders(),
    invoiceDetails: getAllInvoiceDetails(),
    retur: getReturLog(),
    pengeluaran: getPengeluaran(),
    bukuKas: getBukuKasWithIdx(),
    users: getMasterUser()
  };
}

// --- INVOICE & FIFO ---
function saveInvoice(invoiceData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('Invoice_Header');
  const detailSheet = ss.getSheetByName('Invoice_Detail');
  const stokSheet = ss.getSheetByName('Inbound_Stok');
  
  const stokData = stokSheet.getDataRange().getValues();
  
  // Manual string formatting for timezone issues
  const d = new Date();
  const dateStr = d.getFullYear() + String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0');
  const timeStr = String(d.getHours()).padStart(2,'0') + String(d.getMinutes()).padStart(2,'0') + String(d.getSeconds()).padStart(2,'0');
  const noInvoice = 'INV-' + dateStr + '-' + timeStr;
  
  let totalOmset = 0;
  let totalHPP = 0;
  
  const header = invoiceData.header;
  const details = invoiceData.details;
  
  let stokUpdates = [];
  let detailRows = [];
  
  for(let i=0; i<details.length; i++) {
    const det = details[i];
    const qtyDibutuhkan = parseInt(det.Qty);
    let qtySisaUntukDipenuhi = qtyDibutuhkan;
    let hppItemTotal = 0;
    
    // Cari stok batch (FIFO) - dari atas ke bawah
    for(let j=1; j<stokData.length; j++) {
      if(qtySisaUntukDipenuhi <= 0) break;
      
      const row = stokData[j];
      const idProdukBatch = row[2];
      let qtySisaBatch = parseInt(row[6]);
      const hargaBeliBersih = parseFloat(row[9] || 0);
      
      if(idProdukBatch === det.ID_Produk && qtySisaBatch > 0) {
        let qtyDiambil = Math.min(qtySisaBatch, qtySisaUntukDipenuhi);
        
        hppItemTotal += (qtyDiambil * hargaBeliBersih);
        qtySisaUntukDipenuhi -= qtyDiambil;
        qtySisaBatch -= qtyDiambil;
        
        stokUpdates.push({ rowIdx: j + 1, colIdx: 7, val: qtySisaBatch }); // Kolom G
        stokData[j][6] = qtySisaBatch; 
      }
    }
    
    if (qtySisaUntukDipenuhi > 0) {
      throw new Error("Gagal: Stok tidak mencukupi untuk Produk " + det.ID_Produk + ". Sisa kurang: " + qtySisaUntukDipenuhi);
    }
    
    const omsetItem = qtyDibutuhkan * parseFloat(det.Harga_Jual_Aktual);
    totalOmset += omsetItem;
    totalHPP += hppItemTotal;
    
    detailRows.push([noInvoice, det.ID_Produk, qtyDibutuhkan, det.Harga_Jual_Aktual, hppItemTotal]);
  }
  
  // Eksekusi Update Stok (Batch)
  stokUpdates.forEach(u => {
    stokSheet.getRange(u.rowIdx, u.colIdx).setValue(u.val);
  });
  
  // Eksekusi Insert Detail
  if (detailRows.length > 0) {
    detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length).setValues(detailRows);
  }
  
  // Hitung Laba Bersih
  const feeAdmin = parseFloat(header.Total_Fee_Admin) || 0;
  const biayaPelayanan = parseFloat(header.Biaya_Pelayanan) || 0;
  const biayaLain = parseFloat(header.Biaya_Lain) || 0;
  
  const labaBersih = totalOmset - totalHPP - feeAdmin - biayaLain - biayaPelayanan;
  
  headerSheet.appendRow([
    noInvoice, new Date(header.Tanggal), header.Shift, header.Tipe, header.Marketplace, header.Nama_Toko,
    totalOmset, feeAdmin, biayaPelayanan, biayaLain, totalHPP, labaBersih,
    'Belum', '' // Status_Dana & Tanggal_Withdraw
  ]);
  
  
  return noInvoice;
}

function deleteInvoice(noInvoice) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('Invoice_Header');
  const detailSheet = ss.getSheetByName('Invoice_Detail');
  const stokSheet = ss.getSheetByName('Inbound_Stok');
  
  // 1. Get Details to reverse stock
  const details = getInvoiceDetails(noInvoice);
  const stokData = stokSheet.getDataRange().getValues();
  
  details.forEach(det => {
    let qtyToReturn = parseInt(det.Qty);
    // Return to the first batch found for this product (simplified)
    for (let j = 1; j < stokData.length; j++) {
      if (stokData[j][2] === det.ID_Produk) {
        let currentQtySisa = parseInt(stokData[j][6]);
        stokSheet.getRange(j + 1, 7).setValue(currentQtySisa + qtyToReturn);
        break;
      }
    }
  });
  
  // 2. Delete Header
  const headerData = headerSheet.getDataRange().getValues();
  for (let i = 1; i < headerData.length; i++) {
    if (headerData[i][0] === noInvoice) {
      headerSheet.deleteRow(i + 1);
      break;
    }
  }
  
  // 3. Delete Details
  const detailDataArr = detailSheet.getDataRange().getValues();
  for (let i = detailDataArr.length - 1; i >= 1; i--) {
    if (detailDataArr[i][0] === noInvoice) {
      detailSheet.deleteRow(i + 1);
    }
  }
  
  return { success: true };
}

// --- FINANCE & RETUR ---
function getInvoiceHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Invoice_Header');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => serializeRow(headers, row));
}

function processWithdraw(noInvoice, rekening) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Invoice_Header');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === noInvoice && data[i][12] === 'Belum') { // Kolom M: Status_Dana
      sheet.getRange(i + 1, 13).setValue('Sudah');
      sheet.getRange(i + 1, 14).setValue(new Date()); // Kolom N: Tanggal_Withdraw
      
      // HOOK: Append to Buku_Kas - gunakan Laba_Bersih (Net), bukan Total_Omset (Gross)
      // Invoice_Header kolom: [0]=No_Invoice,[1]=Tanggal,[2]=Shift,[3]=Tipe,[4]=Marketplace,
      // [5]=Nama_Toko,[6]=Total_Omset,[7]=Total_Fee_Admin,[8]=Biaya_Pelayanan,
      // [9]=Biaya_Lain,[10]=Total_HPP,[11]=Laba_Bersih,[12]=Status_Dana,[13]=Tanggal_Withdraw
      const labaBersih = parseFloat(data[i][11]) || 0; // Index 11 = Laba_Bersih (Net Revenue)
      appendBukuKas(
        new Date(),
        'Pencairan Penjualan',
        'Withdraw invoice ' + noInvoice,
        labaBersih,  // <-- NET revenue, bukan Gross
        0,
        rekening || 'BCA'
      );
      
      return true;
    }
  }
  return false;
}

function getInvoiceDetails(noInvoice) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Invoice_Detail');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  
  let details = [];
  rows.forEach(row => {
    if (row[0] === noInvoice) {
      details.push(serializeRow(headers, row));
    }
  });
  return details;
}

function getAllInvoiceDetails() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Invoice_Detail');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => serializeRow(headers, row));
}

function getReturLog() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Retur_Log');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => serializeRow(headers, row));
}

function processRetur(returData) {
  // returData = { No_Invoice_Asli, Target_Invoice_Potongan, ID_Produk, Qty, Kondisi, Kerugian_Manual }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const returSheet = ss.getSheetByName('Retur_Log');
  const stokSheet = ss.getSheetByName('Inbound_Stok');
  const detailSheet = ss.getSheetByName('Invoice_Detail');
  const headerSheet = ss.getSheetByName('Invoice_Header');
  
  const idRetur = 'RTR-' + new Date().getTime();
  const tanggal = new Date();
  let qtyRetur = parseInt(returData.Qty);
  
  // Find original detail for HPP and Harga Jual
  let hargaBeliBersihPerItem = 0;
  let hargaJualAktual = 0;
  const detailData = detailSheet.getDataRange().getValues();
  for (let i = 1; i < detailData.length; i++) {
    if (detailData[i][0] === returData.No_Invoice_Asli && detailData[i][1] === returData.ID_Produk) {
        let qtyTerjual = parseInt(detailData[i][2]);
        hargaJualAktual = parseFloat(detailData[i][3]) || 0;
        let totalHPP = parseFloat(detailData[i][4]);
        if(qtyTerjual > 0) hargaBeliBersihPerItem = totalHPP / qtyTerjual;
        break;
    }
  }
  
  // Deduct from Target Invoice Laba Bersih
  const deductionAmount = qtyRetur * hargaJualAktual;
  const headerData = headerSheet.getDataRange().getValues();
  for (let i = 1; i < headerData.length; i++) {
    if (headerData[i][0] === returData.Target_Invoice_Potongan) {
       let currentLabaBersih = parseFloat(headerData[i][11]) || 0;
       headerSheet.getRange(i + 1, 12).setValue(currentLabaBersih - deductionAmount);
       break;
    }
  }
  
  // Record Retur Log
  returSheet.appendRow([
    idRetur, tanggal, returData.No_Invoice_Asli, returData.ID_Produk, qtyRetur, returData.Kondisi, returData.Kerugian_Manual, returData.Target_Invoice_Potongan
  ]);
  
  // If "Bagus" -> Return to Stock to the LATEST active batch
  if (returData.Kondisi === 'Bagus') {
      const stokData = stokSheet.getDataRange().getValues();
      let foundBatch = false;
      // Search backwards
      for (let i = stokData.length - 1; i >= 1; i--) {
         if (stokData[i][2] === returData.ID_Produk) {
            let currentQtyMasuk = parseInt(stokData[i][5]) || 0;
            let currentQtySisa = parseInt(stokData[i][6]) || 0;
            stokSheet.getRange(i + 1, 6).setValue(currentQtyMasuk + qtyRetur);
            stokSheet.getRange(i + 1, 7).setValue(currentQtySisa + qtyRetur);
            foundBatch = true;
            break;
         }
      }
      if (!foundBatch) {
         // Fallback
         const idBatch = 'RSTK-' + new Date().getTime();
         stokSheet.appendRow([
           idBatch, tanggal, returData.ID_Produk, 'RETUR CUSTOMER', returData.No_Invoice_Asli,
           qtyRetur, qtyRetur, hargaBeliBersihPerItem, 0, hargaBeliBersihPerItem
         ]);
      }
  }
  
  // Kerugian Manual -> Pengeluaran
  let kerugian = parseFloat(returData.Kerugian_Manual);
  if (kerugian > 0) {
      savePengeluaran({
          Tanggal: tanggal,
          Kategori: 'Kerugian Retur',
          Nominal: kerugian,
          Keterangan: `Retur produk ${returData.ID_Produk} dari invoice ${returData.No_Invoice_Asli}`
      });
  }
  
  return idRetur;
}

// --- PENGELUARAN CRUD ---
function getPengeluaran() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Pengeluaran');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => serializeRow(headers, row));
}

function savePengeluaran(pengeluaran) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Pengeluaran');
  if(!sheet) {
      sheet = ss.insertSheet('Pengeluaran');
      sheet.appendRow(['Tanggal', 'Kategori', 'Nominal', 'Keterangan', 'Rekening']);
  }
  const tgl = new Date(pengeluaran.Tanggal);
  const nominal = parseFloat(pengeluaran.Nominal) || 0;
  const rekening = pengeluaran.Rekening || 'BCA';
  
  // Validasi: cek saldo bank cukup
  const balances = getBalancesPerBank();
  const saldoRekening = balances[rekening.toUpperCase()] || 0;
  const saldoTotal = Object.values(balances).reduce((a, b) => a + b, 0);
  
  if (saldoTotal < nominal) {
    return { success: false, message: `Saldo keseluruhan tidak cukup (Saldo: Rp ${saldoTotal.toLocaleString()}, Dibutuhkan: Rp ${nominal.toLocaleString()})` };
  }
  if (saldoRekening < nominal) {
    return { success: false, message: `Saldo ${rekening} tidak cukup (Saldo: Rp ${saldoRekening.toLocaleString()}, Dibutuhkan: Rp ${nominal.toLocaleString()}). Silakan pilih rekening lain.` };
  }
  
  // Cek apakah kolom Rekening sudah ada, jika belum tambahkan
  const headers = sheet.getDataRange().getValues()[0];
  if (headers.length < 5) {
    sheet.getRange(1, 5).setValue('Rekening');
  }
  
  sheet.appendRow([tgl, pengeluaran.Kategori, nominal, pengeluaran.Keterangan, rekening]);
  
  // HOOK: Append to Buku_Kas as Pengeluaran
  appendBukuKas(
    tgl,
    'Pengeluaran',
    pengeluaran.Kategori + (pengeluaran.Keterangan ? ' - ' + pengeluaran.Keterangan : ''),
    0,
    nominal,
    rekening
  );
  
  return { success: true };
}

function deletePengeluaranByData(peng) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Pengeluaran');
  if(!sheet) return { success: false, message: 'Sheet Pengeluaran tidak ditemukan.' };
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
    let tglMatch = false;
    if(data[i][0] instanceof Date) {
      let d = data[i][0];
      let str = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,'0') + "-" + String(d.getDate()).padStart(2,'0');
      if(str === peng.Tanggal) tglMatch = true;
    } else {
      if(String(data[i][0]) === String(peng.Tanggal)) tglMatch = true;
    }
    
    if(tglMatch && 
       String(data[i][1]) === String(peng.Kategori) && 
       parseFloat(data[i][2]) === parseFloat(peng.Nominal) && 
       String(data[i][3]) === String(peng.Keterangan)) {
      sheet.deleteRow(i+1);
      
      // JUGA hapus entri terkait dari Buku_Kas
      _deleteBukuKasMatchingPengeluaran(peng);
      
      return { success: true };
    }
  }
  return { success: false, message: 'Data pengeluaran tidak ditemukan.' };
}

// Helper: hapus entri Buku_Kas yang cocok dengan pengeluaran ini
function _deleteBukuKasMatchingPengeluaran(peng) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Buku_Kas');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 1; i--) {
    let tglMatch = false;
    let tglCell = data[i][0];
    if (tglCell instanceof Date) {
      let str = tglCell.getFullYear() + "-" + String(tglCell.getMonth()+1).padStart(2,'0') + "-" + String(tglCell.getDate()).padStart(2,'0');
      if (str === peng.Tanggal) tglMatch = true;
    } else {
      if (String(tglCell) === String(peng.Tanggal)) tglMatch = true;
    }
    
    const tipe = String(data[i][1] || '');
    const keluar = parseFloat(data[i][4]) || 0;
    
    if (tglMatch && tipe === 'Pengeluaran' && keluar === parseFloat(peng.Nominal)) {
      // Hapus baris ini dan rekalkuasi saldo sesudahnya
      sheet.deleteRow(i + 1);
      _recalculateBukuKasSaldo(sheet, i);
      return;
    }
  }
}

// Rekalkuasi saldo berjalan Buku_Kas mulai dari baris startRow (1-indexed dari baris data)
function _recalculateBukuKasSaldo(sheet, startDataIdx) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  // Temukan saldo sebelum startDataIdx
  let saldoBefore = 0;
  const sheetStartRow = startDataIdx + 1; // Convert to 1-indexed sheet row
  if (startDataIdx >= 1) {
    // Baris sebelum yang dihapus
    const prevRow = startDataIdx - 1; // baris data sebelumnya (0-indexed dari data array, 1-indexed dari baris sheet)
    if (prevRow >= 1) {
      saldoBefore = parseFloat(data[prevRow][5]) || 0;
    }
  }
  
  // Update semua baris dari startDataIdx sampai akhir
  for (let i = startDataIdx; i < data.length; i++) {
    const masuk = parseFloat(data[i][3]) || 0;
    const keluar = parseFloat(data[i][4]) || 0;
    saldoBefore = saldoBefore + masuk - keluar;
    sheet.getRange(i + 1, 6).setValue(saldoBefore);
  }
}

function getTopProducts(start, end) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('Invoice_Header');
  const detailSheet = ss.getSheetByName('Invoice_Detail');
  
  if(!headerSheet || !detailSheet) return [];
  
  const headers = headerSheet.getDataRange().getValues();
  const details = detailSheet.getDataRange().getValues();
  
  // Find valid invoice numbers within date range
  const validInvoices = new Set();
  for(let i=1; i<headers.length; i++) {
    let tglMatch = true;
    let tgl = headers[i][1];
    let str = "";
    if(tgl instanceof Date) {
      str = tgl.getFullYear() + "-" + String(tgl.getMonth()+1).padStart(2,'0') + "-" + String(tgl.getDate()).padStart(2,'0');
    } else {
      str = String(tgl);
    }
    
    if(start && str < start) tglMatch = false;
    if(end && str > end) tglMatch = false;
    
    if(tglMatch) validInvoices.add(headers[i][0]);
  }
  
  // Aggregate Qty by ID_Produk
  const productQtys = {};
  for(let i=1; i<details.length; i++) {
    const noInv = details[i][0];
    if(validInvoices.has(noInv)) {
       const idProd = details[i][1];
       const qty = parseInt(details[i][2]) || 0;
       if(!productQtys[idProd]) productQtys[idProd] = 0;
       productQtys[idProd] += qty;
    }
  }
  
  // Sort and pick top 5
  const sorted = Object.keys(productQtys)
    .map(id => ({ id: id, qty: productQtys[id] }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 5);
    
  return sorted;
}

// --- LOGIN VERIFICATION ---
function verifyLogin(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_User');
  if (!sheet) return { success: false, message: 'Tabel user belum dibuat. Jalankan setupDatabase().' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim() &&
        String(data[i][1]).trim() === String(password).trim()) {
      return { success: true, username: data[i][0], role: data[i][2] };
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

function getMasterUser() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_User');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => { obj[header] = row[index]; });
    return obj;
  });
}

function saveMasterUser(user) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_User');
  if (!sheet) throw new Error('Master_User sheet not found. Run setupDatabase().');
  
  if (!user.Username || !user.Password || !user.Role) {
    throw new Error('Username, Password, dan Role wajib diisi.');
  }
  
  // Check duplicate username
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(user.Username).trim()) {
      throw new Error('Username sudah digunakan.');
    }
  }
  
  sheet.appendRow([user.Username, user.Password, user.Role]);
  return true;
}

function deleteMasterUser(username) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Master_User');
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim()) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// --- BATCH INBOUND STOK (Multi-Item Cart) ---
function saveBatchInbound(cartData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Inbound_Stok');
  
  const tanggal = new Date(cartData.tanggal);
  const namaSupplier = cartData.supplier;
  const noInvoicePO = cartData.noInvoicePO;
  const items = cartData.items;
  
  if (!items || items.length === 0) {
    throw new Error('Tidak ada item dalam keranjang.');
  }
  
  const baseTimestamp = new Date().getTime();
  let rows = [];
  
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const idBatch = 'BATCH-' + baseTimestamp + '-' + i;
    const qtyMasuk = parseInt(item.Qty_Masuk, 10);
    const hargaKotor = parseFloat(item.Harga_Beli_Kotor);
    const diskon = parseFloat(item.Diskon_Supplier);
    const hargaBersih = parseFloat(item.Harga_Beli_Bersih);
    
    rows.push([
      idBatch, tanggal, item.ID_Produk, namaSupplier, noInvoicePO,
      qtyMasuk, qtyMasuk, hargaKotor, diskon, hargaBersih
    ]);
  }
  
  // Batch insert all rows at once
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  return { success: true, count: rows.length, noInvoicePO: noInvoicePO };
}

// --- BUKU KAS (CASH FLOW LEDGER) ---

/**
 * Appends a row to the Buku_Kas ledger, auto-calculating Saldo_Berjalan.
 */
function appendBukuKas(tanggal, tipe, keterangan, masuk, keluar, rekening) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Buku_Kas');
  if (!sheet) {
    sheet = ss.insertSheet('Buku_Kas');
    sheet.appendRow(['Tanggal', 'Tipe_Transaksi', 'Keterangan', 'Masuk_Debit', 'Keluar_Kredit', 'Saldo_Berjalan', 'Rekening']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  
  const lastRow = sheet.getLastRow();
  let saldoSebelumnya = 0;
  if (lastRow > 1) {
    saldoSebelumnya = parseFloat(sheet.getRange(lastRow, 6).getValue()) || 0;
  }
  
  const saldoBaru = saldoSebelumnya + (parseFloat(masuk) || 0) - (parseFloat(keluar) || 0);
  sheet.appendRow([tanggal, tipe, keterangan, masuk, keluar, saldoBaru, rekening || '']);
  return saldoBaru;
}

function getBukuKas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Buku_Kas');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => serializeRow(headers, row));
}

function getSaldoKas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Buku_Kas');
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  return parseFloat(sheet.getRange(lastRow, 6).getValue()) || 0;
}

function tambahModal(data) {
  const tgl = new Date(data.Tanggal);
  const nominal = parseFloat(data.Nominal) || 0;
  const keterangan = data.Keterangan || 'Penambahan Modal';
  const rekening = data.Rekening || 'BCA';
  
  if (nominal <= 0) throw new Error('Nominal modal harus lebih dari 0.');
  
  const saldoBaru = appendBukuKas(tgl, 'Modal', keterangan, nominal, 0, rekening);
  return { success: true, saldoBaru: saldoBaru };
}

function getBalancesPerBank() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Buku_Kas');
    const balances = { BCA: 0, BNI: 0, BRI: 0, MANDIRI: 0 };
    
    if (!sheet) return balances;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return balances;
    
    // Header check to find columns index safely
    const headers = data[0];
    const colMasuk = headers.indexOf('Masuk_Debit');
    const colKeluar = headers.indexOf('Keluar_Kredit');
    const colRekening = headers.indexOf('Rekening');
    
    if (colMasuk === -1 || colKeluar === -1 || colRekening === -1) return balances;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const masuk = parseFloat(row[colMasuk]) || 0;
      const keluar = parseFloat(row[colKeluar]) || 0;
      const rekRaw = String(row[colRekening] || '');
      const rek = rekRaw.trim().toUpperCase();
      
      if (balances.hasOwnProperty(rek)) {
        balances[rek] += (masuk - keluar);
      }
    }
    return balances;
  } catch (e) {
    console.error('Error in getBalancesPerBank:', e.message);
    return { BCA: 0, BNI: 0, BRI: 0, MANDIRI: 0 };
  }
}

function deleteInboundData(idBatch) {
  try {
    const res = deleteInboundStok(idBatch);
    if(res) return { success: true };
    return { success: false, message: 'Gagal menghapus. Pastikan stok belum ada yang terjual.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateData(id, type, data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetName = "";
  let pkIndex = 0; 
  
  switch(type) {
    case 'user': sheetName = 'Master_User'; pkIndex = 0; break;
    case 'produk': sheetName = 'Master_Produk'; pkIndex = 0; break;
    case 'toko': sheetName = 'Master_Toko'; pkIndex = 0; break;
    case 'supplier': sheetName = 'Master_Supplier'; pkIndex = 0; break;
    case 'inbound': sheetName = 'Inbound_Stok'; pkIndex = 0; break;
    case 'invoice': sheetName = 'Invoice_Header'; pkIndex = 0; break;
    case 'retur': sheetName = 'Retur_Log'; pkIndex = 0; break;
  }
  
  if(!sheetName) return { success: false, message: 'Invalid type' };
  
  const sheet = ss.getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  
  for(let i=1; i<rows.length; i++) {
    if(String(rows[i][pkIndex]) === String(id)) {
      const updateRow = headers.map((h, idx) => {
        if(data.hasOwnProperty(h)) {
           let val = data[h];
           if(h.toLowerCase().includes('tanggal') && val) return new Date(val);
           return val;
        }
        return rows[i][idx];
      });
      sheet.getRange(i+1, 1, 1, updateRow.length).setValues([updateRow]);
      return { success: true };
    }
  }
  return { success: false, message: 'ID not found' };
}

function deleteData(id, type) {
  switch(type) {
    case 'user': return { success: deleteMasterUser(id) };
    case 'produk': return { success: deleteMasterProduk(id) };
    case 'toko': return { success: deleteMasterToko(id) };
    case 'inbound': return deleteInboundData(id);
    case 'invoice': return deleteInvoice(id);
    case 'retur': return deleteReturData(id);
    default: return { success: false, message: 'Type handler not implemented' };
  }
}

function deleteReturData(idRetur) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Retur_Log');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idRetur) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'ID Retur tidak ditemukan.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function saveTambahModal(data) {
  try {
    const tgl = new Date(data.Tanggal);
    const nominal = parseFloat(data.Nominal);
    appendBukuKas(tgl, 'Modal', data.Keterangan, nominal, 0, data.Rekening);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateSalesData(invoiceId, updatedData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const headerSheet = ss.getSheetByName('Invoice_Header');
  const detailSheet = ss.getSheetByName('Invoice_Detail');
  const stokSheet = ss.getSheetByName('Inbound_Stok');
  
  // 1. REVERT STOCK
  const oldDetails = getInvoiceDetails(invoiceId);
  const stokData = stokSheet.getDataRange().getValues();
  
  oldDetails.forEach(det => {
    let qtyToReturn = parseInt(det.Qty);
    for (let j = 1; j < stokData.length; j++) {
      if (stokData[j][2] === det.ID_Produk) {
        let currentQtySisa = parseInt(stokData[j][6]);
        stokSheet.getRange(j + 1, 7).setValue(currentQtySisa + qtyToReturn);
        stokData[j][6] = currentQtySisa + qtyToReturn; // Update in-memory for next allocation
        break;
      }
    }
  });
  
  // 2. DELETE OLD DETAILS
  const detailDataArr = detailSheet.getDataRange().getValues();
  for (let i = detailDataArr.length - 1; i >= 1; i--) {
    if (detailDataArr[i][0] === invoiceId) {
      detailSheet.deleteRow(i + 1);
    }
  }
  
  // 3. ALLOCATE NEW STOCK (FIFO)
  let totalOmset = 0;
  let totalHPP = 0;
  const header = updatedData.header;
  const details = updatedData.details;
  let stokUpdates = [];
  let newDetailRows = [];
  
  for(let i=0; i<details.length; i++) {
    const det = details[i];
    const qtyDibutuhkan = parseInt(det.Qty);
    let qtySisaUntukDipenuhi = qtyDibutuhkan;
    let hppItemTotal = 0;
    
    for(let j=1; j<stokData.length; j++) {
      if(qtySisaUntukDipenuhi <= 0) break;
      const row = stokData[j];
      if(row[2] === det.ID_Produk && parseInt(row[6]) > 0) {
        let qtyDiambil = Math.min(parseInt(row[6]), qtySisaUntukDipenuhi);
        hppItemTotal += (qtyDiambil * parseFloat(row[9] || 0));
        qtySisaUntukDipenuhi -= qtyDiambil;
        row[6] = parseInt(row[6]) - qtyDiambil;
        stokUpdates.push({ rowIdx: j + 1, colIdx: 7, val: row[6] });
      }
    }
    
    if (qtySisaUntukDipenuhi > 0) {
      throw new Error("Stok tidak cukup untuk " + det.ID_Produk);
    }
    
    totalOmset += (qtyDibutuhkan * parseFloat(det.Harga_Jual_Aktual));
    totalHPP += hppItemTotal;
    newDetailRows.push([invoiceId, det.ID_Produk, qtyDibutuhkan, det.Harga_Jual_Aktual, hppItemTotal]);
  }
  
  // 4. SAVE STOCK UPDATES
  stokUpdates.forEach(u => {
    stokSheet.getRange(u.rowIdx, u.colIdx).setValue(u.val);
  });
  
  // 5. INSERT NEW DETAILS
  if (newDetailRows.length > 0) {
    detailSheet.getRange(detailSheet.getLastRow() + 1, 1, newDetailRows.length, 5).setValues(newDetailRows);
  }
  
  // 6. UPDATE HEADER
  const feeAdmin = parseFloat(header.Total_Fee_Admin) || 0;
  const biayaPelayanan = parseFloat(header.Biaya_Pelayanan) || 0;
  const biayaLain = parseFloat(header.Biaya_Lain) || 0;
  const labaBersih = totalOmset - totalHPP - feeAdmin - biayaLain - biayaPelayanan;
  
  const headerData = headerSheet.getDataRange().getValues();
  for (let i = 1; i < headerData.length; i++) {
    if (headerData[i][0] === invoiceId) {
      headerSheet.getRange(i + 1, 2, 1, 11).setValues([[
        new Date(header.Tanggal),
        header.Shift,
        header.Tipe,
        header.Marketplace,
        header.Nama_Toko,
        totalOmset,
        feeAdmin,
        biayaPelayanan,
        biayaLain,
        totalHPP,
        labaBersih
      ]]);
      break;
    }
  }
  
  return { success: true, noInvoice: invoiceId };
}

// ============================================================
// FUNGSI BARU: PENGELUARAN EDIT, WITHDRAW REVERT & UPDATE, BUKU KAS CRUD
// ============================================================

/**
 * Update pengeluaran yang ada (hapus lama, buat baru).
 * oldPeng: data lama {Tanggal, Kategori, Nominal, Keterangan, Rekening}
 * newPeng: data baru {Tanggal, Kategori, Nominal, Keterangan, Rekening}
 */
function updatePengeluaran(oldPeng, newPeng) {
  try {
    // 1. Hapus pengeluaran & entri Buku_Kas lama
    deletePengeluaranByData(oldPeng);
    
    // 2. Simpan pengeluaran baru (savePengeluaran sudah validasi saldo)
    const result = savePengeluaran(newPeng);
    if (!result.success) return result;
    
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Mengembalikan invoice dari Status_Dana='Sudah' ke 'Belum' (revert withdraw).
 * Juga menghapus entri Buku_Kas pencairan terkait.
 */
function revertWithdraw(noInvoice) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const headerSheet = ss.getSheetByName('Invoice_Header');
    const data = headerSheet.getDataRange().getValues();
    
    let labaBersih = 0;
    let withdrawDate = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === noInvoice && data[i][12] === 'Sudah') {
        labaBersih = parseFloat(data[i][11]) || 0;
        withdrawDate = data[i][13]; // Tanggal_Withdraw
        // Reset status
        headerSheet.getRange(i + 1, 13).setValue('Belum');
        headerSheet.getRange(i + 1, 14).setValue('');
        break;
      }
    }
    
    // Hapus entri Buku_Kas pencairan ini
    const kasSheet = ss.getSheetByName('Buku_Kas');
    if (kasSheet && withdrawDate) {
      const kasData = kasSheet.getDataRange().getValues();
      for (let i = kasData.length - 1; i >= 1; i--) {
        const tipe = String(kasData[i][1] || '');
        const keterangan = String(kasData[i][2] || '');
        if (tipe === 'Pencairan Penjualan' && keterangan.includes(noInvoice)) {
          kasSheet.deleteRow(i + 1);
          _recalculateBukuKasSaldo(kasSheet, i);
          break;
        }
      }
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Edit nominal Laba_Bersih pada Invoice_Header yang sudah cair.
 * Juga update entri Buku_Kas terkait.
 */
function updateWithdrawLaba(noInvoice, nominalBaru) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const headerSheet = ss.getSheetByName('Invoice_Header');
    const data = headerSheet.getDataRange().getValues();
    
    const nominal = parseFloat(nominalBaru) || 0;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === noInvoice) {
        // Update Laba_Bersih di kolom 12 (index 11)
        headerSheet.getRange(i + 1, 12).setValue(nominal);
        break;
      }
    }
    
    // Update entri Buku_Kas pencairan ini
    const kasSheet = ss.getSheetByName('Buku_Kas');
    if (kasSheet) {
      const kasData = kasSheet.getDataRange().getValues();
      for (let i = 1; i < kasData.length; i++) {
        const tipe = String(kasData[i][1] || '');
        const keterangan = String(kasData[i][2] || '');
        if (tipe === 'Pencairan Penjualan' && keterangan.includes(noInvoice)) {
          const oldMasuk = parseFloat(kasData[i][3]) || 0;
          kasSheet.getRange(i + 1, 4).setValue(nominal);
          // Rekalkuasi saldo dari baris ini
          _recalculateBukuKasSaldo(kasSheet, i);
          break;
        }
      }
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Update satu entri di Buku_Kas berdasarkan index baris (row number sheet, 1-indexed, header=1).
 */
function updateBukuKasEntry(rowIdx, updatedEntry) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Buku_Kas');
    if (!sheet) return { success: false, message: 'Sheet Buku_Kas tidak ditemukan.' };
    
    const sheetRow = parseInt(rowIdx) + 1; // +1 karena header di baris 1
    if (sheetRow < 2 || sheetRow > sheet.getLastRow()) {
      return { success: false, message: 'Baris tidak valid.' };
    }
    
    const tgl = new Date(updatedEntry.Tanggal);
    const masuk = parseFloat(updatedEntry.Masuk_Debit) || 0;
    const keluar = parseFloat(updatedEntry.Keluar_Kredit) || 0;
    
    sheet.getRange(sheetRow, 1, 1, 7).setValues([[
      tgl,
      updatedEntry.Tipe_Transaksi,
      updatedEntry.Keterangan,
      masuk,
      keluar,
      0, // Saldo akan direkalkuasi
      updatedEntry.Rekening || ''
    ]]);
    
    // Rekalkuasi saldo dari baris ini sampai akhir
    _recalculateBukuKasSaldo(sheet, sheetRow - 1); // sheetRow-1 = data index (0-indexed from data array)
    
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Hapus satu entri di Buku_Kas berdasarkan index baris (row number sheet).
 */
function deleteBukuKasEntry(rowIdx) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Buku_Kas');
    if (!sheet) return { success: false, message: 'Sheet Buku_Kas tidak ditemukan.' };
    
    const sheetRow = parseInt(rowIdx) + 1; // +1 karena header di baris 1
    if (sheetRow < 2 || sheetRow > sheet.getLastRow()) {
      return { success: false, message: 'Baris tidak valid.' };
    }
    
    const dataIdxAfterDelete = sheetRow - 1 - 1; // index data setelah baris dihapus
    sheet.deleteRow(sheetRow);
    _recalculateBukuKasSaldo(sheet, dataIdxAfterDelete);
    
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// getInitialData juga perlu menyertakan rowIdx untuk Buku_Kas agar frontend bisa referensi baris
function getBukuKasWithIdx() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Buku_Kas');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map((row, idx) => {
    let obj = serializeRow(headers, row);
    obj._rowIdx = idx + 1; // 1-indexed dari data array (baris sheet = idx+2)
    return obj;
  });
}
