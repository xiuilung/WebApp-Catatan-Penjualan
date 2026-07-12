/**
 * Fungsi utama untuk meluncurkan Aplikasi Web Apps Script.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('wiraniaga')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Membuat sheet database baru jika belum terdaftar, 
 * serta menginisiasi data default untuk uji coba awal.
 */
function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsConfig = {
    "MasterProduk": ["id", "nama", "sku", "harga"],
    "MasterArea": ["id", "nama", "kode"],
    "Sales": ["id", "tanggal", "namaToko", "area", "items", "grandTotal"],
    "Competitors": ["id", "tanggal", "area", "namaKompetitor", "programPromo", "produkTerdampakId"],
    "CompetitorMedia": ["id", "tanggal", "area", "namaKompetitor", "produk", "mediaBase64", "catatan"],
    "TargetQty": ["id", "produkId", "target"],
    "TargetValue": ["id", "target"],
    "TargetEC": ["id", "target"],
    "TargetECBrand": ["id", "produkId", "target"],
    "Users": ["id", "username", "password", "nama", "role"]
  };

  for (var sheetName in sheetsConfig) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheetsConfig[sheetName]);
    }
  }

  // Akun Default Admin & Sales (Hanya diisi jika kosong)
  var userSheet = ss.getSheetByName("Users");
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow(["u1", "admin", "admin123", "Administrator Pusat", "Admin"]);
    userSheet.appendRow(["u2", "sales", "sales123", "Sales Lapangan", "Sales"]);
  }

  // Katalog Produk Default
  var prodSheet = ss.getSheetByName("MasterProduk");
  if (prodSheet.getLastRow() === 1) {
    prodSheet.appendRow(["p1", "Kopi Bubuk Premium", "SKU-KBP-01", 50000]);
    prodSheet.appendRow(["p2", "Kopi Bubuk Arabika", "SKU-KBA-02", 75000]);
    prodSheet.appendRow(["p3", "Teh Melati Wangi", "SKU-TMW-03", 25000]);
  }

  // Area Kerja Lapangan Default
  var areaSheet = ss.getSheetByName("MasterArea");
  if (areaSheet.getLastRow() === 1) {
    areaSheet.appendRow(["a1", "Jakarta Selatan", "JKT-SEL"]);
    areaSheet.appendRow(["a2", "Jakarta Barat", "JKT-BAR"]);
    areaSheet.appendRow(["a3", "Bandung", "BDG-001"]);
  }

  // Target Omzet Global Awal (Default)
  var tvSheet = ss.getSheetByName("TargetValue");
  if (tvSheet.getLastRow() === 1) {
    tvSheet.appendRow(["global_val", 150000000]);
  }

  // Target EC Global Awal (Default)
  var tecSheet = ss.getSheetByName("TargetEC");
  if (tecSheet.getLastRow() === 1) {
    tecSheet.appendRow(["global_ec", 50]);
  }

  // Tambah Sales default agar tren perbandingan terlihat sejak awal
  var salesSheet = ss.getSheetByName("Sales");
  if (salesSheet.getLastRow() === 1) {
    // Penjualan Bulan Lalu (Juni 2026)
    salesSheet.appendRow([
      "s_init_1", "2026-06-20", "Toko Sejahtera", "Jakarta Selatan", 
      JSON.stringify([{ produkId: "p1", qty: 200, hargaTotal: 10000000 }]), 
      10000000
    ]);
    salesSheet.appendRow([
      "s_init_2", "2026-06-25", "Toko Makmur", "Jakarta Barat", 
      JSON.stringify([{ produkId: "p2", qty: 100, hargaTotal: 7500000 }]), 
      7500000
    ]);
    // Penjualan Minggu Lalu (22-28 Juni 2026)
    salesSheet.appendRow([
      "s_init_3", "2026-06-26", "Toko Maju Jaya", "Jakarta Selatan", 
      JSON.stringify([{ produkId: "p1", qty: 150, hargaTotal: 7500000 }]), 
      7500000
    ]);
    // Penjualan Minggu Ini / Bulan Ini (Juli 2026)
    salesSheet.appendRow([
      "s_init_4", "2026-07-02", "Toko Baru Hoki", "Jakarta Selatan", 
      JSON.stringify([{ produkId: "p1", qty: 300, hargaTotal: 15000000 }, { produkId: "p2", qty: 100, hargaTotal: 7500000 }]), 
      22500000
    ]);
    salesSheet.appendRow([
      "s_init_5", "2026-07-03", "Warung Berkah", "Bandung", 
      JSON.stringify([{ produkId: "p3", qty: 200, hargaTotal: 5000000 }]), 
      5000000
    ]);
  }
}

/**
 * Mengambil semua data dari Google Sheets untuk dikirimkan ke frontend React.
 */
function getAllData() {
  try {
    initDatabase(); // Pastikan seluruh database siap digunakan
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetsConfig = {
      "MasterProduk": ["id", "nama", "sku", "harga"],
      "MasterArea": ["id", "nama", "kode"],
      "Sales": ["id", "tanggal", "namaToko", "area", "items", "grandTotal"],
      "Competitors": ["id", "tanggal", "area", "namaKompetitor", "programPromo", "produkTerdampakId"],
      "CompetitorMedia": ["id", "tanggal", "area", "namaKompetitor", "produk", "mediaBase64", "catatan"],
      "TargetQty": ["id", "produkId", "target"],
      "TargetValue": ["id", "target"],
      "TargetEC": ["id", "target"],
      "TargetECBrand": ["id", "produkId", "target"],
      "Users": ["id", "username", "password", "nama", "role"]
    };

    var result = {};

    for (var sheetName in sheetsConfig) {
      var sheet = ss.getSheetByName(sheetName);
      var headers = sheetsConfig[sheetName];
      var data = [];

      if (sheet.getLastRow() > 1) {
        var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
        for (var i = 0; i < values.length; i++) {
          var rowObj = {};
          for (var j = 0; j < headers.length; j++) {
            var val = values[i][j];
            
            // Format format tanggal dan parsing array JSON
            if (val instanceof Date) {
              try {
                rowObj[headers[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
              } catch(e) {
                rowObj[headers[j]] = val.toISOString().split('T')[0];
              }
            } else if (headers[j] === 'items' && typeof val === 'string' && val !== '') {
              try {
                rowObj[headers[j]] = JSON.parse(val);
              } catch(e) {
                rowObj[headers[j]] = [];
              }
            } else {
              rowObj[headers[j]] = val;
            }
          }
          data.push(rowObj);
        }
      }
      result[sheetName] = data;
    }

    return result;
  } catch (err) {
    throw new Error("Gagal memuat data dari Spreadsheet: " + err.toString());
  }
}

/**
 * Menyimpan pembaruan array data dari React kembali ke Google Sheets.
 */
function saveData(sheetName, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Tab sheet tidak ditemukan" };

    var sheetsConfig = {
      "MasterProduk": ["id", "nama", "sku", "harga"],
      "MasterArea": ["id", "nama", "kode"],
      "Sales": ["id", "tanggal", "namaToko", "area", "items", "grandTotal"],
      "Competitors": ["id", "tanggal", "area", "namaKompetitor", "programPromo", "produkTerdampakId"],
      "CompetitorMedia": ["id", "tanggal", "area", "namaKompetitor", "produk", "mediaBase64", "catatan"],
      "TargetQty": ["id", "produkId", "target"],
      "TargetValue": ["id", "target"],
      "TargetEC": ["id", "target"],
      "TargetECBrand": ["id", "produkId", "target"],
      "Users": ["id", "username", "password", "nama", "role"]
    };

    var headers = sheetsConfig[sheetName];

    // Bersihkan data lama di bawah baris header
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
    }

    // Tulis data teranyar secara massal
    if (data && data.length > 0) {
      var rowsToWrite = [];
      for (var i = 0; i < data.length; i++) {
        var row = [];
        for (var j = 0; j < headers.length; j++) {
          var val = data[i][headers[j]];
          if (headers[j] === 'items') {
            row.push(JSON.stringify(val || []));
          } else {
            row.push(val !== undefined ? val : "");
          }
        }
        rowsToWrite.push(row);
      }
      sheet.getRange(2, 1, rowsToWrite.length, headers.length).setValues(rowsToWrite);
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Memverifikasi proses login pengguna berdasarkan tab 'Users'
 */
function loginUser(username, password) {
  try {
    initDatabase();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    var lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = 0; i < values.length; i++) {
        var dbUser = values[i][1];
        var dbPass = values[i][2];
        if (dbUser.toString().toLowerCase() === username.toLowerCase() && dbPass.toString() === password.toString()) {
          return {
            success: true,
            user: {
              id: values[i][0],
              username: values[i][1],
              nama: values[i][3],
              role: values[i][4]
            }
          };
        }
      }
    }
    return { success: false, message: "Username atau password salah!" };
  } catch(err) {
    return { success: false, message: "Koneksi database gagal: " + err.toString() };
  }
}

/**
 * Memanggil Gemini API untuk menghasilkan analisis taktis sales yang terstruktur.
 * Mengimplementasikan exponential backoff maksimal 5 kali retry.
 */
function generateAiRecommendations(competitorDataJson, masterProdukJson) {
  const apiKey = ""; // Disediakan otomatis oleh execution environment pada saat runtime
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=" + apiKey;
  
  const systemPrompt = "Anda adalah FMCG & Retail Distribution Sales Strategist handal di Indonesia. Tugas Anda adalah menganalisis program promosi kompetitor di lapangan dan memetakan respon taktis/program tandingan yang terarah, tajam, dan siap dieksekusi oleh tim sales di area terdampak.";
  
  const userQuery = "Berikut adalah temuan pergerakan kompetitor terbaru: " + competitorDataJson + 
                    "\n\nBerikut adalah Master Katalog Produk yang kita miliki: " + masterProdukJson + 
                    "\n\nPetunjuk Khusus: Petakan setiap temuan kompetitor dengan analisis ancaman, urgensi (Tinggi, Sedang, atau Rendah), dan 3 langkah strategi tangkisan taktis harian yang super detail, terarah, dan inovatif menggunakan bahasa Indonesia yang lugas.";

  const payload = {
    contents: [{ parts: [{ text: userQuery }] }],
    systemInstruction: { parts: [{ text: systemPrompt }] },
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          recommendations: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                competitorId: { type: "STRING" },
                area: { type: "STRING" },
                produk: { type: "STRING" },
                namaKompetitor: { type: "STRING" },
                analisisAncaman: { type: "STRING" },
                strategiTangkisan: {
                  type: "ARRAY",
                  items: { type: "STRING" }
                },
                urgensi: { type: "STRING" }
              },
              required: ["competitorId", "area", "produk", "namaKompetitor", "analisisAncaman", "strategiTangkisan", "urgensi"]
            }
          }
        },
        required: ["recommendations"]
      }
    }
  };

  let delay = 1000;
  for (let i = 0; i < 5; i++) {
    try {
      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const resText = response.getContentText();
      const resJson = JSON.parse(resText);
      
      if (response.getResponseCode() === 200) {
        const text = resJson.candidates?.[0]?.content?.parts?.[0]?.text;
        return text; // Mengembalikan string JSON mentah hasil dari model Gemini
      }
    } catch (e) {
      // Abaikan dan lanjutkan retry berikutnya
    }
    Utilities.sleep(delay);
    delay *= 2;
  }
  throw new Error("Gagal menghubungi AI Gemini setelah beberapa percobaan. Silakan coba sesaat lagi.");
}