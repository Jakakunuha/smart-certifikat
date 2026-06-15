// =========================================================================
// BACKEND ENGINE SMART-CERT - INTEGRASI GOOGLE DRIVE & SHEETS
// =========================================================================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let sheetUsers = ss.getSheetByName("Users");
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet("Users");
    sheetUsers.appendRow(["Username", "Password", "Role"]);
    sheetUsers.appendRow(["admin", "admin123", "admin"]);
  }

  let sheetData = ss.getSheetByName("Data_Sertifikat");
  if (!sheetData) {
    sheetData = ss.insertSheet("Data_Sertifikat");
    sheetData.appendRow([
      "Timestamp",
      "Nomor_Sertifikat",
      "Nama_Peserta",
      "Email",
      "No_HP",
      "Status_Sertifikat",
      "Link_Download_Drive",
    ]);
  }

  // --- SHEET BARU UNTUK MENYIMPAN PENGATURAN DESAIN CLOUD ---
  let sheetPengaturan = ss.getSheetByName("Pengaturan");
  if (!sheetPengaturan) {
    sheetPengaturan = ss.insertSheet("Pengaturan");
    sheetPengaturan.appendRow([
      "ID_Konfigurasi",
      "File_ID_Drive",
      "X_Pos",
      "Y_Pos",
      "Font_Size",
      "Color",
    ]);
  }
}

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (action === "ambil_semua") {
    const sheet = ss.getSheetByName("Data_Sertifikat");
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const hasilJson = [];
    for (let i = 1; i < values.length; i++) {
      let obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = values[i][j];
      }
      hasilJson.push(obj);
    }
    return buatResponJSON({ status: "success", data: hasilJson });
  }

  // --- FUNGSI MENGAMBIL DESAIN DARI CLOUD (DRIVE KE BASE64) ---
  if (action === "ambil_pengaturan") {
    const sheet = ss.getSheetByName("Pengaturan");
    if (sheet.getLastRow() < 2) return buatResponJSON({ status: "empty" });

    const data = sheet.getRange(2, 1, 1, 6).getValues()[0];
    const fileId = data[1];
    let base64 = "";

    // Membaca file dari Google Drive dan mengubahnya kembali menjadi Base64 untuk Web
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        const b64Str = Utilities.base64Encode(blob.getBytes());
        base64 = "data:" + blob.getContentType() + ";base64," + b64Str;
      } catch (err) {
        // Abaikan jika file terhapus di Drive
      }
    }

    const config = {
      xPos: data[2],
      yPos: data[3],
      fontSize: data[4],
      color: data[5],
      templateBase64: base64,
    };
    return buatResponJSON({ status: "success", data: config });
  }

  return buatResponJSON({ status: "error", message: "Aksi tidak dikenali" });
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const payload = postData.data;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === "tambah_sertifikat") {
      const sheet = ss.getSheetByName("Data_Sertifikat");
      sheet.appendRow([
        payload.Timestamp || new Date().toLocaleString("id-ID"),
        payload.Nomor_Sertifikat || "-",
        payload.Nama_Peserta || "-",
        payload.Email || "-",
        payload.No_HP || "-",
        "Tercetak",
        "",
      ]);
      return buatResponJSON({ status: "success" });
    }

    if (action === "tambah_sertifikat_bulk") {
      const sheet = ss.getSheetByName("Data_Sertifikat");
      const arrData = payload;
      if (!arrData || arrData.length === 0) throw new Error("Data kosong");

      const rowsToAppend = arrData.map((item) => [
        item.Timestamp || new Date().toLocaleString("id-ID"),
        item.Nomor_Sertifikat || "-",
        item.Nama_Peserta || "-",
        item.Email || "-",
        item.No_HP || "-",
        "Tercetak",
        "",
      ]);
      sheet
        .getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 7)
        .setValues(rowsToAppend);
      return buatResponJSON({
        status: "success",
        message: "Bulk insert sukses",
      });
    }

    // --- FUNGSI BARU: MENYIMPAN GAMBAR KE GOOGLE DRIVE & DATA KE SPREADSHEET ---
    if (action === "simpan_pengaturan") {
      const sheet = ss.getSheetByName("Pengaturan");
      let fileId = "";

      // Memecah format Base64 dan membuat Blob (File) di Google Drive Root Folder
      if (payload.base64Image && payload.base64Image.includes("data:image")) {
        const splitBase = payload.base64Image.split(",");
        const type = splitBase[0].split(";")[0].replace("data:", "");
        const byteCharacters = Utilities.base64Decode(splitBase[1]);
        const blob = Utilities.newBlob(
          byteCharacters,
          type,
          "Template_Sertifikat_Master",
        );

        const folder = DriveApp.getRootFolder();
        const file = folder.createFile(blob);
        file.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW,
        );
        fileId = file.getId();
      }

      // Bersihkan pengaturan lama
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
      }

      // Simpan konfigurasi baru ke dalam sel Spreadsheet
      sheet
        .getRange(2, 1, 1, 6)
        .setValues([
          [
            "DESAIN_MASTER",
            fileId,
            payload.xPos,
            payload.yPos,
            payload.fontSize,
            payload.color,
          ],
        ]);

      return buatResponJSON({
        status: "success",
        message: "Desain tersimpan di Cloud Drive",
      });
    }

    return buatResponJSON({
      status: "error",
      message: "Aksi Post tidak ditemukan",
    });
  } catch (err) {
    return buatResponJSON({ status: "error", message: err.toString() });
  }
}

function buatResponJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
