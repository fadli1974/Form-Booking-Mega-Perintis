// Jalankan fungsi ini SATU KALI dengan menekan tombol "Jalankan" (Run) di editor
// Ini bertujuan untuk memancing Google meminta Izin (Otorisasi) pengiriman Email.
function OtorisasiSistem() {
  var email = Session.getActiveUser().getEmail();
  MailApp.sendEmail(email, "Otorisasi Sukses", "Sistem LJR Booking sudah siap mengirim email.");
}

function doGet(e) {
  try {
    // If getting settings
    if (e.parameter && e.parameter.action === 'get_settings') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var settingSheet = ss.getSheetByName("Setting Kedatangan");
      
      var dataObj = { "tglAwal": "", "tglAkhir": "", "jamOpen": "", "jamClose": "", "quota": "" };
      
      if (settingSheet) {
        var lastRow = settingSheet.getLastRow();
        if (lastRow > 1) {
          var dataRow = settingSheet.getRange(lastRow, 1, 1, 5).getDisplayValues()[0];
          dataObj = {
            "tglAwal": dataRow[0] || "",
            "tglAkhir": dataRow[1] || "",
            "jamOpen": dataRow[2] || "",
            "jamClose": dataRow[3] || "",
            "quota": dataRow[4] || ""
          };
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "sukses",
        "data": dataObj
      })).setMimeType(ContentService.MimeType.JSON);
    } else if (e.parameter && e.parameter.action === 'get_users') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var userSheet = ss.getSheetByName("User");
      var userData = [];
      
      if (userSheet) {
        var data = userSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] || data[i][1]) { // If ID or Nama exists
            userData.push({
              rowIndex: i + 1,
              idUser: data[i][0],
              nama: data[i][1],
              email: data[i][2],
              jabatan: data[i][3],
              password: data[i][4],
              role: data[i][5]
            });
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        "status": "sukses",
        "data": userData
      })).setMimeType(ContentService.MimeType.JSON);
    } else if (e.parameter && e.parameter.action === 'get_pemalang') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Pemalang");
      var data = [];
      if (sheet) {
        var rawData = sheet.getDataRange().getDisplayValues();
        for (var i = 1; i < rawData.length; i++) {
          if (rawData[i][4] || rawData[i][2]) { // Check Nama Perusahaan or Timestamp
            data.push({
              rowIndex: i + 1,
              actualKedatangan: rawData[i][0],
              tro: rawData[i][1],
              timestamp: rawData[i][2],
              tanggalKedatangan: rawData[i][3],
              namaPerusahaan: rawData[i][4],
              jumlahToko: rawData[i][5],
              jumlahBlister: rawData[i][6],
              jumlahQty: rawData[i][7],
              tujuanPengiriman: rawData[i][8],
              pic: rawData[i][9]
            });
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        "status": "sukses",
        "data": data
      })).setMimeType(ContentService.MimeType.JSON);
    } else if (e.parameter && e.parameter.action === 'get_all_data') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var responseData = { booking: [], pemalang: [], users: [], settings: {} };

      // 1. Settings
      var settingSheet = ss.getSheetByName("Setting Kedatangan");
      if (settingSheet && settingSheet.getLastRow() > 1) {
        var sRow = settingSheet.getRange(settingSheet.getLastRow(), 1, 1, 5).getDisplayValues()[0];
        responseData.settings = { tglAwal: sRow[0], tglAkhir: sRow[1], jamOpen: sRow[2], jamClose: sRow[3], quota: sRow[4] };
      }

      // 2. Booking Data (Data Booking sheet is used when action is empty usually, let's grab it)
      var dataSheet = ss.getSheetByName("Data Booking") || ss.getSheets()[0];
      if (dataSheet) {
        var bData = dataSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < bData.length; i++) {
          if (bData[i][2] || bData[i][5]) {
            responseData.booking.push({
              rowIndex: i + 1,
              tro: bData[i][0], 
              actualKedatangan: bData[i][1],
              tanggalKedatangan: bData[i][3], 
              namaPerusahaan: bData[i][5], 
              brands: bData[i][6], 
              style: bData[i][7], 
              noShipment: bData[i][8], 
              jumlahToko: bData[i][10],
              jumlahBlister: bData[i][11], 
              jumlahQty: bData[i][12], 
              tujuanPengiriman: bData[i][13], 
              aksi: bData[i][17]
            });
          }
        }
      }

      // 3. Pemalang
      var pSheet = ss.getSheetByName("Pemalang");
      if (pSheet) {
        var pData = pSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < pData.length; i++) {
          if (pData[i][4] || pData[i][2]) {
            responseData.pemalang.push({
              rowIndex: i + 1, actualKedatangan: pData[i][0], tro: pData[i][1], timestamp: pData[i][2],
              tanggalKedatangan: pData[i][3], namaPerusahaan: pData[i][4], jumlahToko: pData[i][5],
              jumlahBlister: pData[i][6], jumlahQty: pData[i][7], tujuanPengiriman: pData[i][8], pic: pData[i][9]
            });
          }
        }
      }

      // 4. Users
      var uSheet = ss.getSheetByName("User");
      if (uSheet) {
        var uData = uSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < uData.length; i++) {
          if (uData[i][0] || uData[i][1]) {
            responseData.users.push({
              rowIndex: i + 1, idUser: uData[i][0], nama: uData[i][1],
              email: uData[i][2], jabatan: uData[i][3], password: uData[i][4], role: uData[i][5]
            });
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "sukses", data: responseData })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getDisplayValues(); // Use getDisplayValues to get formatted dates/strings
    
    var result = [];
    // Start from row 2 (index 1) to skip headers
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Only push rows that actually have data (e.g. check Nama Perusahaan or Timestamp)
      if (row[2] || row[5]) { 
        result.push({
          rowIndex: i + 1, // Row number in spreadsheet
          tro: row[0], // Kolom A
          actualKedatangan: row[1], // Kolom B
          tanggalKedatangan: row[3], // Kolom D
          namaPerusahaan: row[5], // Kolom F
          brands: row[6], // Kolom G
          style: row[7], // Kolom H
          noShipment: row[8], // Kolom I
          jumlahToko: row[10], // Kolom K
          jumlahBlister: row[11], // Kolom L
          jumlahQty: row[12], // Kolom M
          tujuanPengiriman: row[13], // Kolom N
          aksi: row[17] // Kolom R
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ "status": "sukses", "data": result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Check if this is an update action
    if (data.action === "update_actual") {
      sheet.getRange(data.rowIndex, 2).setValue(data.actualKedatangan); // Update Kolom B
      sheet.getRange(data.rowIndex, 18).setValue("Received"); // Update Kolom R
      return ContentService.createTextOutput(JSON.stringify({ "status": "sukses", "message": "Berhasil diupdate" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "update_tro") {
      sheet.getRange(data.rowIndex, 1).setValue(data.tro); // Update Kolom A
      return ContentService.createTextOutput(JSON.stringify({ "status": "sukses", "message": "TRO Berhasil diupdate" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "save_settings") {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var settingSheet = ss.getSheetByName("Setting Kedatangan");
      
      // Buat tab jika belum ada
      if (!settingSheet) {
        settingSheet = ss.insertSheet("Setting Kedatangan");
      }
      
      // Buat header jika kosong
      if (settingSheet.getLastRow() === 0) {
        var headers = [["Tanggal Awal", "Tanggal Akhir", "Jam Open", "Jam Close", "Quota"]];
        settingSheet.getRange(1, 1, 1, 5).setValues(headers);
        settingSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9ead3"); // Style dikit
      }
      
      // Simpan data di baris kedua
      var rowData = [[data.tglAwal, data.tglAkhir, data.jamOpen, data.jamClose, data.quota]];
      settingSheet.getRange(2, 1, 1, 5).setValues(rowData);
      
      return ContentService.createTextOutput(JSON.stringify({ "status": "sukses", "message": "Settings saved to spreadsheet" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "add_user") {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var userSheet = ss.getSheetByName("User");
      
      // Buat tab jika belum ada
      if (!userSheet) {
        userSheet = ss.insertSheet("User");
      }
      
      // Buat header jika kosong
      if (userSheet.getLastRow() === 0) {
        var headers = [["ID User", "Nama", "Email", "Jabatan", "Password", "Role"]];
        userSheet.getRange(1, 1, 1, 6).setValues(headers);
        userSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#cfe2f3"); 
      }
      
      // Buat ID User (Tahun-BulanTanggal-Nomor)
      var now = new Date();
      var prefix = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
      // Append ke baris terakhir
      userSheet.appendRow([prefix, data.nama, data.email, data.jabatan, data.password, data.role]);
      
      return ContentService.createTextOutput(JSON.stringify({ "status": "sukses", "message": "User berhasil ditambahkan" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "login") {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var userSheet = ss.getSheetByName("User");
      if (userSheet) {
        var users = userSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < users.length; i++) {
          // users[i][1] is Nama (Username), users[i][4] is Password
          var sheetNama = (users[i][1] || "").toString().toLowerCase();
          var inputNama = (data.username || "").toString().toLowerCase();
          
          if (sheetNama === inputNama && sheetNama !== "" && users[i][4] === data.password) {
            return ContentService.createTextOutput(JSON.stringify({ 
              "status": "sukses", 
              "data": { "idUser": users[i][0], "nama": users[i][1], "email": users[i][2], "jabatan": users[i][3], "role": users[i][5] }
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Username atau password salah" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "input_manual_pemalang") {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Pemalang");
      if (!sheet) {
        sheet = ss.insertSheet("Pemalang");
        sheet.appendRow(["Actual Kedatangan", "TRO", "Timestamp", "Tanggal Kedatangan", "Nama Perusahaan ( PT atau CV )", "Jumlah Toko", "Jumlah Blister", "Jumlah Qty", "Tujuan Pengiriman", "PIC"]);
      }
      sheet.appendRow([
        data.actualKedatangan, data.tro, new Date(), data.tanggalKedatangan, data.namaPerusahaan, 
        data.jumlahToko, data.jumlahBlister, data.jumlahQty, data.tujuanPengiriman, data.pic
      ]);
      return ContentService.createTextOutput(JSON.stringify({ "status": "sukses" })).setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "upload_pemalang") {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Pemalang");
      if (!sheet) {
        sheet = ss.insertSheet("Pemalang");
        sheet.appendRow(["Actual Kedatangan", "TRO", "Timestamp", "Tanggal Kedatangan", "Nama Perusahaan ( PT atau CV )", "Jumlah Toko", "Jumlah Blister", "Jumlah Qty", "Tujuan Pengiriman", "PIC"]);
      }
      if (data.rows && data.rows.length > 0) {
        // Skip header if it exists in the uploaded rows
        var rowsToAppend = [];
        for (var r = 0; r < data.rows.length; r++) {
            if (r === 0 && (data.rows[r][1] === "TRO" || data.rows[r][4] && data.rows[r][4].includes("Nama Perusahaan"))) {
                continue; // Skip header row
            }
            rowsToAppend.push(data.rows[r]);
        }
        if (rowsToAppend.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ "status": "sukses", "message": "Data berhasil diupload" })).setMimeType(ContentService.MimeType.JSON);
    } else if (data.action === "edit_pemalang") {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pemalang");
      if (sheet) {
        // Assume rowIndex is passed (and it corresponds to sheet row directly or we need +1 for header)
        // Usually frontend sends rowIndex. We will assume rowIndex is 1-based, or frontend sends the actual row index of the sheet
        var actualRow = parseInt(data.rowIndex);
        sheet.getRange(actualRow, 1, 1, 10).setValues([[
          data.actualKedatangan, data.tro, new Date(), data.tanggalKedatangan, data.namaPerusahaan, 
          data.jumlahToko, data.jumlahBlister, data.jumlahQty, data.tujuanPengiriman, data.pic
        ]]);
        return ContentService.createTextOutput(JSON.stringify({ "status": "sukses" })).setMimeType(ContentService.MimeType.JSON);
      }
    } else if (data.action === "delete_pemalang") {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pemalang");
      if (sheet) {
        var actualRow = parseInt(data.rowIndex);
        sheet.deleteRow(actualRow);
        return ContentService.createTextOutput(JSON.stringify({ "status": "sukses" })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // Siapkan folder untuk menampung file upload di Google Drive (buat otomatis jika belum ada)
    var folderName = "Lampiran Booking Mega Perintis";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    // Fungsi kecil untuk upload file base64 ke Google Drive dan mereturn URL-nya
    var uploadFile = function(fileObj) {
      if (fileObj && fileObj.data) {
        var blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, fileObj.name);
        var file = folder.createFile(blob);
        return file.getUrl();
      }
      return "";
    };

    // Proses upload lampiran (jika ada)
    var urlPackingList = uploadFile(data.filePackingList);
    var urlInspection = uploadFile(data.fileInspection);
    var urlKedatangan = uploadFile(data.fileKedatangan);
    
    // Format Waktu Pengisian (Waktu Indonesia Barat / GMT+7)
    var timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    // Susun data sesuai dengan urutan kolom
    var rowData = [
      data.tro2 || "",                  // A: TRO 2
      data.actualKedatangan || "",      // B: Actual Kedatangan
      timestamp,                        // C: Timestamp
      data.tanggalKedatangan || "",     // D: Tanggal Kedatangan
      data.email || "",                 // E: Alamat email
      data.namaPerusahaan || "",        // F: Nama Perusahaan
      data.brands || "",                // G: Brands
      data.style || "",                 // H: Style
      data.noShipment || "",            // I: No Shipment Confirmation
      data.noTlp ? "'" + data.noTlp : "",// J: No Tlp / Hp (ditambah kutip agar angka 0 tidak hilang di excel)
      data.jumlahToko || "",            // K: Jumlah Toko
      data.jumlahBlister || "",         // L: Jumlah Blister
      data.jumlahQty || "",             // M: Jumlah Qty
      data.tujuanPengiriman || "",      // N: Tujuan Pengiriman
      urlPackingList,                   // O: Upload Excell
      urlInspection,                    // P: INSPECTION REPORT
      urlKedatangan                     // Q: Form kedatangan
    ];
    
    // Masukkan data tersebut ke baris baru paling bawah di Spreadsheet
    sheet.appendRow(rowData);
    
    // ==========================================
    // BAGIAN PENGIRIMAN EMAIL DAN PEMBUATAN PDF
    // ==========================================
    var lastRow = sheet.getLastRow();
    var nomer = "00001";
    
    // Cek nomor terakhir di kolom Q (kolom 17) pada baris sebelumnya (karena baris saat ini baru saja di-append)
    if (lastRow > 2) {
      var lastVal = sheet.getRange(lastRow - 1, 17).getValue(); // Baris sebelum data yang baru di-append
      var lastNum = parseInt(lastVal, 10);
      if (!isNaN(lastNum)) {
        nomer = (lastNum + 1).toString().padStart(5, '0');
      } else {
        // Jika teks biasa, gunakan nomor baris saja
        nomer = (lastRow - 1).toString().padStart(5, '0');
      }
    }
    
    // Update data baris yang baru di-append dengan nomer pendaftaran di Kolom Q
    sheet.getRange(lastRow, 17).setValue("'" + nomer);

    var namaVendor = data.namaPerusahaan || "-";
    
    // 1. Buat desain PDF dalam bentuk HTML
    var htmlContent = `
    <div style="font-family: Arial, sans-serif; color: #000; padding: 20px;">
        
        <div style="text-align: center; border-bottom: 2px solid #1d4ed8; padding-bottom: 10px; margin-bottom: 30px;">
            <h1 style="color: #1e40af; font-style: italic; margin:0;">LJR LOGISTICS</h1>
            <p style="margin:0; font-size:12px; color: #1d4ed8; font-weight:bold;">Domestic Cargo Service</p>
        </div>
        
        <div style="border: 2px solid #94a3b8; border-radius: 15px; width: 300px; margin: 0 auto; text-align: center; padding: 20px;">
            <h1 style="font-size: 55px; margin: 0; font-weight: bold; color: #0f172a;">${nomer}</h1>
            <p style="font-weight: bold; margin: 10px 0 0 0; font-size: 16px; color: #334155;">NOMER PENDAFTARAN</p>
        </div>

        <table style="margin-top: 50px; width: 100%; font-size: 14px; font-weight: bold;" cellpadding="8">
            <tr><td width="200">Tanggal ke datangan</td><td width="10">:</td><td>${data.tanggalKedatangan || '-'}</td></tr>
            <tr><td>Nama Vendor/suplier</td><td>:</td><td>${namaVendor}</td></tr>
            <tr><td>Brands</td><td>:</td><td>${data.brands || '-'}</td></tr>
            <tr><td>Style</td><td>:</td><td>${data.style || '-'}</td></tr>
            <tr><td>Jumlah Toko</td><td>:</td><td>${data.jumlahToko || '-'}</td></tr>
            <tr><td>Jumlah Blister</td><td>:</td><td>${data.jumlahBlister || '-'}</td></tr>
            <tr><td>Jumlah QTY</td><td>:</td><td>${data.jumlahQty || '-'}</td></tr>
            <tr><td>No. Handphone</td><td>:</td><td>${data.noTlp || '-'}</td></tr>
            <tr><td>Shipment Confirmation</td><td>:</td><td>${data.noShipment || '-'}</td></tr>
        </table>
    </div>
    `;
    
    // Convert HTML menjadi PDF file
    var pdfBlob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF);
    pdfBlob.setName(namaVendor + ".pdf");
    
    // 2. Isi Teks Email
    var emailBody = "Terimakasih Sudah Melakukan Pengisian Form Booking Melalui Aplikasi\n\n" +
      "1. Setiap supplier MP yang akan kirim barang ke LJR harus isi form kedatangan Maxsimal H-1\n" +
      "2. Link akan dibuka dari Jam 8:00 s/d jam 17:00 otomatis terkunci sendiri ( senin s/d sabtu )\n" +
      "3. Penerimaan barang di WH LJR dari hari senin s/d Jumat dari jam 08:00 s/d 15:00 apabila kedatangan lebih dari hari / jam yang di tentukan maka kami perhak menolak ( dibongkar di hari berikutnya )\n" +
      "4. Setiap hari sabtu kami tidak terima barang masuk dan bongkar barang.\n" +
      "5. Setiap Supplier / CMT Mp yang akan kirim barang di haruskan berpakaian sopan ( memakai celana panjang & Sepatu ).\n\n" +
      "LJR Logistics\n\n" +
      "Fadly ( 0817-159-258 )";

    // 3. Pengaturan Penerima & Kirim Email
    var subject = "No.Pendaftaran " + nomer + " " + namaVendor;
    var recipient = data.email;
    var ccEmails = "rivaldi.fahreza@lestarijayaraya.com, fadli.alamsyah@lestarijayaraya.com, desy.ekowati@megaperintis.co.id";
    
    if (recipient) {
      MailApp.sendEmail({
        to: recipient,
        cc: ccEmails,
        subject: subject,
        body: emailBody,
        attachments: [pdfBlob]
      });
    }

    // Berikan respons sukses kembali ke website HTML dengan nomor pendaftaran yang di-generate
    return ContentService.createTextOutput(JSON.stringify({ "status": "sukses", "nomerPendaftaran": nomer }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
