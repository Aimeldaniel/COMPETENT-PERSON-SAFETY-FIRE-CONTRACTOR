// ==========================================
// KONFIGURASI SISTEM
// ==========================================
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const FOLDER_ATTACHMENTS_ID = "1m08sUQ3DM9Z6689AYyrdfMEF7HRWHI2O"; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistem Pendaftaran Competent Person Fire Safety Contractor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// PENGURUSAN PENGGUNA & AUTHENTICATION
// ==========================================
function loginUser(email, password) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === email && data[i][4] === password) {
      if (data[i][6] !== "Aktif") {
        return { success: false, message: "Akaun anda telah dinyahaktifkan." };
      }
      return {
        success: true,
        user: {
          userId: data[i][0],
          name: data[i][1],
          ic: data[i][2],
          email: data[i][3],
          role: data[i][5]
        }
      };
    }
  }
  return { success: false, message: "Emel atau Kata Laluan tidak sah!" };
}

function registerUser(userData) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === userData.email) {
      return { success: false, message: "Emel telah berdaftar dalam sistem!" };
    }
  }
  
  const newId = "USR-" + Math.floor(100000 + Math.random() * 900000);
  const role = userData.role || "awam"; // default role awam
  
  sheet.appendRow([
    newId,
    userData.fullName,
    userData.icNo,
    userData.email,
    userData.password,
    role,
    "Aktif",
    new Date()
  ]);
  
  return { success: true, message: "Pendaftaran berjaya! Sila log masuk." };
}

// ==========================================
// SIMPAN & HANTAR PERMOHONAN
// ==========================================
function submitApplication(payload) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Applications");
    const folder = DriveApp.getFolderById(FOLDER_ATTACHMENTS_ID);
    
    // Upload Fail IC
    let icUrl = payload.existingIcUrl || "";
    if (payload.icFileBase64) {
      icUrl = uploadBase64ToDrive(folder, payload.icFileBase64, "IC_" + payload.icNo);
    }

    // Upload Sijil Latihan Kompetensi
    let cert1Url = payload.existingCert1Url || "";
    if (payload.cert1Base64) {
      cert1Url = uploadBase64ToDrive(folder, payload.cert1Base64, "CertComp1_" + payload.icNo);
    }
    let cert2Url = payload.existingCert2Url || "";
    if (payload.cert2Base64) {
      cert2Url = uploadBase64ToDrive(folder, payload.cert2Base64, "CertComp2_" + payload.icNo);
    }

    // Upload Surat Pengesahan Majikan
    let employerLetterUrl = payload.existingEmployerLetterUrl || "";
    if (payload.employerLetterBase64) {
      employerLetterUrl = uploadBase64ToDrive(folder, payload.employerLetterBase64, "EmployerLetter_" + payload.icNo);
    }

    const appId = payload.appId || ("APP-" + new Date().getFullYear() + "-" + Math.floor(1000 + Math.random() * 9000));
    const status = payload.isDraft ? "Draf" : "Dalam Proses";
    const now = new Date();

    const rowData = [
      appId,
      payload.email,
      payload.fullName,
      payload.icNo,
      icUrl,
      payload.gender,
      payload.dob,
      payload.address,
      payload.phone,
      JSON.stringify(payload.categoryData),
      JSON.stringify(payload.academicData),
      cert1Url,
      cert2Url,
      JSON.stringify(payload.workExpData),
      employerLetterUrl,
      status,
      "", "", "", "", "", "", "", "", // peg1, peg2, payment, cert info
      now,
      now
    ];

    // Semak jika update draft
    const data = sheet.getDataRange().getValues();
    let updated = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === appId) {
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        updated = true;
        break;
      }
    }

    if (!updated) {
      sheet.appendRow(rowData);
    }

    return { success: true, appId: appId, status: status, message: payload.isDraft ? "Draf disimpan." : "Permohonan berjaya dihantar!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// Helper untuk Upload Base64 ke Google Drive
function uploadBase64ToDrive(folder, base64Data, filename) {
  const splitData = base64Data.split(',');
  const contentType = splitData[0].split(':')[1].split(';')[0];
  const bytes = Utilities.base64Decode(splitData[1]);
  const blob = Utilities.newBlob(bytes, contentType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ==========================================
// ALUR KELULUSAN & SEMAKAN STATUS
// ==========================================
function getApplicationsForRole(userEmail, role) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Applications");
  const data = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const app = {
      appId: row[0],
      userEmail: row[1],
      fullName: row[2],
      icNo: row[3],
      icUrl: row[4],
      gender: row[5],
      dob: row[6],
      address: row[7],
      phone: row[8],
      categoryData: parseJSON(row[9]),
      academicData: parseJSON(row[10]),
      cert1Url: row[11],
      cert2Url: row[12],
      workExpData: parseJSON(row[13]),
      employerLetterUrl: row[14],
      status: row[15],
      peg1Review: row[16],
      peg1Date: row[17],
      peg2Review: row[18],
      peg2Date: row[19],
      paymentReceiptUrl: row[20],
      paymentDate: row[21],
      certNumber: row[22],
      certUrl: row[23],
      createdAt: row[24],
      updatedAt: row[25]
    };

    if (role === "awam" && app.userEmail === userEmail) {
      results.push(app);
    } else if (role === "pegawai_1" && ["Dalam Proses", "Dalam Pembayaran"].includes(app.status)) {
      results.push(app);
    } else if (role === "pegawai_2" && ["Kelulusan Pertama", "Dalam Pembayaran"].includes(app.status)) {
      results.push(app);
    } else if (role === "admin") {
      results.push(app);
    }
  }
  return results;
}

function processApproval(appId, newStatus, reviewText, role) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Applications");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === appId) {
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, 16).setValue(newStatus); // Status
      sheet.getRange(rowIndex, 26).setValue(new Date()); // Updated_At

      if (role === "pegawai_1") {
        sheet.getRange(rowIndex, 17).setValue(reviewText);
        sheet.getRange(rowIndex, 18).setValue(new Date());
      } else if (role === "pegawai_2") {
        sheet.getRange(rowIndex, 19).setValue(reviewText);
        sheet.getRange(rowIndex, 20).setValue(new Date());
      }

      // Notifikasi Emel
      const applicantEmail = data[i][1];
      const applicantName = data[i][2];
      sendNotificationEmail(applicantEmail, applicantName, appId, newStatus);

      // Jika Lulus Sepenuhnya, Janakan Sijil
      if (newStatus === "Lulus Sepenuhnya") {
        const certNo = "CERT-BOMBA-" + Math.floor(100000 + Math.random() * 900000);
        sheet.getRange(rowIndex, 23).setValue(certNo);
        sheet.getRange(rowIndex, 24).setValue("https://drive.google.com/sample_cert_" + appId); // Pautan Sijil
      }

      return { success: true, message: "Status permohonan " + appId + " berjaya dikemaskini ke: " + newStatus };
    }
  }
  return { success: false, message: "Permohonan tidak dijumpai." };
}

function submitPaymentReceipt(appId, receiptBase64) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Applications");
  const folder = DriveApp.getFolderById(FOLDER_ATTACHMENTS_ID);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === appId) {
      const receiptUrl = uploadBase64ToDrive(folder, receiptBase64, "Resit_" + appId);
      const rowIndex = i + 1;
      
      sheet.getRange(rowIndex, 16).setValue("Dalam Pembayaran");
      sheet.getRange(rowIndex, 21).setValue(receiptUrl);
      sheet.getRange(rowIndex, 22).setValue(new Date());
      sheet.getRange(rowIndex, 26).setValue(new Date());

      return { success: true, message: "Bukti pembayaran berjaya dimuat naik dan sedang disemak." };
    }
  }
  return { success: false, message: "Permohonan tidak wujud." };
}

function sendNotificationEmail(email, name, appId, status) {
  const subject = `[Status Permohonan Competent Person] - ${appId}: ${status}`;
  const body = `Salam Sejahtera ${name},\n\nStatus permohonan anda (${appId}) telah dikemaskini kepada: ${status}.\n\nSila log masuk ke dalam sistem untuk maklumat lanjut.\n\nSekian, terima kasih.`;
  MailApp.sendEmail(email, subject, body);
}

function parseJSON(str) {
  try { return JSON.parse(str); } catch(e) { return {}; }
}
