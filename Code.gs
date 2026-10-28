// SPREADSHEET IDs
const MASTER_DROPDOWN_SS_ID = "1KOJamWs_pc15iqdGXyA4vR-NTeuKLXHWKZDjerbJPbo"; // "Wil Lingk K7"
const FORM_RESPON_SS_ID     = "1P9hJCC_kVG0JgLwtBRhOHVOgU3yT6ImtMS9Z5kV-rL0"; // Sheet respon "penyalalilin"
const JADWAL_SS_ID          = "1nU48nL0W7ehkle7FZIeEYZaYe7kfUlVgb3QFZYqiexc"; // Spreadsheet penyalalilinjadwal.xlsx ("jadwalpenyala")

function doGet(e) {
  const action = e ? e.parameter.action : '';
  const callback = e ? e.parameter.callback : '';

  // 1. ENDPOINT JSONP: Dropdown Data (Form Pendaftaran)
  if (action === 'getDropdownData') {
    const ss = SpreadsheetApp.openById(MASTER_DROPDOWN_SS_ID);
    const sheet = ss.getSheetByName("wildanlingkKategorial");
    
    if (!sheet) {
      return sendJsonpResponse(callback, { status: "error", message: "Sheet wildanlingkKategorial tidak ditemukan" });
    }

    const values = sheet.getDataRange().getValues();
    const rows = values.slice(1);

    const wilayahMap = {};
    const kategorialList = [];

    rows.forEach(function(row) {
      const wilayah = row[0] ? String(row[0]).trim() : "";
      
      if (wilayah) {
        const lingkungans = [];
        for (let i = 1; i <= 4; i++) {
          if (row[i] && String(row[i]).trim() !== "") {
            lingkungans.push(String(row[i]).trim());
          }
        }
        wilayahMap[wilayah] = lingkungans;
      }

      if (row[8] && String(row[8]).trim() !== "") {
        kategorialList.push(String(row[8]).trim());
      }
    });

    return sendJsonpResponse(callback, {
      status: "success",
      wilayahMap: wilayahMap,
      kategorialList: kategorialList
    });
  }

  // 2. ENDPOINT JSONP: Report Data (report.html)
  if (action === 'getReportData') {
    return getReportDataResponse(callback);
  }

  // Fallback Web App Status
  return ContentService.createTextOutput("Service Active");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);

  if (!hasLock) {
    return ContentService.createTextOutput(
      JSON.stringify({
        result: 'busy',
        message: 'Server sedang sibuk karena banyaknya pendaftaran bersamaan. Silakan coba klik Kirim Pendaftaran beberapa saat lagi.'
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    // Bot honeypot check
    if (e.parameter.mobile_number && e.parameter.mobile_number !== '') {
      return ContentService.createTextOutput(
        JSON.stringify({ result: 'success', message: 'Bot detected' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const params = e.parameter;
    const ss = SpreadsheetApp.openById(FORM_RESPON_SS_ID);
    let sheet = ss.getSheetByName("penyalalilin");

    if (!sheet) {
      sheet = ss.insertSheet("penyalalilin");
      sheet.appendRow([
        "id", "timestamp", "email", "Wilayah", "Lingkungan", 
        "Pengurus atau Kategorial lain", "Nama Suami", "No. HP/WA Suami", 
        "Nama Istri", "No. HP/WA Istri"
      ]);
    }

    const jalur = params.jalur; // 'wilayah' atau 'kategorial'
    const wilayah = jalur === 'wilayah' ? (params.wilayah || '') : '';
    const lingkungan = jalur === 'wilayah' ? (params.lingkungan || '') : '';
    const kategorial = jalur === 'kategorial' ? (params.kategorial || '') : '';

    const namaSuami = params.nama_suami || '';
    const hpSuami = params.hp_suami ? "'" + params.hp_suami : '';
    const namaIstri = params.nama_istri || '';
    const hpIstri = params.hp_istri ? "'" + params.hp_istri : '';
    const email = params.email || '';

    // Cek Duplikasi di Sheet "penyalalilin"
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const existingWilayah = String(data[i][3]).trim();
      const existingLingkungan = String(data[i][4]).trim();
      const existingKategorial = String(data[i][5]).trim();
      const existingSuami = data[i][6];
      const existingIstri = data[i][8];

      if (jalur === 'wilayah' && existingWilayah === wilayah) {
        return ContentService.createTextOutput(
          JSON.stringify({
            result: 'duplicate',
            message: `Maaf, Pendaftaran tidak dapat dilakukan, Sebab ${wilayah} sudah terdaftar atas nama ${existingSuami} dan ${existingIstri} dari lingkungan ${existingLingkungan}.`
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      if (jalur === 'kategorial' && existingKategorial === kategorial) {
        return ContentService.createTextOutput(
          JSON.stringify({
            result: 'duplicate',
            message: `Maaf, Pendaftaran tidak dapat dilakukan, Sebab ${kategorial} sudah terdaftar atas nama ${existingSuami} dan ${existingIstri}.`
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Simpan Data Baru
    const newId = Utilities.getUuid();
    const timestamp = new Date();

    sheet.appendRow([
      newId, timestamp, email, wilayah, lingkungan, 
      kategorial, namaSuami, hpSuami, namaIstri, hpIstri
    ]);

    // Kirim Email Konfirmasi
    if (email !== '') {
      const senderName = 'Panitia Natal 2026 - Wilayah Theresia dari Kanak-Kanak Yesus';
      const subject = 'Pendaftaran Petugas Penyalaan Lilin Adven 2026';
      
      const bodyDetail = `
Syalom Bapak/Ibu,

Terima kasih telah mendaftar sebagai Petugas Pasutri Penyala Lilin Adven 2026 Paroki Katedral Jakarta.

Berikut rincian data pendaftaran Anda:
--------------------------------------------------
ID Pendaftaran : ${newId}
Waktu          : ${Utilities.formatDate(timestamp, "GMT+7", "dd/MM/yyyy HH:mm:ss")} WIB
Mewakili       : ${jalur === 'wilayah' ? 'Wilayah & Lingkungan' : 'Pengurus / Kategorial Lain'}
${jalur === 'wilayah' ? `Wilayah        : ${wilayah}\nLingkungan     : ${lingkungan}` : `Pengurus/Kategorial: ${kategorial}`}
Nama Suami     : ${namaSuami} (${params.hp_suami || ''})
Nama Istri     : ${namaIstri} (${params.hp_istri || ''})
--------------------------------------------------

Jadwal penugasan akan diingatkan kembali melalui WhatsApp.

Tuhan memberkati,
Panitia Natal 2026 - Wil St. Theresia dari Kanak-Kanak Yesus
Paroki Katedral Jakarta
      `;

      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: bodyDetail,
          name: senderName
        });
      } catch (mailErr) {
        Logger.log("Gagal mengirim email: " + mailErr.toString());
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        result: 'success',
        details: {
          id: newId,
          jalur: jalur,
          wilayah: wilayah,
          lingkungan: lingkungan,
          kategorial: kategorial,
          namaSuami: namaSuami,
          hpSuami: hpSuami,
          namaIstri: namaIstri,
          hpIstri: hpIstri,
          email: email
        }
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'error', message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// FUNGSI OLAH REPORT DATA UNTUK REPORT.HTML
function getReportDataResponse(callback) {
  try {
    // 1. Ambil data Jadwal dari Spreadsheet "penyalalilinjadwal"
    const ssJadwal = SpreadsheetApp.openById(JADWAL_SS_ID);
    const sheetJadwal = ssJadwal.getSheetByName("jadwalpenyala");
    const jadwalValues = sheetJadwal.getDataRange().getValues();
    const jadwalRows = jadwalValues.slice(1);

    // 2. Ambil data Respon dari Sheet "penyalalilin"
    let responRows = [];
    try {
      const ssRespon = SpreadsheetApp.openById(FORM_RESPON_SS_ID);
      const sheetRespon = ssRespon.getSheetByName("penyalalilin");
      if (sheetRespon) {
        responRows = sheetRespon.getDataRange().getValues().slice(1);
      }
    } catch (err) {
      Logger.log("Gagal membuka sheet respon: " + err.toString());
    }

    // 3. Pencocokan Data Respon ke Jadwal
    const reportList = jadwalRows.map(row => {
      const no = row[0];
      const tanggal = row[1] instanceof Date ? Utilities.formatDate(row[1], "GMT+7", "yyyy-MM-dd") : row[1];
      
      // PERBAIKAN FORMAT JAM (Mengabaikan Offset Historis Timezone 1899)
      let jam = "";
      if (row[2] instanceof Date) {
        const h = String(row[2].getHours()).padStart(2, '0');
        const m = String(row[2].getMinutes()).padStart(2, '0');
        jam = `${h}:${m}`;
      } else {
        jam = String(row[2] || '');
      }

      const eventName = row[3];
      const asalPetugas = String(row[4] || '').trim();

      let namaSuami = "-";
      let namaIstri = "-";

      for (let i = 0; i < responRows.length; i++) {
        const respWilayah = String(responRows[i][3] || '').trim();
        const respKategorial = String(responRows[i][5] || '').trim();
        const respSuami = responRows[i][6] || '-';
        const respIstri = responRows[i][8] || '-';

        let isMatch = false;

        if (respWilayah !== "") {
          if (respWilayah.toLowerCase() === asalPetugas.toLowerCase()) {
            isMatch = true;
          }
        } else {
          if (respKategorial !== "") {
            if (respKategorial.toLowerCase() === asalPetugas.toLowerCase()) {
              isMatch = true;
            }
          }
        }

        if (isMatch) {
          namaSuami = respSuami;
          namaIstri = respIstri;
          break;
        }
      }

      return {
        no: no,
        tanggal: tanggal,
        jam: jam,
        event: eventName,
        asalPetugas: asalPetugas,
        namaSuami: namaSuami,
        namaIstri: namaIstri
      };
    });

    return sendJsonpResponse(callback, { status: "success", data: reportList });

  } catch (error) {
    return sendJsonpResponse(callback, { status: "error", message: error.toString() });
  }
}

function sendJsonpResponse(callback, data) {
  const jsonOutput = JSON.stringify(data);
  const output = callback ? `${callback}(${jsonOutput})` : jsonOutput;
  return ContentService.createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
