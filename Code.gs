const SPREADSHEET_ID = "1KOJamWs_pc15iqdGXyA4vR-NTeuKLXHWKZDjerbJPbo"; // Gunakan Spreadsheet reference "Wil Lingk K7"


function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;

  // Endpoint JSONP untuk mengambil data dropdown
  if (action === 'getDropdownData') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("wildanlingkKategorial");
    
    if (!sheet) {
      return sendJsonpResponse(callback, { status: "error", message: "Sheet wildanlingkKategorial tidak ditemukan" });
    }

    const values = sheet.getDataRange().getValues();
    // Menggunakan slice(1) untuk mengabaikan baris header
    const rows = values.slice(1);

    const wilayahMap = {};
    const kategorialList = [];

    rows.forEach(function(row) {
      const wilayah = row[0] ? String(row[0]).trim() : "";
      
      // Ambil Lingkungan (Kolom B sampai E / Index 1-4)
      if (wilayah) {
        const lingkungans = [];
        for (let i = 1; i <= 4; i++) {
          if (row[i] && String(row[i]).trim() !== "") {
            lingkungans.push(String(row[i]).trim());
          }
        }
        wilayahMap[wilayah] = lingkungans;
      }

      // Ambil Pengurus / Kategorial Lain (Kolom I / Index 8)
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

  // Jika tidak ada action, tampilkan halaman jika di-deploy sebagai Web App
  return ContentService.createTextOutput("Service Active");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);

  // Jika gagal mendapatkan kunci (antrian penuh / timeout)
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
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // Gunakan Spreadsheet Aktif
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
    const hpSuami = params.hp_suami || '';
    const namaIstri = params.nama_istri || '';
    const hpIstri = params.hp_istri || '';
    const email = params.email || '';

    // Cek Duplikasi di Sheet "penyalalilin"
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const existingWilayah = String(data[i][3]).trim();
      const existingLingkungan = String(data[i][4]).trim();
      const existingKategorial = String(data[i][5]).trim();
      const existingSuami = data[i][6];
      const existingIstri = data[i][8];

      // Jika memilih jalur Wilayah: cek apakah Wilayah & Lingkungan sudah terdaftar
      if (jalur === 'wilayah' && existingWilayah === wilayah && existingLingkungan === lingkungan) {
        return ContentService.createTextOutput(
          JSON.stringify({
            result: 'duplicate',
            message: `Maaf, Pendaftaran tidak dapat dilakukan, Sebab Lingkungan ${lingkungan} (${wilayah}) sudah terdaftar atas nama ${existingSuami} dan ${existingIstri}.`
          })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // Jika memilih jalur Kategorial: cek apakah Kategorial sudah terdaftar
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

    // Kirim Email Konfirmasi jika mengisi email
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
${jalur === 'wilayah' ? `Wilayah        : ${wilayah}\nLingkungan     :${lingkungan}` : `Pengurus/Kategorial: ${kategorial}`}
Nama Suami     : ${namaSuami} (${hpSuami})
Nama Istri     : ${namaIstri} (${hpIstri})
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

function sendJsonpResponse(callback, data) {
  const jsonOutput = JSON.stringify(data);
  const output = callback ? `${callback}(${jsonOutput})` : jsonOutput;
  return ContentService.createTextOutput(output)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}