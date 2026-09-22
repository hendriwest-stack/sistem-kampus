/****************************************************
 * GAS PRESENSI MAHASISWA
 * Versi 1.0
 ****************************************************/


/* ==================================================
   KONFIGURASI
================================================== */

// ID GOOGLE SPREADSHEET
const SPREADSHEET_ID = '143woF2y9cfqH90lze4K-JVlq_emqc_WGPE_sC9MumrM';

// NAMA SHEET
const SHEET_MAHASISWA = 'MAHASISWA';
const SHEET_DOSEN = 'DOSEN';
const SHEET_PRESENSI = 'PRESENSI';


/* ==================================================
   LOKASI KAMPUS
================================================== */

// GANTI DENGAN KOORDINAT KAMPUS DARI GOOGLE MAPS

const LATITUDE_KAMPUS = -7.565375;
const LONGITUDE_KAMPUS = 110.864881;

// Radius maksimal presensi dalam meter
const RADIUS_KAMPUS = 100;


/* ==================================================
   WEB APP
================================================== */

function doGet(e) {

  // Jika dipanggil dengan ?api=1
  if (e && e.parameter && e.parameter.api === '1') {
    return handleAPI(e);
  }

  // Jika dibuka biasa, tampilkan aplikasi web
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Presensi Mahasiswa')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


function handleAPI(e) {

  try {

    const action =
      String(e.parameter.action || '').trim();


    switch (action) {

      case 'serverTime':

        return jsonResponse({
          success: true,
          data: getServerTime()
        });


      case 'dosen':

        return jsonResponse({
          success: true,
          data: getDosen()
        });


      case 'mahasiswa':

        return jsonResponse(
          apiMahasiswa(e)
        );


      case 'cekPresensi':

        return jsonResponse(
          apiCekPresensi(e)
        );


      case 'cekLokasi':

        return jsonResponse(
          apiCekLokasi(e)
        );


      default:

        return jsonResponse({

          success: false,

          message:
            'Action API tidak dikenal.'

        });

    }


  } catch (error) {

    return jsonResponse({

      success: false,

      message:
        error.message

    });

  }

}

function jsonResponse(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}

function apiMahasiswa(e) {

  const nim =
    String(
      e.parameter.nim || ''
    ).trim();


  if (!nim) {

    return {

      success: false,

      message:
        'NIM wajib diisi.'

    };

  }


  const mahasiswa =
    getMahasiswa(nim);


  return {

    success: true,

    data: mahasiswa

  };

}

function apiCekPresensi(e) {

  const nim =
    String(
      e.parameter.nim || ''
    ).trim();


  if (!nim) {

    return {

      success: false,

      message:
        'NIM wajib diisi.'

    };

  }


  const hasil =
    cekPresensiHariIni(nim);


  return {

    success: true,

    data: hasil

  };

}

function apiCekLokasi(e) {

  const latitude =
    Number(
      e.parameter.latitude
    );


  const longitude =
    Number(
      e.parameter.longitude
    );


  if (
    isNaN(latitude) ||
    isNaN(longitude)
  ) {

    return {

      success: false,

      message:
        'Latitude dan longitude wajib diisi.'

    };

  }


  const hasil =
    cekLokasi(
      latitude,
      longitude
    );


  return {

    success: true,

    data: hasil

  };

}


function doPost(e) {

  try {

    let data;


    /*
     * JSON
     */

    if (
      e.postData &&
      e.postData.contents
    ) {

      data =
        JSON.parse(
          e.postData.contents
        );

    } else {

      data = e.parameter;

    }


    const action =
      String(
        data.action || ''
      ).trim();


    switch (action) {


      case 'simpanPresensi':

        return jsonResponse({

          success: true,

          data:
            simpanPresensi(data)

        });


      default:

        return jsonResponse({

          success: false,

          message:
            'Action POST tidak dikenal.'

        });

    }


  } catch (error) {

    return jsonResponse({

      success: false,

      message:
        error.message

    });

  }

}



/* ==================================================
   INCLUDE FILE
================================================== */

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


/* ==================================================
   WAKTU SERVER
================================================== */

function getServerTime() {

  const now = new Date();

  const timezone =
    Session.getScriptTimeZone();

  return {

    tanggal: Utilities.formatDate(
      now,
      timezone,
      'dd MMMM yyyy'
    ),

    waktu: Utilities.formatDate(
      now,
      timezone,
      'HH:mm:ss'
    )

  };

}


/* ==================================================
   DATA MAHASISWA
================================================== */

function getMahasiswa(nim) {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      SHEET_MAHASISWA
    );

  if (!sheet) {
    throw new Error(
      'Sheet MAHASISWA tidak ditemukan.'
    );
  }

  const data =
    sheet.getDataRange().getValues();

  nim =
    String(nim).trim();

  for (let i = 1; i < data.length; i++) {

    const nimSheet =
      String(data[i][0]).trim();

    if (nimSheet === nim) {

      return {

        ditemukan: true,

        nim: nimSheet,

        nama: String(data[i][1]),

        prodi: String(data[i][2]),

        status: String(data[i][3])

      };

    }

  }

  return {
    ditemukan: false
  };

}


/* ==================================================
   DATA DOSEN
================================================== */

function getDosen() {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      SHEET_DOSEN
    );

  if (!sheet) {

    throw new Error(
      "Sheet DOSEN tidak ditemukan."
    );

  }

  const data =
    sheet
      .getDataRange()
      .getDisplayValues();


  const hasil = [];


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const nidn =
      String(data[i][0] || "")
        .trim();

    const nama =
      String(data[i][1] || "")
        .trim();

    const status =
      String(data[i][2] || "")
        .trim()
        .toUpperCase();


    if (
      nidn !== "" &&
      nama !== "" &&
      status === "AKTIF"
    ) {

      hasil.push({

        nidn:
          nidn,

        nama:
          nama,

        status:
          status

      });

    }

  }


  return hasil;

}


/* ==================================================
   HITUNG JARAK GPS
   Haversine
================================================== */

function hitungJarak(
  lat1,
  lon1,
  lat2,
  lon2
) {

  const R = 6371000;

  const dLat =
    (lat2 - lat1) *
    Math.PI / 180;

  const dLon =
    (lon2 - lon1) *
    Math.PI / 180;

  const a =

    Math.sin(dLat / 2) *
    Math.sin(dLat / 2)

    +

    Math.cos(
      lat1 * Math.PI / 180
    )

    *

    Math.cos(
      lat2 * Math.PI / 180
    )

    *

    Math.sin(dLon / 2) *
    Math.sin(dLon / 2);

  const c =
    2 *
    Math.atan2(
      Math.sqrt(a),
      Math.sqrt(1 - a)
    );

  return R * c;

}


/* ==================================================
   CEK LOKASI
================================================== */

function cekLokasi(
  latitude,
  longitude
) {

  latitude =
    Number(latitude);

  longitude =
    Number(longitude);


  if (
    isNaN(latitude) ||
    isNaN(longitude)
  ) {

    return {

      valid: false,

      pesan:
        'Koordinat GPS tidak valid.'

    };

  }


  const jarak =
    hitungJarak(

      latitude,
      longitude,

      LATITUDE_KAMPUS,
      LONGITUDE_KAMPUS

    );


  const valid =
    jarak <= RADIUS_KAMPUS;


  return {

    valid: valid,

    jarak: Math.round(jarak),

    radius: RADIUS_KAMPUS,

    latitudeKampus:
      LATITUDE_KAMPUS,

    longitudeKampus:
      LONGITUDE_KAMPUS,

    pesan: valid

      ? 'Anda berada di dalam area kampus.'

      : 'Anda berada di luar area kampus.'

  };

}


/* ==================================================
   CEK PRESENSI HARI INI
================================================== */

function cekPresensiHariIni(nim) {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  const sheet =
    ss.getSheetByName(
      SHEET_PRESENSI
    );

  if (!sheet) {

    throw new Error(
      'Sheet PRESENSI tidak ditemukan.'
    );

  }


  const data =
    sheet.getDataRange().getValues();


  const timezone =
    Session.getScriptTimeZone();


  const hariIni =
    Utilities.formatDate(
      new Date(),
      timezone,
      'yyyy-MM-dd'
    );


  nim =
    String(nim).trim();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const nimData =
      String(data[i][3]).trim();

    const tanggalData =
      data[i][1];


    let tanggal =
      '';


    if (
      tanggalData instanceof Date
    ) {

      tanggal =
        Utilities.formatDate(
          tanggalData,
          timezone,
          'yyyy-MM-dd'
        );

    } else {

      tanggal =
        String(tanggalData).trim();

    }


    if (
      nimData === nim &&
      tanggal === hariIni
    ) {

      return {

        sudah: true,

        waktu:
          String(data[i][2]),

        nama:
          String(data[i][4]),

        dosen:
          String(data[i][5])

      };

    }

  }


  return {

    sudah: false

  };

}


/* ==================================================
   SIMPAN PRESENSI
================================================== */

function simpanPresensi(data) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {

    /* ==================================================
       VALIDASI DATA AWAL
    ================================================== */

    if (!data) {

      throw new Error(
        'Data presensi tidak tersedia.'
      );

    }


    const nim =
      String(data.nim || '').trim();

    const nidn =
      String(data.nidn || '').trim();

    const latitude =
      Number(data.latitude);

    const longitude =
      Number(data.longitude);


    if (!nim) {

      throw new Error(
        'NIM tidak tersedia.'
      );

    }


    if (!nidn) {

      throw new Error(
        'Dosen belum dipilih.'
      );

    }


    /* ==================================================
       BUKA SPREADSHEET
    ================================================== */

    const ss =
      SpreadsheetApp.openById(
        SPREADSHEET_ID
      );


    /* ==================================================
       VALIDASI MAHASISWA
    ================================================== */

    const mahasiswa =
      getMahasiswa(nim);


    if (!mahasiswa ||
        !mahasiswa.ditemukan) {

      throw new Error(
        'NIM tidak ditemukan.'
      );

    }


    if (
      String(mahasiswa.status)
        .trim()
        .toUpperCase() !== 'AKTIF'
    ) {

      throw new Error(
        'Status mahasiswa tidak aktif.'
      );

    }


    /* ==================================================
       VALIDASI DOSEN
    ================================================== */

    const sheetDosen =
      ss.getSheetByName(
        SHEET_DOSEN
      );


    if (!sheetDosen) {

      throw new Error(
        'Sheet DOSEN tidak ditemukan.'
      );

    }


    /*
     * getDisplayValues()
     * supaya NIDN dibaca persis
     * seperti yang tampil di Sheet
     */

    const dataDosen =
      sheetDosen
        .getDataRange()
        .getDisplayValues();


    let dosen = null;


    for (
      let i = 1;
      i < dataDosen.length;
      i++
    ) {

      const nidnSheet =
        String(
          dataDosen[i][0] || ''
        ).trim();

      const namaSheet =
        String(
          dataDosen[i][1] || ''
        ).trim();

      const statusSheet =
        String(
          dataDosen[i][2] || ''
        )
        .trim()
        .toUpperCase();


      /*
       * Cocokkan NIDN
       */

      if (
        nidnSheet === nidn
      ) {

        /*
         * NIDN ditemukan
         */

        if (
          statusSheet !== 'AKTIF'
        ) {

          throw new Error(
            'Dosen ' +
            namaSheet +
            ' tidak aktif.'
          );

        }


        dosen = {

          nidn:
            nidnSheet,

          nama:
            namaSheet,

          status:
            statusSheet

        };


        break;

      }

    }


    /*
     * DOSEN TIDAK DITEMUKAN
     */

    if (!dosen) {

      throw new Error(
        'NIDN dosen "' +
        nidn +
        '" tidak ditemukan pada sheet DOSEN.'
      );

    }


    /* ==================================================
       CEK PRESENSI GANDA
    ================================================== */

    const sudah =
      cekPresensiHariIni(
        nim
      );


    if (sudah.sudah) {

      throw new Error(

        'Mahasiswa sudah melakukan ' +
        'presensi hari ini pada pukul ' +
        sudah.waktu +
        '.'

      );

    }


    /* ==================================================
       VALIDASI GPS
    ================================================== */

    if (
      isNaN(latitude) ||
      isNaN(longitude)
    ) {

      throw new Error(
        'Lokasi GPS tidak valid.'
      );

    }


    /* ==================================================
       CEK JARAK KAMPUS
    ================================================== */

    const lokasi =
      cekLokasi(
        latitude,
        longitude
      );


    if (!lokasi.valid) {

      throw new Error(

        'Presensi ditolak. ' +
        'Jarak Anda dari kampus ' +
        lokasi.jarak +
        ' meter. ' +
        'Maksimal radius ' +
        lokasi.radius +
        ' meter.'

      );

    }


    /* ==================================================
       WAKTU SERVER
    ================================================== */

    const sekarang =
      new Date();

    const timezone =
      Session.getScriptTimeZone();


    const tanggal =
      Utilities.formatDate(
        sekarang,
        timezone,
        'yyyy-MM-dd'
      );


    const waktu =
      Utilities.formatDate(
        sekarang,
        timezone,
        'HH:mm:ss'
      );


    /* ==================================================
       SHEET PRESENSI
    ================================================== */

    const sheetPresensi =
      ss.getSheetByName(
        SHEET_PRESENSI
      );


    if (!sheetPresensi) {

      throw new Error(
        'Sheet PRESENSI tidak ditemukan.'
      );

    }


    /* ==================================================
       SIMPAN
    ================================================== */

    sheetPresensi.appendRow([

      sekarang,

      tanggal,

      waktu,

      mahasiswa.nim,

      mahasiswa.nama,

      dosen.nama,

      latitude,

      longitude,

      Math.round(
        lokasi.jarak
      ),

      'HADIR'

    ]);


    /* ==================================================
       HASIL
    ================================================== */

    return {

      sukses: true,

      nim:
        mahasiswa.nim,

      nama:
        mahasiswa.nama,

      dosen:
        dosen.nama,

      tanggal:
        tanggal,

      waktu:
        waktu,

      jarak:
        Math.round(
          lokasi.jarak
        ),

      status:
        'HADIR'

    };


  } finally {

    lock.releaseLock();

  }

}
