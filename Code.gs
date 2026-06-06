/**
 * Koperasi Syariah MIKA
 * Backend Google Apps Script untuk Google Sheets.
 */

const SPREADSHEET_ID = '1rsX3XUeippTQCUlgrA_UdKf-iRB2-vmkC-X8ji8cMWQ';
const APP_TIMEZONE = 'Asia/Jakarta';
const DEFAULT_OFFICER = 'Administrator';
const SETUP_VERSION = 'MIKA_SCHEMA_2026_06_V1';

const SCHEMA = {
  Anggota: [
    'ID Anggota', 'Tanggal Gabung', 'Nama Lengkap', 'Kategori', 'No HP',
    'Alamat', 'Simpanan Pokok', 'Status', 'Tanggal Keluar', 'Keterangan'
  ],
  Simpanan: [
    'ID Transaksi', 'Tanggal', 'ID Anggota', 'Nama Anggota', 'Jenis Simpanan',
    'Tipe Transaksi', 'Jumlah', 'Keterangan', 'Petugas'
  ],
  Pinjaman: [
    'ID Pinjaman', 'Tanggal Pengajuan', 'ID Anggota', 'Nama Anggota',
    'Pokok Pinjaman', 'Tenor', 'Margin/Bunga', 'Biaya Admin', 'Total Tagihan',
    'Angsuran Per Bulan', 'Sisa Pinjaman', 'Status', 'Keterangan'
  ],
  Angsuran: [
    'ID Angsuran', 'Tanggal Bayar', 'ID Pinjaman', 'ID Anggota', 'Nama Anggota',
    'Angsuran Ke', 'Jumlah Bayar', 'Sisa Pinjaman Setelah Bayar',
    'Keterangan', 'Petugas'
  ],
  Kas: [
    'ID Kas', 'Tanggal', 'Kategori', 'Deskripsi', 'Masuk', 'Keluar',
    'Sumber Transaksi', 'Ref ID'
  ],
  Pengaturan: ['Key', 'Value', 'Deskripsi']
};

const DEFAULT_SETTINGS = [
  ['nama_koperasi', 'Koperasi Syariah MIKA', 'Nama resmi koperasi'],
  ['alamat_koperasi', 'Alamat koperasi belum diatur', 'Alamat lengkap koperasi'],
  ['telepon_koperasi', '-', 'Nomor telepon koperasi'],
  ['email_koperasi', '-', 'Email resmi koperasi'],
  ['logo_url', '', 'URL logo koperasi'],
  ['simpanan_pokok', '165000', 'Simpanan pokok anggota baru'],
  ['simpanan_wajib', '50000', 'Simpanan wajib bulanan'],
  ['margin_pinjaman', '5', 'Persentase margin pembiayaan'],
  ['biaya_admin', '20000', 'Biaya administrasi pembiayaan']
];

function doGet() {
  ensureSetup();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Koperasi Syariah MIKA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDb_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Spreadsheet aktif tidak ditemukan.');
  return active;
}

function ensureSetup(force) {
  var ss = getDb_();
  var properties = PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
  var sheetsReady = Object.keys(SCHEMA).every(function(name) {
    return !!ss.getSheetByName(name);
  });
  if (!force && sheetsReady && properties.getProperty('MIKA_SETUP_VERSION') === SETUP_VERSION) {
    return {
      success: true,
      message: 'Database siap digunakan.',
      spreadsheetUrl: ss.getUrl()
    };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Object.keys(SCHEMA).forEach(function(name) {
      ensureSheet_(ss, name, SCHEMA[name]);
    });
    seedSettings_();
    var cashSheet = ss.getSheetByName('Kas');
    var hasTransactions = ['Simpanan', 'Pinjaman', 'Angsuran'].some(function(name) {
      return ss.getSheetByName(name).getLastRow() > 1;
    });
    if (cashSheet.getLastRow() < 2 && hasTransactions) rebuildCashLedger_();
    formatSheets_();
    properties.setProperty('MIKA_SETUP_VERSION', SETUP_VERSION);
    return {
      success: true,
      message: 'Database siap digunakan.',
      spreadsheetUrl: ss.getUrl()
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  var currentHeaders = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length))
    .getDisplayValues()[0]
    .slice(0, headers.length);

  if (headers.join('|') === currentHeaders.join('|')) return;

  var legacyValues = sheet.getDataRange().getValues();
  backupLegacySheet_(ss, sheet, name);
  var migratedRows = migrateLegacyRows_(name, legacyValues);

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (migratedRows.length) {
    sheet.getRange(2, 1, migratedRows.length, headers.length).setValues(migratedRows);
  }
}

function backupLegacySheet_(ss, sheet, name) {
  var suffix = Utilities.formatDate(new Date(), APP_TIMEZONE, 'yyyyMMdd_HHmmss');
  var base = ('_Legacy_' + name + '_' + suffix).substring(0, 90);
  var backupName = base;
  var counter = 1;
  while (ss.getSheetByName(backupName)) {
    backupName = (base + '_' + counter).substring(0, 99);
    counter++;
  }
  var backup = sheet.copyTo(ss).setName(backupName);
  try {
    backup.hideSheet();
  } catch (error) {
    // Sheet cadangan tetap dipertahankan jika tidak dapat disembunyikan.
  }
}

function migrateLegacyRows_(sheetName, values) {
  if (!values || values.length < 2) return [];
  var oldHeaders = values[0].map(function(header) {
    return String(header || '').trim();
  });
  var rows = values.slice(1).filter(function(row) {
    return row.some(function(cell) { return cell !== '' && cell !== null; });
  });
  var index = {};
  oldHeaders.forEach(function(header, i) {
    index[normalizeHeader_(header)] = i;
  });

  function value(row, aliases) {
    for (var i = 0; i < aliases.length; i++) {
      var key = normalizeHeader_(aliases[i]);
      if (Object.prototype.hasOwnProperty.call(index, key)) return row[index[key]];
    }
    return '';
  }

  if (sheetName === 'Anggota') {
    return rows.map(function(row) {
      var category = value(row, ['Kategori', 'Keterangan']) || 'Umum';
      if (['Dosen', 'Mahasiswa', 'Umum'].indexOf(String(category)) === -1) category = 'Umum';
      return [
        value(row, ['ID Anggota', 'ID_Anggota']),
        value(row, ['Tanggal Gabung', 'Tanggal_Daftar', 'Tanggal']),
        value(row, ['Nama Lengkap', 'Nama']),
        category,
        String(value(row, ['No HP', 'Kontak', 'No_Telepon']) || ''),
        value(row, ['Alamat']),
        number_(value(row, ['Simpanan Pokok', 'Simpanan_Pokok'])),
        value(row, ['Status']) || 'Aktif',
        value(row, ['Tanggal Keluar']),
        value(row, ['Catatan']) || ''
      ];
    });
  }

  if (sheetName === 'Simpanan') {
    var memberMap = getMemberMap_();
    return rows.map(function(row) {
      var amount = number_(value(row, ['Jumlah', 'Nominal']));
      var note = String(value(row, ['Keterangan']) || '');
      var type = value(row, ['Jenis Simpanan']);
      if (!type) {
        if (/pokok/i.test(note)) type = 'Simpanan Pokok';
        else if (/wajib/i.test(note)) type = 'Simpanan Wajib';
        else type = 'Simpanan Sukarela';
      }
      var memberId = value(row, ['ID Anggota', 'ID_Anggota']);
      return [
        value(row, ['ID Transaksi', 'ID_Transaksi']),
        value(row, ['Tanggal']),
        memberId,
        value(row, ['Nama Anggota']) || (memberMap[memberId] || ''),
        type,
        value(row, ['Tipe Transaksi']) || (amount < 0 ? 'Tarik Tunai' : 'Setor Tunai'),
        Math.abs(amount),
        note,
        value(row, ['Petugas']) || 'Migrasi Data'
      ];
    });
  }

  if (sheetName === 'Pinjaman') {
    var loanMemberMap = getMemberMap_();
    return rows.map(function(row) {
      var memberId = value(row, ['ID Anggota', 'ID_Anggota']);
      var principal = number_(value(row, ['Pokok Pinjaman', 'Plafon']));
      var tenor = number_(value(row, ['Tenor', 'Tenor_Bulan']));
      var margin = number_(value(row, ['Margin/Bunga', 'Bunga_Persen']));
      if (margin > 0 && margin <= 1) margin *= 100;
      var admin = number_(value(row, ['Biaya Admin']));
      var total = number_(value(row, ['Total Tagihan', 'Total_Tagihan'])) || principal;
      var remaining = number_(value(row, ['Sisa Pinjaman', 'Sisa_Tagihan']));
      return [
        value(row, ['ID Pinjaman', 'ID_Pinjaman']),
        value(row, ['Tanggal Pengajuan', 'Tanggal']),
        memberId,
        loanMemberMap[memberId] || value(row, ['Nama Anggota', 'Nama_Peminjam']),
        principal,
        tenor,
        margin,
        admin,
        total,
        tenor > 0 ? total / tenor : total,
        remaining,
        value(row, ['Status']) || (remaining <= 0 ? 'Lunas' : 'Aktif'),
        value(row, ['Keterangan']) || 'Migrasi data lama'
      ];
    });
  }

  if (sheetName === 'Angsuran') {
    var loanMap = getLoanMap_();
    var sequence = {};
    return rows.map(function(row) {
      var loanId = value(row, ['ID Pinjaman', 'ID_Pinjaman']);
      sequence[loanId] = (sequence[loanId] || 0) + 1;
      var loan = loanMap[loanId] || {};
      return [
        value(row, ['ID Angsuran', 'ID_Angsuran']),
        value(row, ['Tanggal Bayar', 'Tanggal_Bayar']),
        loanId,
        loan.memberId || value(row, ['ID Anggota']),
        loan.memberName || value(row, ['Nama Anggota']),
        number_(value(row, ['Angsuran Ke'])) || sequence[loanId],
        number_(value(row, ['Jumlah Bayar', 'Nominal_Bayar'])),
        number_(value(row, ['Sisa Pinjaman Setelah Bayar', 'Sisa_Tagihan'])),
        value(row, ['Keterangan']) || 'Migrasi data lama',
        value(row, ['Petugas']) || 'Migrasi Data'
      ];
    });
  }

  if (sheetName === 'Kas') {
    return rows.map(function(row) {
      var nominal = number_(value(row, ['Nominal']));
      var type = String(value(row, ['Tipe']) || '');
      var masuk = number_(value(row, ['Masuk'])) || (/masuk/i.test(type) ? nominal : 0);
      var keluar = number_(value(row, ['Keluar'])) || (/keluar/i.test(type) ? nominal : 0);
      if (!masuk && !keluar) return null;
      return [
        value(row, ['ID Kas', 'ID_Kas']),
        value(row, ['Tanggal']),
        value(row, ['Kategori']) || 'Lainnya',
        value(row, ['Deskripsi', 'Keterangan']) || 'Migrasi data lama',
        masuk,
        keluar,
        value(row, ['Sumber Transaksi']) || 'Migrasi',
        value(row, ['Ref ID']) || ''
      ];
    }).filter(Boolean);
  }

  if (sheetName === 'Pengaturan') {
    return rows.map(function(row) {
      var rawKey = String(value(row, ['Key', 'Parameter']) || '');
      var keyMap = {
        'nama koperasi': 'nama_koperasi',
        'bunga pinjaman': 'margin_pinjaman',
        'margin pinjaman': 'margin_pinjaman'
      };
      var key = keyMap[rawKey.toLowerCase()] || rawKey;
      var settingValue = value(row, ['Value', 'Nilai']);
      if (key === 'margin_pinjaman' && number_(settingValue) > 0 && number_(settingValue) <= 1) {
        settingValue = number_(settingValue) * 100;
      }
      return [key, settingValue, value(row, ['Deskripsi'])];
    }).filter(function(row) { return row[0]; });
  }

  return [];
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/[\s_\/-]+/g, '');
}

function seedSettings_() {
  var sheet = getDb_().getSheetByName('Pengaturan');
  var existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function(row) {
      existing[String(row[0])] = true;
    });
  }
  var missing = DEFAULT_SETTINGS.filter(function(row) { return !existing[row[0]]; });
  if (missing.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }
}

function formatSheets_() {
  var ss = getDb_();
  Object.keys(SCHEMA).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var width = SCHEMA[name].length;
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, width)
      .setBackground('#123B73')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.autoResizeColumns(1, width);
  });

  setFormats_('Anggota', ['B:B', 'G:G', 'I:I'], ['yyyy-mm-dd', '#,##0', 'yyyy-mm-dd']);
  setFormats_('Simpanan', ['B:B', 'G:G'], ['yyyy-mm-dd', '#,##0']);
  setFormats_('Pinjaman', ['B:B', 'E:E', 'G:K'], ['yyyy-mm-dd', '#,##0', '#,##0.00']);
  setFormats_('Angsuran', ['B:B', 'G:H'], ['yyyy-mm-dd', '#,##0']);
  setFormats_('Kas', ['B:B', 'E:F'], ['yyyy-mm-dd', '#,##0']);
}

function setFormats_(sheetName, ranges, formats) {
  var sheet = getDb_().getSheetByName(sheetName);
  if (!sheet) return;
  ranges.forEach(function(range, i) {
    sheet.getRange(range).setNumberFormat(formats[i]);
  });
}

function getAppData() {
  ensureSetup();
  var members = getMembers({ status: 'Aktif' });
  return {
    settings: getSettings(),
    activeMembers: members.items,
    dashboard: getDashboardData({ period: '6months' }),
    spreadsheetUrl: getDb_().getUrl()
  };
}

function getDashboardData(filter) {
  ensureSetup();
  filter = filter || { period: '6months' };
  var period = resolvePeriod_(filter);
  var cashRows = getObjects_('Kas');
  var savingsRows = getObjects_('Simpanan');
  var loanRows = getObjects_('Pinjaman');
  var memberRows = getObjects_('Anggota');

  var totalCash = cashRows.reduce(function(total, row) {
    return total + number_(row['Masuk']) - number_(row['Keluar']);
  }, 0);
  var totalSavings = savingsRows.reduce(function(total, row) {
    var sign = row['Tipe Transaksi'] === 'Tarik Tunai' ? -1 : 1;
    return total + sign * number_(row['Jumlah']);
  }, 0);
  var remainingLoans = loanRows.reduce(function(total, row) {
    return total + Math.max(0, number_(row['Sisa Pinjaman']));
  }, 0);

  var categories = { Dosen: 0, Mahasiswa: 0, Umum: 0 };
  var activeMembers = memberRows.filter(function(row) {
    var active = row['Status'] === 'Aktif';
    if (active && categories[row['Kategori']] !== undefined) categories[row['Kategori']]++;
    return active;
  });

  var periodCashRows = cashRows.filter(function(row) {
    return isWithinPeriod_(row['Tanggal'], period);
  });
  var cashInPeriod = sum_(periodCashRows, 'Masuk');
  var cashOutPeriod = sum_(periodCashRows, 'Keluar');

  var monthly = {};
  periodCashRows.forEach(function(row) {
    var date = parseDate_(row['Tanggal']);
    if (!date) return;
    var key = Utilities.formatDate(date, APP_TIMEZONE, 'yyyy-MM');
    if (!monthly[key]) monthly[key] = { in: 0, out: 0 };
    monthly[key].in += number_(row['Masuk']);
    monthly[key].out += number_(row['Keluar']);
  });
  var monthKeys = Object.keys(monthly).sort();
  var cashFlow = {
    labels: monthKeys.map(function(key) {
      return monthLabel_(key);
    }),
    incoming: monthKeys.map(function(key) { return monthly[key].in; }),
    outgoing: monthKeys.map(function(key) { return monthly[key].out; })
  };

  var composition = {
    'Simpanan Pokok': 0,
    'Simpanan Wajib': 0,
    'Simpanan Sukarela': 0
  };
  savingsRows.forEach(function(row) {
    var type = row['Jenis Simpanan'];
    if (composition[type] === undefined) return;
    var sign = row['Tipe Transaksi'] === 'Tarik Tunai' ? -1 : 1;
    composition[type] += sign * number_(row['Jumlah']);
  });

  var healthy = totalCash > 0 && (totalSavings <= 0 || totalCash >= totalSavings * 0.1);
  return {
    totals: {
      cash: totalCash,
      savings: totalSavings,
      remainingLoans: remainingLoans,
      activeMembers: activeMembers.length,
      lecturers: categories.Dosen,
      students: categories.Mahasiswa,
      publicMembers: categories.Umum,
      cashInPeriod: cashInPeriod,
      cashOutPeriod: cashOutPeriod
    },
    health: healthy ? 'Sehat' : 'Perlu Perhatian',
    cashFlow: cashFlow,
    savingsComposition: {
      labels: Object.keys(composition),
      values: Object.keys(composition).map(function(key) {
        return Math.max(0, composition[key]);
      })
    },
    period: {
      start: formatDate_(period.start),
      end: formatDate_(period.end)
    }
  };
}

function getMembers(filters) {
  ensureSetup();
  filters = filters || {};
  var query = String(filters.query || '').toLowerCase().trim();
  var items = getObjects_('Anggota').filter(function(row) {
    if (query) {
      var haystack = (row['ID Anggota'] + ' ' + row['Nama Lengkap'] + ' ' + row['No HP']).toLowerCase();
      if (haystack.indexOf(query) === -1) return false;
    }
    if (filters.status && filters.status !== 'Semua' && row['Status'] !== filters.status) return false;
    if (filters.category && filters.category !== 'Semua' && row['Kategori'] !== filters.category) return false;
    if (filters.startDate || filters.endDate) {
      var period = resolvePeriod_({
        period: 'custom',
        startDate: filters.startDate || '1900-01-01',
        endDate: filters.endDate || '2999-12-31'
      });
      if (!isWithinPeriod_(row['Tanggal Gabung'], period)) return false;
    }
    return true;
  });
  items.sort(function(a, b) {
    return String(b['ID Anggota']).localeCompare(String(a['ID Anggota']));
  });
  return { items: items, total: items.length };
}

function addMember(data) {
  return withLock_(function() {
    validateRequired_(data, ['name', 'category', 'phone', 'joinDate']);
    if (['Dosen', 'Mahasiswa', 'Umum'].indexOf(data.category) === -1) {
      throw new Error('Kategori anggota tidak valid.');
    }
    var principalSaving = positiveOrZero_(data.principalSaving, 'Simpanan pokok');
    var memberId = nextId_('Anggota', 'MIKA-', 4);
    var joinDate = parseRequiredDate_(data.joinDate, 'Tanggal gabung');
    var name = cleanText_(data.name, 150);
    var sheet = getDb_().getSheetByName('Anggota');
    sheet.appendRow([
      memberId, joinDate, name, data.category, cleanText_(data.phone, 30),
      cleanText_(data.address, 500), principalSaving, 'Aktif', '',
      cleanText_(data.notes, 500)
    ]);

    if (principalSaving > 0) {
      var transactionId = nextId_('Simpanan', 'SMP-', 6);
      appendSaving_({
        transactionId: transactionId,
        date: joinDate,
        memberId: memberId,
        memberName: name,
        savingType: 'Simpanan Pokok',
        transactionType: 'Setor Tunai',
        amount: principalSaving,
        notes: 'Simpanan pokok anggota baru',
        officer: cleanText_(data.officer, 100) || DEFAULT_OFFICER
      });
      appendCash_({
        date: joinDate,
        category: 'Simpanan Masuk',
        description: 'Simpanan Pokok - ' + name,
        incoming: principalSaving,
        outgoing: 0,
        source: 'Simpanan',
        refId: transactionId
      });
    }
    SpreadsheetApp.flush();
    return { success: true, message: 'Anggota berhasil ditambahkan.', id: memberId };
  });
}

function updateMember(data) {
  return withLock_(function() {
    validateRequired_(data, ['id', 'name', 'category', 'phone', 'status']);
    if (['Dosen', 'Mahasiswa', 'Umum'].indexOf(data.category) === -1) {
      throw new Error('Kategori anggota tidak valid.');
    }
    if (['Aktif', 'Keluar', 'Nonaktif'].indexOf(data.status) === -1) {
      throw new Error('Status anggota tidak valid.');
    }
    var found = findRowById_('Anggota', data.id);
    if (!found) throw new Error('Data anggota tidak ditemukan.');
    if (data.status === 'Keluar' && found.values[7] !== 'Keluar') {
      throw new Error('Gunakan tombol Proses Keluar agar saldo dan kewajiban anggota dihitung.');
    }
    var row = found.values;
    row[2] = cleanText_(data.name, 150);
    row[3] = data.category;
    row[4] = cleanText_(data.phone, 30);
    row[5] = cleanText_(data.address, 500);
    row[7] = data.status;
    row[9] = cleanText_(data.notes, 500);
    if (data.status !== 'Keluar') row[8] = '';
    getDb_().getSheetByName('Anggota').getRange(found.row, 1, 1, row.length).setValues([row]);
    return { success: true, message: 'Data anggota berhasil diperbarui.' };
  });
}

function getMemberFinancialSummary(memberId) {
  ensureSetup();
  return getMemberFinancialSummaryInternal_(memberId);
}

function getMemberFinancialSummaryInternal_(memberId) {
  var member = getMemberById_(memberId);
  if (!member) throw new Error('Data anggota tidak ditemukan.');
  var balances = getSavingBalances_(memberId);
  var loans = getObjects_('Pinjaman').filter(function(row) {
    return row['ID Anggota'] === memberId && number_(row['Sisa Pinjaman']) > 0;
  });
  var outstanding = loans.reduce(function(total, row) {
    return total + number_(row['Sisa Pinjaman']);
  }, 0);
  var totalSavings = balances.pokok + balances.wajib + balances.sukarela;
  return {
    member: member,
    savings: balances,
    totalSavings: totalSavings,
    outstandingLoan: outstanding,
    refundable: totalSavings - outstanding,
    hasObligation: outstanding > totalSavings
  };
}

function processMemberExit(memberId, force) {
  return withLock_(function() {
    var memberFound = findRowById_('Anggota', memberId);
    if (!memberFound) throw new Error('Data anggota tidak ditemukan.');
    if (memberFound.values[7] === 'Keluar') throw new Error('Anggota sudah berstatus Keluar.');
    var summary = getMemberFinancialSummaryInternal_(memberId);
    if (summary.hasObligation && !force) {
      throw new Error('Anggota masih memiliki kewajiban. Centang konfirmasi untuk melanjutkan.');
    }

    var exitDate = new Date();
    var credit = Math.min(summary.totalSavings, summary.outstandingLoan);
    if (credit > 0) applyExitCreditToLoans_(memberId, credit, exitDate);

    var balanceMap = [
      ['Simpanan Pokok', summary.savings.pokok],
      ['Simpanan Wajib', summary.savings.wajib],
      ['Simpanan Sukarela', summary.savings.sukarela]
    ];
    balanceMap.forEach(function(item) {
      if (item[1] <= 0) return;
      appendSaving_({
        transactionId: nextId_('Simpanan', 'SMP-', 6),
        date: exitDate,
        memberId: memberId,
        memberName: memberFound.values[2],
        savingType: item[0],
        transactionType: 'Tarik Tunai',
        amount: item[1],
        notes: 'Penutupan saldo karena anggota keluar',
        officer: DEFAULT_OFFICER
      });
    });

    var refundable = Math.max(0, summary.refundable);
    if (refundable > 0) {
      appendCash_({
        date: exitDate,
        category: 'Penarikan Simpanan',
        description: 'Pengembalian dana anggota keluar - ' + memberFound.values[2],
        incoming: 0,
        outgoing: refundable,
        source: 'Anggota Keluar',
        refId: memberId
      });
    }

    memberFound.values[7] = 'Keluar';
    memberFound.values[8] = exitDate;
    memberFound.values[9] = appendNote_(memberFound.values[9], 'Diproses keluar pada ' + formatDate_(exitDate));
    getDb_().getSheetByName('Anggota')
      .getRange(memberFound.row, 1, 1, memberFound.values.length)
      .setValues([memberFound.values]);
    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Proses anggota keluar berhasil.',
      refundable: refundable,
      remainingObligation: Math.max(0, summary.outstandingLoan - summary.totalSavings)
    };
  });
}

function applyExitCreditToLoans_(memberId, credit, date) {
  var sheet = getDb_().getSheetByName('Pinjaman');
  var loans = getObjectsWithRows_('Pinjaman').filter(function(item) {
    return item.data['ID Anggota'] === memberId && number_(item.data['Sisa Pinjaman']) > 0;
  });
  var remainingCredit = credit;
  loans.forEach(function(item) {
    if (remainingCredit <= 0) return;
    var balance = number_(item.data['Sisa Pinjaman']);
    var applied = Math.min(balance, remainingCredit);
    var after = Math.max(0, balance - applied);
    var installmentId = nextId_('Angsuran', 'ANG-', 6);
    var installmentNo = getNextInstallmentNumber_(item.data['ID Pinjaman']);
    getDb_().getSheetByName('Angsuran').appendRow([
      installmentId, date, item.data['ID Pinjaman'], memberId,
      item.data['Nama Anggota'], installmentNo, applied, after,
      'Kompensasi simpanan saat anggota keluar', DEFAULT_OFFICER
    ]);
    sheet.getRange(item.row, 11).setValue(after);
    sheet.getRange(item.row, 12).setValue(after <= 0 ? 'Lunas' : item.data['Status']);
    remainingCredit -= applied;
  });
}

function getSavings(filters) {
  ensureSetup();
  filters = filters || {};
  var query = String(filters.query || '').toLowerCase().trim();
  var period = filters.startDate || filters.endDate ? resolvePeriod_({
    period: 'custom',
    startDate: filters.startDate || '1900-01-01',
    endDate: filters.endDate || '2999-12-31'
  }) : null;
  var items = getObjects_('Simpanan').filter(function(row) {
    if (query) {
      var haystack = (row['ID Transaksi'] + ' ' + row['Nama Anggota'] + ' ' + row['ID Anggota']).toLowerCase();
      if (haystack.indexOf(query) === -1) return false;
    }
    if (filters.savingType && filters.savingType !== 'Semua' &&
        row['Jenis Simpanan'] !== filters.savingType) return false;
    if (filters.transactionType && filters.transactionType !== 'Semua' &&
        row['Tipe Transaksi'] !== filters.transactionType) return false;
    if (period && !isWithinPeriod_(row['Tanggal'], period)) return false;
    return true;
  });
  items.sort(sortByDateDesc_('Tanggal'));
  return { items: items, total: items.length };
}

function addSavingTransaction(data) {
  return withLock_(function() {
    validateRequired_(data, ['date', 'memberId', 'savingType', 'transactionType', 'amount']);
    if (['Simpanan Pokok', 'Simpanan Wajib', 'Simpanan Sukarela'].indexOf(data.savingType) === -1) {
      throw new Error('Jenis simpanan tidak valid.');
    }
    if (['Setor Tunai', 'Tarik Tunai'].indexOf(data.transactionType) === -1) {
      throw new Error('Tipe transaksi tidak valid.');
    }
    var amount = positive_(data.amount, 'Jumlah transaksi');
    var member = getMemberById_(data.memberId);
    if (!member) throw new Error('Anggota tidak ditemukan.');
    if (member['Status'] !== 'Aktif') throw new Error('Hanya anggota aktif yang dapat bertransaksi.');
    if (data.transactionType === 'Tarik Tunai') {
      var balances = getSavingBalances_(data.memberId);
      var key = savingBalanceKey_(data.savingType);
      if (amount > balances[key]) {
        throw new Error('Penarikan melebihi saldo ' + data.savingType + '.');
      }
    }

    var id = nextId_('Simpanan', 'SMP-', 6);
    var date = parseRequiredDate_(data.date, 'Tanggal transaksi');
    appendSaving_({
      transactionId: id,
      date: date,
      memberId: data.memberId,
      memberName: member['Nama Lengkap'],
      savingType: data.savingType,
      transactionType: data.transactionType,
      amount: amount,
      notes: cleanText_(data.notes, 500),
      officer: cleanText_(data.officer, 100) || DEFAULT_OFFICER
    });
    appendCash_({
      date: date,
      category: data.transactionType === 'Setor Tunai' ? 'Simpanan Masuk' : 'Penarikan Simpanan',
      description: data.savingType + ' - ' + member['Nama Lengkap'],
      incoming: data.transactionType === 'Setor Tunai' ? amount : 0,
      outgoing: data.transactionType === 'Tarik Tunai' ? amount : 0,
      source: 'Simpanan',
      refId: id
    });
    SpreadsheetApp.flush();
    return { success: true, message: 'Transaksi simpanan berhasil disimpan.', id: id };
  });
}

function generateSavingReceiptPdf(transactionId) {
  ensureSetup();
  var transaction = getObjectById_('Simpanan', transactionId);
  if (!transaction) throw new Error('Transaksi simpanan tidak ditemukan.');
  var settings = getSettings();
  return buildReceiptPdf_({
    title: 'STRUK TRANSAKSI SIMPANAN',
    idLabel: 'ID Transaksi',
    id: transaction['ID Transaksi'],
    date: transaction['Tanggal'],
    rows: [
      ['Nama Anggota', transaction['Nama Anggota']],
      ['Jenis Simpanan', transaction['Jenis Simpanan']],
      ['Tipe Transaksi', transaction['Tipe Transaksi']],
      ['Jumlah', formatRupiah_(transaction['Jumlah'])],
      ['Petugas', transaction['Petugas']]
    ]
  }, settings, 'Struk-Simpanan-' + transactionId + '.pdf');
}

function getLoans(filters) {
  ensureSetup();
  filters = filters || {};
  var query = String(filters.query || '').toLowerCase().trim();
  var period = filters.startDate || filters.endDate ? resolvePeriod_({
    period: 'custom',
    startDate: filters.startDate || '1900-01-01',
    endDate: filters.endDate || '2999-12-31'
  }) : null;
  var items = getObjects_('Pinjaman').filter(function(row) {
    if (query) {
      var haystack = (row['ID Pinjaman'] + ' ' + row['Nama Anggota'] + ' ' + row['ID Anggota']).toLowerCase();
      if (haystack.indexOf(query) === -1) return false;
    }
    if (filters.status && filters.status !== 'Semua' && row['Status'] !== filters.status) return false;
    if (period && !isWithinPeriod_(row['Tanggal Pengajuan'], period)) return false;
    return true;
  }).map(function(row) {
    row['Sudah Dibayar'] = Math.max(0, number_(row['Total Tagihan']) - number_(row['Sisa Pinjaman']));
    return row;
  });
  items.sort(sortByDateDesc_('Tanggal Pengajuan'));
  return { items: items, total: items.length };
}

function addLoan(data) {
  return withLock_(function() {
    validateRequired_(data, ['date', 'memberId', 'principal', 'tenor', 'margin', 'adminFee']);
    var member = getMemberById_(data.memberId);
    if (!member) throw new Error('Anggota tidak ditemukan.');
    if (member['Status'] !== 'Aktif') throw new Error('Pinjaman hanya dapat dibuat untuk anggota aktif.');

    var principal = positive_(data.principal, 'Pokok pinjaman');
    var tenor = Math.floor(positive_(data.tenor, 'Tenor'));
    var marginPercent = positiveOrZero_(data.margin, 'Margin');
    var adminFee = positiveOrZero_(data.adminFee, 'Biaya admin');
    if (tenor > 360) throw new Error('Tenor maksimal 360 bulan.');
    if (marginPercent > 1000) throw new Error('Persentase margin terlalu besar.');

    var marginNominal = principal * marginPercent / 100;
    var totalBill = principal + marginNominal + adminFee;
    var installment = totalBill / tenor;
    var loanId = nextId_('Pinjaman', 'PJM-', 4);
    var date = parseRequiredDate_(data.date, 'Tanggal pengajuan');

    getDb_().getSheetByName('Pinjaman').appendRow([
      loanId, date, data.memberId, member['Nama Lengkap'], principal, tenor,
      marginPercent, adminFee, totalBill, installment, totalBill, 'Aktif',
      cleanText_(data.notes, 500)
    ]);
    appendCash_({
      date: date,
      category: 'Pencairan Pinjaman',
      description: 'Pencairan pembiayaan - ' + member['Nama Lengkap'],
      incoming: 0,
      outgoing: principal,
      source: 'Pinjaman',
      refId: loanId
    });
    if (adminFee > 0) {
      appendCash_({
        date: date,
        category: 'Biaya Admin',
        description: 'Biaya admin pembiayaan - ' + member['Nama Lengkap'],
        incoming: adminFee,
        outgoing: 0,
        source: 'Pinjaman',
        refId: loanId
      });
    }
    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Pembiayaan berhasil dibuat.',
      id: loanId,
      calculation: {
        marginNominal: marginNominal,
        totalBill: totalBill,
        installment: installment
      }
    };
  });
}

function payInstallment(data) {
  return withLock_(function() {
    validateRequired_(data, ['loanId', 'date', 'amount']);
    var found = findRowById_('Pinjaman', data.loanId);
    if (!found) throw new Error('Data pinjaman tidak ditemukan.');
    if (found.values[11] !== 'Aktif') throw new Error('Pinjaman ini tidak berstatus Aktif.');
    var amount = positive_(data.amount, 'Jumlah bayar');
    var remainingBefore = number_(found.values[10]);
    if (amount > remainingBefore + 0.01) throw new Error('Jumlah bayar melebihi sisa pinjaman.');

    var date = parseRequiredDate_(data.date, 'Tanggal bayar');
    var remainingAfter = Math.max(0, remainingBefore - amount);
    var installmentId = nextId_('Angsuran', 'ANG-', 6);
    var installmentNo = getNextInstallmentNumber_(data.loanId);
    var officer = cleanText_(data.officer, 100) || DEFAULT_OFFICER;

    getDb_().getSheetByName('Angsuran').appendRow([
      installmentId, date, data.loanId, found.values[2], found.values[3],
      installmentNo, amount, remainingAfter, cleanText_(data.notes, 500), officer
    ]);
    var loanSheet = getDb_().getSheetByName('Pinjaman');
    loanSheet.getRange(found.row, 11).setValue(remainingAfter);
    loanSheet.getRange(found.row, 12).setValue(remainingAfter <= 0.01 ? 'Lunas' : 'Aktif');

    var totalBill = number_(found.values[8]);
    var principal = number_(found.values[4]);
    var marginNominal = principal * number_(found.values[6]) / 100;
    var marginRatio = totalBill > 0 ? marginNominal / totalBill : 0;
    var marginPart = Math.min(amount * marginRatio, marginNominal);
    var principalPart = Math.max(0, amount - marginPart);

    if (principalPart > 0) {
      appendCash_({
        date: date,
        category: 'Pembayaran Angsuran',
        description: 'Angsuran pokok #' + installmentNo + ' - ' + found.values[3],
        incoming: principalPart,
        outgoing: 0,
        source: 'Angsuran',
        refId: installmentId
      });
    }
    if (marginPart > 0) {
      appendCash_({
        date: date,
        category: 'Pendapatan Margin',
        description: 'Margin angsuran #' + installmentNo + ' - ' + found.values[3],
        incoming: marginPart,
        outgoing: 0,
        source: 'Angsuran',
        refId: installmentId
      });
    }
    SpreadsheetApp.flush();
    return {
      success: true,
      message: remainingAfter <= 0.01 ? 'Pembayaran berhasil. Pinjaman telah lunas.' : 'Pembayaran angsuran berhasil.',
      id: installmentId,
      remaining: remainingAfter
    };
  });
}

function getInstallmentHistory(loanId) {
  ensureSetup();
  var loan = getObjectById_('Pinjaman', loanId);
  if (!loan) throw new Error('Data pinjaman tidak ditemukan.');
  var items = getObjects_('Angsuran').filter(function(row) {
    return row['ID Pinjaman'] === loanId;
  });
  items.sort(function(a, b) {
    return number_(a['Angsuran Ke']) - number_(b['Angsuran Ke']);
  });
  return { loan: loan, items: items, total: items.length };
}

function generateInstallmentReceiptPdf(installmentId) {
  ensureSetup();
  var installment = getObjectById_('Angsuran', installmentId);
  if (!installment) throw new Error('Data angsuran tidak ditemukan.');
  var settings = getSettings();
  return buildReceiptPdf_({
    title: 'STRUK PEMBAYARAN ANGSURAN',
    idLabel: 'ID Angsuran',
    id: installment['ID Angsuran'],
    date: installment['Tanggal Bayar'],
    rows: [
      ['Nama Anggota', installment['Nama Anggota']],
      ['ID Pinjaman', installment['ID Pinjaman']],
      ['Angsuran Ke', installment['Angsuran Ke']],
      ['Jumlah Bayar', formatRupiah_(installment['Jumlah Bayar'])],
      ['Sisa Pinjaman', formatRupiah_(installment['Sisa Pinjaman Setelah Bayar'])],
      ['Petugas', installment['Petugas']]
    ]
  }, settings, 'Struk-Angsuran-' + installmentId + '.pdf');
}

function getFinancialReport(filter) {
  ensureSetup();
  filter = filter || {};
  var period = resolvePeriod_({
    period: 'custom',
    startDate: filter.startDate || Utilities.formatDate(new Date(new Date().getFullYear(), 0, 1), APP_TIMEZONE, 'yyyy-MM-dd'),
    endDate: filter.endDate || Utilities.formatDate(new Date(), APP_TIMEZONE, 'yyyy-MM-dd')
  });
  var allCash = getObjects_('Kas').sort(function(a, b) {
    return parseDate_(a['Tanggal']) - parseDate_(b['Tanggal']);
  });
  var openingBalance = allCash.filter(function(row) {
    var date = parseDate_(row['Tanggal']);
    return date && date < period.start;
  }).reduce(function(total, row) {
    return total + number_(row['Masuk']) - number_(row['Keluar']);
  }, 0);
  var running = openingBalance;
  var mutations = allCash.filter(function(row) {
    return isWithinPeriod_(row['Tanggal'], period);
  }).map(function(row) {
    running += number_(row['Masuk']) - number_(row['Keluar']);
    return {
      date: row['Tanggal'],
      category: row['Kategori'],
      description: row['Deskripsi'],
      incoming: number_(row['Masuk']),
      outgoing: number_(row['Keluar']),
      balance: running,
      refId: row['Ref ID']
    };
  });

  var totalIn = mutations.reduce(function(total, row) { return total + row.incoming; }, 0);
  var totalOut = mutations.reduce(function(total, row) { return total + row.outgoing; }, 0);
  var marginIncome = mutations.filter(function(row) {
    return row.category === 'Pendapatan Margin';
  }).reduce(function(total, row) { return total + row.incoming; }, 0);
  var adminIncome = mutations.filter(function(row) {
    return row.category === 'Biaya Admin';
  }).reduce(function(total, row) { return total + row.incoming; }, 0);

  var allSavings = getObjects_('Simpanan');
  var savingsBalance = {
    pokok: calculateSavingTypeBalance_(allSavings, 'Simpanan Pokok'),
    wajib: calculateSavingTypeBalance_(allSavings, 'Simpanan Wajib'),
    sukarela: calculateSavingTypeBalance_(allSavings, 'Simpanan Sukarela')
  };
  var currentCash = allCash.reduce(function(total, row) {
    return total + number_(row['Masuk']) - number_(row['Keluar']);
  }, 0);
  var receivables = getObjects_('Pinjaman').reduce(function(total, row) {
    return total + Math.max(0, number_(row['Sisa Pinjaman']));
  }, 0);
  var totalSavings = savingsBalance.pokok + savingsBalance.wajib + savingsBalance.sukarela;
  var netIncome = currentCash + receivables - totalSavings;

  return {
    period: { start: formatDate_(period.start), end: formatDate_(period.end) },
    summary: {
      totalIn: totalIn,
      totalOut: totalOut,
      marginIncome: marginIncome,
      adminIncome: adminIncome,
      netCashFlow: totalIn - totalOut
    },
    mutations: mutations,
    balanceSheet: {
      assets: {
        cash: currentCash,
        receivables: receivables,
        total: currentCash + receivables
      },
      liabilities: {
        principalSavings: savingsBalance.pokok,
        mandatorySavings: savingsBalance.wajib,
        voluntarySavings: savingsBalance.sukarela,
        netIncome: netIncome,
        total: totalSavings + netIncome
      }
    }
  };
}

function getSettings() {
  ensureSetup();
  return getSettingsInternal_();
}

function getSettingsInternal_() {
  var result = {};
  getObjects_('Pengaturan').forEach(function(row) {
    result[row['Key']] = row['Value'];
  });
  ['simpanan_pokok', 'simpanan_wajib', 'margin_pinjaman', 'biaya_admin'].forEach(function(key) {
    result[key] = number_(result[key]);
  });
  return result;
}

function saveSettings(data) {
  return withLock_(function() {
    if (!data || !cleanText_(data.nama_koperasi, 150)) throw new Error('Nama koperasi wajib diisi.');
    if (data.logo_url && !/^https?:\/\//i.test(String(data.logo_url).trim())) {
      throw new Error('Logo URL harus menggunakan alamat http atau https.');
    }
    var allowed = {};
    DEFAULT_SETTINGS.forEach(function(row) { allowed[row[0]] = row[2]; });
    var numericKeys = ['simpanan_pokok', 'simpanan_wajib', 'margin_pinjaman', 'biaya_admin'];
    numericKeys.forEach(function(key) {
      data[key] = positiveOrZero_(data[key], allowed[key] || key);
    });

    var sheet = getDb_().getSheetByName('Pengaturan');
    var rows = getObjectsWithRows_('Pengaturan');
    Object.keys(allowed).forEach(function(key) {
      var existing = rows.filter(function(item) { return item.data['Key'] === key; })[0];
      var raw = data[key] === undefined || data[key] === null ? '' : data[key];
      var value = numericKeys.indexOf(key) >= 0 ? String(raw) : cleanText_(raw, 1000);
      if (existing) {
        sheet.getRange(existing.row, 2, 1, 2).setValues([[value, allowed[key]]]);
      } else {
        sheet.appendRow([key, value, allowed[key]]);
      }
    });
    SpreadsheetApp.flush();
    return { success: true, message: 'Pengaturan berhasil disimpan.', settings: getSettingsInternal_() };
  });
}

function syncCashLedger() {
  return withLock_(function() {
    return rebuildCashLedger_();
  });
}

function rebuildCashLedger_() {
  var cashSheet = getDb_().getSheetByName('Kas');
  var manualRows = getObjects_('Kas').filter(function(row) {
    return row['Kategori'] === 'Lainnya' ||
      row['Sumber Transaksi'] === 'Manual' ||
      row['Sumber Transaksi'] === 'Anggota Keluar';
  });
  var entries = [];
  function add(data) {
    entries.push([
      '', parseDate_(data.date) || new Date(), data.category, data.description,
      number_(data.incoming), number_(data.outgoing), data.source, data.refId
    ]);
  }

  getObjects_('Simpanan').forEach(function(row) {
    if (/penutupan saldo karena anggota keluar/i.test(String(row['Keterangan'] || ''))) return;
    var amount = number_(row['Jumlah']);
    add({
      date: row['Tanggal'],
      category: row['Tipe Transaksi'] === 'Setor Tunai' ? 'Simpanan Masuk' : 'Penarikan Simpanan',
      description: row['Jenis Simpanan'] + ' - ' + row['Nama Anggota'],
      incoming: row['Tipe Transaksi'] === 'Setor Tunai' ? amount : 0,
      outgoing: row['Tipe Transaksi'] === 'Tarik Tunai' ? amount : 0,
      source: 'Simpanan',
      refId: row['ID Transaksi']
    });
  });
  getObjects_('Pinjaman').forEach(function(row) {
    add({
      date: row['Tanggal Pengajuan'],
      category: 'Pencairan Pinjaman',
      description: 'Pencairan pembiayaan - ' + row['Nama Anggota'],
      incoming: 0,
      outgoing: number_(row['Pokok Pinjaman']),
      source: 'Pinjaman',
      refId: row['ID Pinjaman']
    });
    if (number_(row['Biaya Admin']) > 0) {
      add({
        date: row['Tanggal Pengajuan'],
        category: 'Biaya Admin',
        description: 'Biaya admin pembiayaan - ' + row['Nama Anggota'],
        incoming: number_(row['Biaya Admin']),
        outgoing: 0,
        source: 'Pinjaman',
        refId: row['ID Pinjaman']
      });
    }
  });
  var loanMap = getLoanMap_();
  getObjects_('Angsuran').forEach(function(row) {
    if (/kompensasi/i.test(String(row['Keterangan'] || ''))) return;
    var loan = loanMap[row['ID Pinjaman']] || {};
    var totalBill = number_(loan.totalBill);
    var marginNominal = number_(loan.principal) * number_(loan.margin) / 100;
    var marginPart = totalBill > 0 ? number_(row['Jumlah Bayar']) * marginNominal / totalBill : 0;
    add({
      date: row['Tanggal Bayar'],
      category: 'Pembayaran Angsuran',
      description: 'Angsuran pokok #' + row['Angsuran Ke'] + ' - ' + row['Nama Anggota'],
      incoming: Math.max(0, number_(row['Jumlah Bayar']) - marginPart),
      outgoing: 0,
      source: 'Angsuran',
      refId: row['ID Angsuran']
    });
    if (marginPart > 0) {
      add({
        date: row['Tanggal Bayar'],
        category: 'Pendapatan Margin',
        description: 'Margin angsuran #' + row['Angsuran Ke'] + ' - ' + row['Nama Anggota'],
        incoming: marginPart,
        outgoing: 0,
        source: 'Angsuran',
        refId: row['ID Angsuran']
      });
    }
  });
  manualRows.forEach(function(row) {
    add({
      date: row['Tanggal'],
      category: row['Kategori'],
      description: row['Deskripsi'],
      incoming: row['Masuk'],
      outgoing: row['Keluar'],
      source: row['Sumber Transaksi'],
      refId: row['Ref ID']
    });
  });

  entries.sort(function(a, b) { return a[1] - b[1]; });
  entries.forEach(function(row, index) {
    row[0] = 'KAS-' + String(index + 1).padStart(7, '0');
  });
  if (cashSheet.getLastRow() > 1) {
    cashSheet.getRange(2, 1, cashSheet.getLastRow() - 1, SCHEMA.Kas.length).clearContent();
  }
  if (entries.length) {
    cashSheet.getRange(2, 1, entries.length, SCHEMA.Kas.length).setValues(entries);
  }
  SpreadsheetApp.flush();
  return { success: true, message: 'Ledger kas berhasil disinkronkan.' };
}

function appendSaving_(data) {
  getDb_().getSheetByName('Simpanan').appendRow([
    data.transactionId, data.date, data.memberId, data.memberName,
    data.savingType, data.transactionType, data.amount, data.notes, data.officer
  ]);
}

function appendCash_(data) {
  getDb_().getSheetByName('Kas').appendRow([
    nextId_('Kas', 'KAS-', 7), parseDate_(data.date) || new Date(),
    data.category, data.description, number_(data.incoming), number_(data.outgoing),
    data.source, data.refId
  ]);
}

function getSavingBalances_(memberId) {
  var balances = { pokok: 0, wajib: 0, sukarela: 0 };
  getObjects_('Simpanan').forEach(function(row) {
    if (row['ID Anggota'] !== memberId) return;
    var key = savingBalanceKey_(row['Jenis Simpanan']);
    if (!key) return;
    var sign = row['Tipe Transaksi'] === 'Tarik Tunai' ? -1 : 1;
    balances[key] += sign * number_(row['Jumlah']);
  });
  Object.keys(balances).forEach(function(key) {
    balances[key] = Math.max(0, balances[key]);
  });
  return balances;
}

function savingBalanceKey_(savingType) {
  return {
    'Simpanan Pokok': 'pokok',
    'Simpanan Wajib': 'wajib',
    'Simpanan Sukarela': 'sukarela'
  }[savingType];
}

function calculateSavingTypeBalance_(rows, type) {
  return rows.filter(function(row) {
    return row['Jenis Simpanan'] === type;
  }).reduce(function(total, row) {
    return total + (row['Tipe Transaksi'] === 'Tarik Tunai' ? -1 : 1) * number_(row['Jumlah']);
  }, 0);
}

function getNextInstallmentNumber_(loanId) {
  var max = 0;
  getObjects_('Angsuran').forEach(function(row) {
    if (row['ID Pinjaman'] === loanId) max = Math.max(max, number_(row['Angsuran Ke']));
  });
  return max + 1;
}

function getMemberMap_() {
  var sheet = getDb_().getSheetByName('Anggota');
  if (!sheet || sheet.getLastRow() < 2) return {};
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var idIndex = headers.indexOf('ID Anggota');
  var nameIndex = headers.indexOf('Nama Lengkap');
  if (idIndex < 0) idIndex = headers.indexOf('ID_Anggota');
  if (nameIndex < 0) nameIndex = headers.indexOf('Nama');
  var map = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues().forEach(function(row) {
    if (row[idIndex]) map[String(row[idIndex])] = String(row[nameIndex] || '');
  });
  return map;
}

function getLoanMap_() {
  var sheet = getDb_().getSheetByName('Pinjaman');
  if (!sheet || sheet.getLastRow() < 2) return {};
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(normalizeHeader_);
  function at(row, names) {
    for (var i = 0; i < names.length; i++) {
      var idx = headers.indexOf(normalizeHeader_(names[i]));
      if (idx >= 0) return row[idx];
    }
    return '';
  }
  var map = {};
  values.slice(1).forEach(function(row) {
    var id = at(row, ['ID Pinjaman', 'ID_Pinjaman']);
    if (!id) return;
    map[id] = {
      memberId: at(row, ['ID Anggota', 'ID_Anggota']),
      memberName: at(row, ['Nama Anggota', 'Nama_Peminjam']),
      principal: number_(at(row, ['Pokok Pinjaman', 'Plafon'])),
      margin: number_(at(row, ['Margin/Bunga', 'Bunga_Persen'])),
      totalBill: number_(at(row, ['Total Tagihan', 'Total_Tagihan']))
    };
  });
  return map;
}

function getObjects_(sheetName) {
  return getObjectsWithRows_(sheetName).map(function(item) { return item.data; });
}

function getObjectsWithRows_(sheetName) {
  var sheet = getDb_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var width = SCHEMA[sheetName].length;
  var values = sheet.getRange(1, 1, sheet.getLastRow(), width).getValues();
  var headers = values[0];
  return values.slice(1).map(function(row, index) {
    var data = {};
    headers.forEach(function(header, column) {
      data[header] = serializeValue_(row[column]);
    });
    return { row: index + 2, data: data };
  }).filter(function(item) {
    return item.data[headers[0]] !== '' && item.data[headers[0]] !== null;
  });
}

function getObjectById_(sheetName, id) {
  var rows = getObjects_(sheetName);
  var idHeader = SCHEMA[sheetName][0];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][idHeader]) === String(id)) return rows[i];
  }
  return null;
}

function getMemberById_(id) {
  return getObjectById_('Anggota', id);
}

function findRowById_(sheetName, id) {
  var sheet = getDb_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var width = SCHEMA[sheetName].length;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return { row: i + 2, values: values[i] };
  }
  return null;
}

function nextId_(sheetName, prefix, digits) {
  var sheet = getDb_().getSheetByName(sheetName);
  var max = 0;
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().forEach(function(row) {
      var match = String(row[0] || '').match(/(\d+)(?!.*\d)/);
      if (match) max = Math.max(max, Number(match[1]));
    });
  }
  return prefix + String(max + 1).padStart(digits, '0');
}

function resolvePeriod_(filter) {
  filter = filter || {};
  var now = new Date();
  var end = endOfDay_(now);
  var start;
  switch (filter.period) {
    case 'month':
      start = new Date(now.getFullYear(), now.getMonth(), 1);
      break;
    case 'year':
      start = new Date(now.getFullYear(), 0, 1);
      break;
    case 'all':
      start = new Date(2000, 0, 1);
      break;
    case 'custom':
      start = filter.startDate ? parseRequiredDate_(filter.startDate, 'Tanggal mulai') : new Date(2000, 0, 1);
      end = filter.endDate ? endOfDay_(parseRequiredDate_(filter.endDate, 'Tanggal selesai')) : end;
      break;
    case '6months':
    default:
      start = new Date(now.getFullYear(), now.getMonth() - 5, 1);
      break;
  }
  if (start > end) throw new Error('Tanggal mulai tidak boleh melewati tanggal selesai.');
  return { start: startOfDay_(start), end: endOfDay_(end) };
}

function isWithinPeriod_(value, period) {
  var date = parseDate_(value);
  return !!date && date >= period.start && date <= period.end;
}

function parseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  var text = String(value).substring(0, 10);
  var parts = text.split('-');
  if (parts.length === 3) {
    var date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    if (!isNaN(date.getTime())) return date;
  }
  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function parseRequiredDate_(value, label) {
  var date = parseDate_(value);
  if (!date) throw new Error(label + ' tidak valid.');
  return date;
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function endOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
}

function monthLabel_(key) {
  var parts = key.split('-');
  var names = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  return names[Number(parts[1]) - 1] + ' ' + parts[0];
}

function sortByDateDesc_(key) {
  return function(a, b) {
    var dateA = parseDate_(a[key]);
    var dateB = parseDate_(b[key]);
    return (dateB ? dateB.getTime() : 0) - (dateA ? dateA.getTime() : 0);
  };
}

function serializeValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDate_(value);
  }
  return value;
}

function formatDate_(value) {
  var date = parseDate_(value);
  return date ? Utilities.formatDate(date, APP_TIMEZONE, 'yyyy-MM-dd') : '';
}

function formatRupiah_(value) {
  return 'Rp ' + Math.round(number_(value)).toLocaleString('id-ID');
}

function number_(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === '') return 0;
  var normalized = String(value).replace(/[^\d,.-]/g, '');
  if (normalized.indexOf(',') >= 0) {
    normalized = normalized.replace(/\./g, '').replace(',', '.');
  }
  var parsed = Number(normalized);
  return isFinite(parsed) ? parsed : 0;
}

function positive_(value, label) {
  var number = number_(value);
  if (!(number > 0)) throw new Error(label + ' harus lebih besar dari 0.');
  return number;
}

function positiveOrZero_(value, label) {
  var number = number_(value);
  if (number < 0) throw new Error(label + ' tidak boleh negatif.');
  return number;
}

function sum_(rows, key) {
  return rows.reduce(function(total, row) { return total + number_(row[key]); }, 0);
}

function validateRequired_(data, fields) {
  if (!data) throw new Error('Data formulir tidak ditemukan.');
  fields.forEach(function(field) {
    if (data[field] === undefined || data[field] === null || data[field] === '') {
      throw new Error('Mohon lengkapi semua field wajib.');
    }
  });
}

function cleanText_(value, maxLength) {
  var text = String(value === undefined || value === null ? '' : value)
    .replace(/[<>]/g, '')
    .trim()
    .substring(0, maxLength || 500);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function appendNote_(current, note) {
  return current ? String(current) + ' | ' + note : note;
}

function withLock_(callback) {
  ensureSetup();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function buildReceiptPdf_(receipt, settings, fileName) {
  var logo = settings.logo_url
    ? '<img class="logo" src="' + escapeHtml_(settings.logo_url) + '" alt="Logo">'
    : '<div class="logo-placeholder">MIKA</div>';
  var rows = receipt.rows.map(function(row) {
    return '<tr><td>' + escapeHtml_(row[0]) + '</td><td>:</td><td>' + escapeHtml_(row[1]) + '</td></tr>';
  }).join('');
  var html = '<!doctype html><html><head><meta charset="utf-8"><style>' +
    '@page{size:A6;margin:10mm}body{font-family:Arial,sans-serif;color:#172033;font-size:10px}' +
    '.head{text-align:center;border-bottom:2px solid #123b73;padding-bottom:10px;margin-bottom:10px}' +
    '.logo{max-width:44px;max-height:44px}.logo-placeholder{width:44px;height:44px;line-height:44px;' +
    'margin:0 auto 6px;background:#123b73;color:#fff;border-radius:10px;font-weight:700}' +
    'h1{font-size:15px;margin:3px 0;color:#123b73}p{margin:2px 0;color:#596579}' +
    'h2{font-size:11px;text-align:center;margin:12px 0}.meta{width:100%;border-collapse:collapse}' +
    '.meta td{padding:4px 2px;vertical-align:top}.meta td:first-child{width:34%;color:#596579}' +
    '.amount{font-weight:700;color:#123b73}.thanks{text-align:center;border-top:1px dashed #9aa7b8;' +
    'margin-top:14px;padding-top:10px;font-weight:700}.footer{text-align:center;color:#7a8799;margin-top:8px}' +
    '</style></head><body><div class="head">' + logo +
    '<h1>' + escapeHtml_(settings.nama_koperasi || 'Koperasi Syariah MIKA') + '</h1>' +
    '<p>' + escapeHtml_(settings.alamat_koperasi || '') + '</p>' +
    '<p>' + escapeHtml_(settings.telepon_koperasi || '') + ' &bull; ' +
    escapeHtml_(settings.email_koperasi || '') + '</p></div>' +
    '<h2>' + escapeHtml_(receipt.title) + '</h2><table class="meta">' +
    '<tr><td>' + escapeHtml_(receipt.idLabel) + '</td><td>:</td><td><b>' + escapeHtml_(receipt.id) + '</b></td></tr>' +
    '<tr><td>Tanggal</td><td>:</td><td>' + escapeHtml_(receipt.date) + '</td></tr>' +
    rows + '</table><div class="thanks">TERIMA KASIH</div>' +
    '<div class="footer">Amanah, Transparan, Sejahtera</div></body></html>';
  var pdf = Utilities.newBlob(html, MimeType.HTML, 'receipt.html')
    .getAs(MimeType.PDF)
    .setName(fileName);
  return {
    fileName: fileName,
    mimeType: 'application/pdf',
    base64: Utilities.base64Encode(pdf.getBytes())
  };
}

function escapeHtml_(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
