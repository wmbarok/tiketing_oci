// =========================================================================
// GOOGLE APPS SCRIPT BACKEND - PT. OUTVENTURE CAKRAWALA INDONESIA
// =========================================================================
// SETUP:
// 1. Paste kode ini di Code.gs
// 2. Pilih fungsi "createSheets" di dropdown → Klik Run ▶️ → Allow
// 3. Deploy > Manage deployments > New version > Deploy
//    Execute as: Me | Who has access: Anyone
// =========================================================================

// ==================== AUTO-CREATE / MIGRATE SHEETS ====================
function createSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  createFlightSheet(ss);
  createHotelSheet(ss);
  createUsersSheet(ss);
  var def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) {
    try { ss.deleteSheet(def); } catch(e) {}
  }
  try {
    SpreadsheetApp.getUi().alert(
      "✅ SETUP / MIGRASI BERHASIL!\n\n" +
      "Sheet:\n• FLIGHT_BOOKINGS\n• HOTEL_BOOKINGS\n• USERS\n\n" +
      "Kolom hotel baru: BookingRefNo, Meal, Nationality, Channel\n" +
      "Kolom pembayaran: PaymentStatus & PaidAt\n\n" +
      "Lanjut: Deploy Web App (Execute as: Me | Access: Anyone)"
    );
  } catch(e) {}
}

function createFlightSheet(ss) {
  var sheet = ss.getSheetByName("FLIGHT_BOOKINGS") || ss.insertSheet("FLIGHT_BOOKINGS");
  ensureHeaders(sheet, [
    "Invoice","PNR","IssueDate","Client","Pax","TicketNo",
    "Airline","DepDate","Time","Cabin","DepAirport",
    "ArrAirport","Route","CabinBag","CheckedBag",
    "BasePrice","AddonPrice","TotalPrice","CreatedAt",
    "PaymentStatus","PaidAt"
  ], "#1e40af");
  return sheet;
}

function createHotelSheet(ss) {
  var sheet = ss.getSheetByName("HOTEL_BOOKINGS") || ss.insertSheet("HOTEL_BOOKINGS");
  ensureHeaders(sheet, [
    "Invoice","Date","PaymentMethod","Client","Pax",
    "HotelName","RoomType","Duration","StayPeriod",
    "PaxQty","RoomQty","UnitPrice","Voucher",
    "GrandTotal","CreatedAt",
    "PaymentStatus","PaidAt",
    "BookingRefNo","Meal","Nationality","Channel"
  ], "#065f46");
  return sheet;
}

function createUsersSheet(ss) {
  var sheet = ss.getSheetByName("USERS") || ss.insertSheet("USERS");
  ensureHeaders(sheet, ["Username","Password","Name","Role","CreatedAt"], "#7c3aed");
  if (sheet.getLastRow() <= 1) {
    var now = new Date();
    sheet.appendRow(["admin","admin123","Master Admin","admin",now]);
    sheet.appendRow(["coadmin","coadmin123","Co-Admin Operational","coadmin",now]);
    sheet.appendRow(["viewer","viewer123","Guest Viewer","viewer",now]);
  }
  return sheet;
}

function ensureHeaders(sheet, headers, color) {
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var isEmpty = existing.length === 0 || existing.every(function(v){ return v === "" || v === null; });

  if (isEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var missing = [];
    for (var i = 0; i < headers.length; i++) {
      if (existing.indexOf(headers[i]) === -1) missing.push(headers[i]);
    }
    if (missing.length > 0) {
      sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    }
  }

  var finalCol = sheet.getLastColumn();
  sheet.getRange(1, 1, 1, finalCol)
    .setFontWeight("bold").setBackground(color || "#1e40af")
    .setFontColor("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 32);
  for (var c = 1; c <= finalCol; c++) sheet.autoResizeColumn(c);
  return sheet;
}

// ==================== WEB APP ====================
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getData') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return jsonOut({
      status: "success",
      flights: sheetToJson(getOrCreateSheet(ss, "FLIGHT_BOOKINGS")),
      hotels:  sheetToJson(getOrCreateSheet(ss, "HOTEL_BOOKINGS")),
      users:   sheetToJson(getOrCreateSheet(ss, "USERS"))
    });
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PT. Outventure Cakrawala Indonesia - Sistem Tiketing & Travel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function doPost(e) {
  try {
    var data;
    if (e.parameter && e.parameter.data) data = JSON.parse(e.parameter.data);
    else if (e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
    else return jsonOut({status:"error", message:"No data received"});

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = data.action || "create";

    if (data.type === "FLIGHT" || data.type === "HOTEL") {
      var sheetName = data.type === "FLIGHT" ? "FLIGHT_BOOKINGS" : "HOTEL_BOOKINGS";
      var sheet = getOrCreateSheet(ss, sheetName);

      if (action === "markPayment") return markPaymentAction(sheet, data);

      var row = data.type === "FLIGHT" ? flightRowFromData(data) : hotelRowFromData(data);

      if (action === "update") {
        var key = data._originalInvoice || data.invoice;
        if (!updateRowByInvoice(sheet, key, row)) {
          return jsonOut({status:"error", message:"Invoice tidak ditemukan: " + key});
        }
      } else if (action === "delete") {
        deleteRowByInvoice(sheet, data.invoice);
      } else {
        row.push("UNPAID", "");
        sheet.appendRow(row);
      }
    } else if (data.type === "USER") {
      var uSheet = getOrCreateSheet(ss, "USERS");
      if (action === "create") {
        if (findUserRow(uSheet, data.username) > 0) return jsonOut({status:"error", message:"Username sudah dipakai"});
        uSheet.appendRow([data.username, data.password, data.name, (data.role||"viewer").toLowerCase(), new Date()]);
      } else if (action === "update") {
        var r = findUserRow(uSheet, data._originalUsername || data.username);
        if (r < 1) return jsonOut({status:"error", message:"User tidak ditemukan"});
        var createdAt = uSheet.getRange(r, 5).getValue() || new Date();
        uSheet.getRange(r, 1, 1, 5).setValues([[data.username, data.password, data.name, (data.role||"viewer").toLowerCase(), createdAt]]);
      } else if (action === "delete") {
        var r2 = findUserRow(uSheet, data.username);
        if (r2 > 0) uSheet.deleteRow(r2);
      } else if (action === "changePassword") {
        var r3 = findUserRow(uSheet, data.username);
        if (r3 < 1) return jsonOut({status:"error", message:"User tidak ditemukan"});
        if (String(uSheet.getRange(r3, 2).getValue()) !== String(data.oldPassword||"")) {
          return jsonOut({status:"error", message:"Password lama salah"});
        }
        uSheet.getRange(r3, 2).setValue(data.newPassword||"");
      }
    }

    return jsonOut({status:"success", message:"Data tersimpan!"});
  } catch (err) {
    return jsonOut({status:"error", message: err.toString()});
  }
}

// ==================== MARK PAYMENT ====================
function markPaymentAction(sheet, data) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var statusIdx = headers.indexOf("PaymentStatus") + 1;
  var paidAtIdx = headers.indexOf("PaidAt") + 1;
  if (statusIdx < 1 || paidAtIdx < 1) {
    return jsonOut({status:"error", message:"Kolom PaymentStatus/PaidAt tidak ada. Jalankan createSheets."});
  }

  var invoices = (data.invoices && data.invoices.length) ? data.invoices : [data.invoice].filter(Boolean);
  if (invoices.length === 0) return jsonOut({status:"error", message:"Tidak ada invoice dipilih"});

  var status = data.status === "UNPAID" ? "UNPAID" : "PAID";
  var values = sheet.getDataRange().getValues();
  var now = new Date();
  var count = 0;

  for (var i = 1; i < values.length; i++) {
    var inv = String(values[i][0]).trim();
    if (invoices.indexOf(inv) !== -1) {
      sheet.getRange(i + 1, statusIdx).setValue(status);
      sheet.getRange(i + 1, paidAtIdx).setValue(status === "PAID" ? now : "");
      count++;
    }
  }
  return jsonOut({status:"success", message: count + " invoice diupdate", count: count});
}

function findUserRow(sheet, username) {
  var v = sheet.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) return i + 1;
  }
  return -1;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function flightRowFromData(d) {
  return [
    d.invoice||"", d.pnr||"", d.issueDate||"", d.client||"", d.pax||"", d.ticketNo||"",
    d.airline||"", d.depDate||"", d.time||"", d.cabin||"", d.depAirport||"", d.arrAirport||"",
    d.route||"", d.cabinBag||"", d.checkedBag||"",
    Number(d.basePrice)||0, Number(d.addonPrice)||0, Number(d.totalPrice)||0, new Date()
  ];
}

function hotelRowFromData(d) {
  return [
    d.invoice||"", d.date||"", d.paymentMethod||"", d.client||"", d.pax||"",
    d.hotelName||"", d.roomType||"", d.duration||"", d.stayPeriod||"",
    Number(d.paxQty)||1, Number(d.roomQty)||1,
    Number(d.unitPrice)||0, Number(d.voucher)||0, Number(d.grandTotal)||0, new Date(),
    "", "",
    d.bookingRefNo||"", d.meal||"", d.nationality||"INDONESIA", d.channel||"GLOBAL TIKET NETWORK"
  ];
}

function updateRowByInvoice(sheet, invoice, rowData) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(invoice).trim()) {
      rowData[rowData.length - 1] = values[i][rowData.length - 1] || new Date();
      sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
      return true;
    }
  }
  return false;
}

function deleteRowByInvoice(sheet, invoice) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(invoice).trim()) { sheet.deleteRow(i + 1); return true; }
  }
  return false;
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    if (name === "FLIGHT_BOOKINGS") sheet = createFlightSheet(ss);
    else if (name === "HOTEL_BOOKINGS") sheet = createHotelSheet(ss);
    else if (name === "USERS") sheet = createUsersSheet(ss);
    else sheet = ss.insertSheet(name);
  }
  return sheet;
}

function sheetToJson(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var headers = rows[0];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var v = rows[i][j];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      obj[headers[j]] = v;
    }
    result.push(obj);
  }
  return result;
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🛠️ Outventure Tools')
      .addItem('🔄 Setup / Migrasi Sheets', 'createSheets')
      .addItem('🧹 Bersihkan Data Booking', 'confirmClearData')
      .addToUi();
  } catch(e) {}
}

function confirmClearData() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('⚠️ KONFIRMASI', 'Hapus semua baris booking? (header & user aman)', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ["FLIGHT_BOOKINGS","HOTEL_BOOKINGS"].forEach(function(n){
      var sh = ss.getSheetByName(n);
      if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
    });
    ui.alert('✅ Data booking dibersihkan.');
  }
}
