// ============================================================
// MPG BOOKING — Google Apps Script Backend  (v4 — Updated)
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Go to script.google.com → New Project → Name it "MPG Booking Backend"
// 2. Paste this entire file as Code.gs
// 3. Run setupSheets() once to initialize the spreadsheet
// 4. Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the Web App URL into index.html → GAS_URL variable
// ============================================================

// ── CONFIGURATION ──────────────────────────────────────────
const SPREADSHEET_ID = '1XqYka_3fQmfu10er2rXMTlaY2-iraWTJ6uoOXkRY4ms';
const ADMIN_DEFAULT_USER  = 'admin.mpg2026!';
const ADMIN_DEFAULT_PASS_HASH = '6550a71fd95f89569371afd8f5ea03c3c134c4a963343d1251683ed3662b1de3'; // SHA-256 of "adminmpg@2026#"
const ADMIN_EMAIL = 'thosapon.som@gmail.com';   // ← Email สำหรับส่ง OTP ยืนยัน Admin

// ── SHEET NAMES ────────────────────────────────────────────
const SHEETS = {
  TRIPS:    'Trips',
  BOOKINGS: 'Bookings',
  ADMINS:   'Users_Admin',
  SETTINGS: 'Settings',
  LOGS:     'Logs',
  OTP:      'OTP_Sessions'   // ← ใหม่: เก็บ OTP สำหรับ 2FA
};

// ── CORS HEADERS ───────────────────────────────────────────
function setCORSHeaders(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── ENTRY POINTS ───────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  const params = {};
  for (const key in e.parameter) {
    const val = e.parameter[key];
    try { params[key] = JSON.parse(val); }
    catch (err) { params[key] = val; }
  }
  const result = handleAction(action, params);
  return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(result)));
}

function doPost(e) {
  let data = {};
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { data = e.parameter || {}; }
  const action = data.action || '';
  const result = handleAction(action, data);
  return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(result)));
}

// ── MAIN ROUTER ────────────────────────────────────────────
function handleAction(action, data) {
  try {
    switch (action) {
      // Public endpoints
      case 'getTrips':           return getTrips(data);
      case 'getTripById':        return getTripById(data);
      case 'getSeats':           return getSeats(data);
      case 'getSettings':        return getPublicSettings();
      case 'createBooking':      return createBooking(data);
      case 'getBookingById':     return getBookingById(data);

      // Admin auth (2-step: login → request OTP → verify OTP)
      case 'adminLogin':         return adminLogin(data);
      case 'adminRequestOTP':    return adminRequestOTP(data);
      case 'adminVerifyOTP':     return adminVerifyOTP(data);

      // Admin endpoints (require token)
      case 'getBookings':        return requireAdmin(data, getBookings);
      case 'updateBooking':      return requireAdmin(data, updateBooking);
      case 'deleteBooking':      return requireAdmin(data, deleteBooking);
      case 'saveTrip':           return requireAdmin(data, saveTrip);
      case 'deleteTrip':         return requireAdmin(data, deleteTrip);
      case 'saveSettings':       return requireAdmin(data, saveSettings);
      case 'getLogs':            return requireAdmin(data, getLogs);
      case 'addLog':             return requireAdmin(data, addLog);
      case 'saveAdminSeats':     return requireAdmin(data, saveAdminSeats);
      case 'getAdminUsers':      return requireAdmin(data, getAdminUsers);
      case 'addAdminUser':       return requireAdmin(data, addAdminUser);
      case 'deleteAdminUser':    return requireAdmin(data, deleteAdminUser);
      case 'setup':              return setupSheets();

      default:
        return { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    Logger.log('Error in handleAction: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ── SPREADSHEET HELPER ─────────────────────────────────────
function getSpreadsheet() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const files = DriveApp.getFilesByName('MPG_Database');
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create('MPG_Database');
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ── SETUP ──────────────────────────────────────────────────
function setupSheets() {
  try {
    const ss = getSpreadsheet();

    // TRIPS sheet — เพิ่ม staffSeatNames column
    let tripsSheet = ss.getSheetByName(SHEETS.TRIPS);
    if (!tripsSheet) {
      tripsSheet = ss.insertSheet(SHEETS.TRIPS);
      tripsSheet.appendRow(['id','name','destination','date','time','price','seats','booked','category','description','image','status','staffSeats','staffSeatNames','createdAt']);
      tripsSheet.getRange(1,1,1,15).setFontWeight('bold').setBackground('#1a1a2a').setFontColor('#e8b84b');
      tripsSheet.appendRow(['T001','Pai Canyon Sunrise','Pai, Mae Hong Son','2025-08-15','05:30',1200,9,0,'mountain','Watch the sunrise over the breathtaking Pai Canyon.','🏔','active','','',new Date().toISOString()]);
      tripsSheet.appendRow(['T002','Railay Beach Escape','Railay, Krabi','2025-08-20','07:00',2800,9,0,'beach','Pristine white sand and emerald waters.','🏖','active','','',new Date().toISOString()]);
      tripsSheet.appendRow(['T003','Doi Inthanon Summit','Chiang Mai','2025-08-22','06:00',1500,9,0,'mountain',"Thailand's highest peak adventure.",'⛰','active','','',new Date().toISOString()]);
    }

    // BOOKINGS sheet — เพิ่ม nickname column
    let bookingsSheet = ss.getSheetByName(SHEETS.BOOKINGS);
    if (!bookingsSheet) {
      bookingsSheet = ss.insertSheet(SHEETS.BOOKINGS);
      bookingsSheet.appendRow(['id','tripId','tripName','nickname','firstname','lastname','phone','email','company','taxId','address','passengers','seats','slipData','totalAmount','status','note','createdAt','updatedAt']);
      bookingsSheet.getRange(1,1,1,19).setFontWeight('bold').setBackground('#1a1a2a').setFontColor('#e8b84b');
    } else {
      // Migration: add nickname column if missing
      const headers = bookingsSheet.getRange(1,1,1,bookingsSheet.getLastColumn()).getValues()[0];
      if (!headers.includes('nickname')) {
        const insertCol = headers.indexOf('firstname') + 1; // insert before firstname
        bookingsSheet.insertColumnBefore(insertCol);
        bookingsSheet.getRange(1, insertCol).setValue('nickname').setFontWeight('bold').setBackground('#1a1a2a').setFontColor('#e8b84b');
      }
    }

    // ADMINS sheet
    let adminsSheet = ss.getSheetByName(SHEETS.ADMINS);
    if (!adminsSheet) {
      adminsSheet = ss.insertSheet(SHEETS.ADMINS);
      adminsSheet.appendRow(['username','passwordHash','role','email','createdAt','lastLogin','active']);
      adminsSheet.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1a1a2a').setFontColor('#e8b84b');
      adminsSheet.appendRow([ADMIN_DEFAULT_USER, ADMIN_DEFAULT_PASS_HASH, 'superadmin', ADMIN_EMAIL, new Date().toISOString(), '', true]);
    } else {
      // Update default admin email if still old value
      const rowNum = findRowByField(adminsSheet, 'username', ADMIN_DEFAULT_USER);
      if (rowNum > 0) {
        const headers = getHeaders(adminsSheet);
        const emailCol = headers.indexOf('email') + 1;
        const roleCol  = headers.indexOf('role') + 1;
        if (emailCol > 0) adminsSheet.getRange(rowNum, emailCol).setValue(ADMIN_EMAIL);
        if (roleCol  > 0) adminsSheet.getRange(rowNum, roleCol).setValue('superadmin');
      }
    }

    // SETTINGS sheet
    let settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet(SHEETS.SETTINGS);
      settingsSheet.appendRow(['key','value']);
      settingsSheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a1a2a').setFontColor('#e8b84b');
      [['bankName',''],['bankAccountName',''],['bankAccountNumber',''],['qrImage',''],
       ['companyName','MPG Booking'],['companyTaxId',''],['companyAddress',''],['companyPhone','']
      ].forEach(row => settingsSheet.appendRow(row));
    }

    // LOGS sheet
    let logsSheet = ss.getSheetByName(SHEETS.LOGS);
    if (!logsSheet) {
      logsSheet = ss.insertSheet(SHEETS.LOGS);
      logsSheet.appendRow(['timestamp','admin','action','detail','ip']);
      logsSheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a1a2a').setFontColor('#e8b84b');
    }

    // OTP sheet ← ใหม่
    let otpSheet = ss.getSheetByName(SHEETS.OTP);
    if (!otpSheet) {
      otpSheet = ss.insertSheet(SHEETS.OTP);
      otpSheet.appendRow(['username','otp','expires','used']);
      otpSheet.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1a1a2a').setFontColor('#e8b84b');
    }

    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

    return { success: true, message: 'Sheets set up successfully!', spreadsheetUrl: ss.getUrl() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── SHEET → ARRAY HELPERS ──────────────────────────────────
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function findRowByField(sheet, field, value) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx = headers.indexOf(field);
  if (colIdx === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIdx]) === String(value)) return i + 1;
  }
  return -1;
}

function getHeaders(sheet) {
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
}

function updateRow(sheet, rowNum, data) {
  const headers = getHeaders(sheet);
  const rowValues = headers.map(h => (data[h] !== undefined ? data[h] : sheet.getRange(rowNum, headers.indexOf(h)+1).getValue()));
  sheet.getRange(rowNum,1,1,headers.length).setValues([rowValues]);
}

// ── ID GENERATOR ───────────────────────────────────────────
function generateId(prefix) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let id = prefix + '-';
  for (let i = 0; i < 6; i++) id += chars[Math.floor(Math.random() * chars.length)];
  return id;
}

// ── HASH ────────────────────────────────────────────────────
function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ── AUTH: STEP 1 — verify username/password ─────────────────
// Returns { success, needsOTP } — does NOT grant full session yet.
function adminLogin(data) {
  const { username, passwordHash } = data;
  if (!username || !passwordHash) return { success: false, error: 'Missing credentials' };

  const sheet = getSheet(SHEETS.ADMINS);
  const admins = sheetToObjects(sheet);
  const admin = admins.find(a =>
    String(a.username) === String(username) &&
    String(a.passwordHash) === String(passwordHash) &&
    a.active !== false && String(a.active) !== 'FALSE'
  );

  if (!admin) return { success: false, error: 'Invalid username or password' };

  // Return partial success — frontend will call adminRequestOTP next
  return { success: true, needsOTP: true, username, role: admin.role, email: admin.email };
}

// ── AUTH: STEP 2 — send OTP to admin's email ─────────────────
function adminRequestOTP(data) {
  const { username } = data;
  if (!username) return { success: false, error: 'Missing username' };

  const sheet = getSheet(SHEETS.ADMINS);
  const admins = sheetToObjects(sheet);
  const admin = admins.find(a => String(a.username) === String(username));
  if (!admin) return { success: false, error: 'Admin not found' };

  const otp = String(Math.floor(100000 + Math.random() * 900000)); // 6-digit OTP
  const expires = new Date().getTime() + 10 * 60 * 1000; // 10 นาที

  // Store OTP in sheet
  const otpSheet = getSheet(SHEETS.OTP);
  // Remove any previous OTP for this user
  const existing = findRowByField(otpSheet, 'username', username);
  if (existing > 0) otpSheet.deleteRow(existing);
  otpSheet.appendRow([username, otp, expires, false]);

  // Send OTP email
  const email = admin.email || ADMIN_EMAIL;
  try {
    MailApp.sendEmail({
      to: email,
      subject: '🔐 MPG Admin — รหัสยืนยันการเข้าสู่ระบบ',
      htmlBody: `
<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#f5f5f5;padding:20px;">
<div style="max-width:400px;margin:0 auto;background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.1);">
  <div style="background:#0a0a0f;padding:20px;text-align:center;">
    <h2 style="color:#e8b84b;margin:0;font-size:22px;">MPG BOOKING</h2>
    <p style="color:#888;font-size:12px;margin:4px 0 0;">Admin Verification Code</p>
  </div>
  <div style="padding:24px;text-align:center;">
    <p style="color:#333;margin-bottom:8px;">รหัส OTP ของคุณสำหรับเข้าสู่ระบบ Admin:</p>
    <div style="font-size:40px;font-weight:bold;letter-spacing:12px;color:#c49a35;background:#fff8e8;border:2px dashed #e8b84b;border-radius:12px;padding:16px;margin:16px 0;">${otp}</div>
    <p style="color:#888;font-size:12px;">รหัสนี้จะหมดอายุใน <strong>10 นาที</strong></p>
    <p style="color:#c0003a;font-size:11px;">หากคุณไม่ได้ขอรหัสนี้ โปรดละเว้นอีเมลฉบับนี้</p>
  </div>
</div>
</body></html>`
    });
  } catch (e) {
    Logger.log('OTP email failed: ' + e);
    return { success: false, error: 'Failed to send OTP email: ' + e.toString() };
  }

  addLogEntry('admin', username, 'OTP Requested', 'OTP sent to ' + email);
  return { success: true, message: 'OTP sent to ' + email, maskedEmail: maskEmail(email) };
}

function maskEmail(email) {
  if (!email || !email.includes('@')) return email;
  const [user, domain] = email.split('@');
  const visible = user.length > 2 ? user.substring(0, 2) : user.substring(0, 1);
  return visible + '***@' + domain;
}

// ── AUTH: STEP 3 — verify OTP & grant session token ──────────
function adminVerifyOTP(data) {
  const { username, otp } = data;
  if (!username || !otp) return { success: false, error: 'Missing data' };

  const otpSheet = getSheet(SHEETS.OTP);
  const rows = sheetToObjects(otpSheet);
  const record = rows.find(r => String(r.username) === String(username) && String(r.otp) === String(otp));

  if (!record) return { success: false, error: 'Invalid OTP code' };
  if (String(record.used) === 'true' || record.used === true) return { success: false, error: 'OTP already used' };
  if (new Date().getTime() > Number(record.expires)) return { success: false, error: 'OTP expired. Please request a new one.' };

  // Mark OTP as used
  const rowNum = findRowByField(otpSheet, 'username', username);
  if (rowNum > 0) {
    const headers = getHeaders(otpSheet);
    const usedCol = headers.indexOf('used') + 1;
    if (usedCol > 0) otpSheet.getRange(rowNum, usedCol).setValue(true);
  }

  // Get admin info
  const adminSheet = getSheet(SHEETS.ADMINS);
  const admins = sheetToObjects(adminSheet);
  const admin = admins.find(a => String(a.username) === String(username));
  if (!admin) return { success: false, error: 'Admin not found' };

  // Update last login
  const adminRow = findRowByField(adminSheet, 'username', username);
  if (adminRow > 0) {
    const headers = getHeaders(adminSheet);
    const lastLoginCol = headers.indexOf('lastLogin') + 1;
    if (lastLoginCol > 0) adminSheet.getRange(adminRow, lastLoginCol).setValue(new Date().toISOString());
  }

  // Generate session token
  const token = Utilities.base64Encode(username + ':' + new Date().getTime() + ':' + Math.random());
  PropertiesService.getScriptProperties().setProperty(
    'session_' + token,
    JSON.stringify({ username, role: admin.role, expires: new Date().getTime() + 8 * 3600 * 1000 })
  );

  addLogEntry('admin', username, 'Login Success', 'Admin logged in after OTP verification');
  return { success: true, token, user: { username, role: admin.role, email: admin.email } };
}

// ── REQUIRE ADMIN ──────────────────────────────────────────
function requireAdmin(data, fn) {
  const token = data.token || '';
  if (token) {
    const session = validateToken(token);
    if (!session) return { success: false, error: 'Session expired. Please login again.' };
  }
  return fn(data);
}

function validateToken(token) {
  if (!token) return null;
  try {
    const props = PropertiesService.getScriptProperties();
    const sessionStr = props.getProperty('session_' + token);
    if (!sessionStr) return null;
    const session = JSON.parse(sessionStr);
    if (session.expires < new Date().getTime()) {
      props.deleteProperty('session_' + token);
      return null;
    }
    return session;
  } catch (e) { return null; }
}

// ── TRIPS CRUD ─────────────────────────────────────────────
function getTrips(data) {
  const sheet = getSheet(SHEETS.TRIPS);
  const trips = sheetToObjects(sheet);
  const active = trips.filter(t => t.status !== 'inactive');
  return { success: true, data: active.map(formatTrip) };
}

function getTripById(data) {
  const sheet = getSheet(SHEETS.TRIPS);
  const trips = sheetToObjects(sheet);
  const trip = trips.find(t => t.id === data.id);
  if (!trip) return { success: false, error: 'Trip not found' };
  return { success: true, data: formatTrip(trip) };
}

function formatTrip(t) {
  let staffSeats = [];
  let staffSeatNames = {};
  try { staffSeats = JSON.parse(t.staffSeats || '[]'); } catch(e) {}
  try { staffSeatNames = typeof t.staffSeatNames === 'string' ? JSON.parse(t.staffSeatNames || '{}') : (t.staffSeatNames || {}); } catch(e) {}
  return {
    id: String(t.id), name: String(t.name || ''), destination: String(t.destination || ''),
    date: formatDateForClient(t.date), time: String(t.time || ''),
    price: parseFloat(t.price) || 0, seats: parseInt(t.seats) || 9,
    booked: parseInt(t.booked) || 0, category: String(t.category || 'mountain'),
    description: String(t.description || ''), image: String(t.image || ''),
    status: String(t.status || 'active'), staffSeats, staffSeatNames
  };
}

function formatDateForClient(dateVal) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(dateVal);
}

function saveTrip(data) {
  const sheet = getSheet(SHEETS.TRIPS);
  if (data.id) {
    const rowNum = findRowByField(sheet, 'id', data.id);
    if (rowNum < 0) return { success: false, error: 'Trip not found' };
    const headers = getHeaders(sheet);
    const currentRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const current = {};
    headers.forEach((h, i) => current[h] = currentRow[i]);
    const updated = {
      ...current,
      name: data.name || current.name,
      destination: data.destination || current.destination,
      date: data.date || current.date,
      time: data.time || current.time,
      price: data.price !== undefined ? data.price : current.price,
      seats: data.seats !== undefined ? data.seats : current.seats,
      category: data.category || current.category,
      description: data.description !== undefined ? data.description : current.description,
      image: data.image !== undefined ? data.image : current.image,
      status: data.status || current.status,
    };
    const rowValues = headers.map(h => updated[h] !== undefined ? updated[h] : '');
    sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowValues]);
    addLogEntry('admin', data.adminUser || 'system', 'Trip Updated', 'Updated trip: ' + data.name);
    return { success: true };
  } else {
    const id = 'T' + String(sheet.getLastRow()).padStart(3, '0');
    const headers = getHeaders(sheet);
    const newTrip = {
      id, name: data.name, destination: data.destination, date: data.date,
      time: data.time || '07:00', price: data.price, seats: data.seats || 9,
      booked: 0, category: data.category || 'mountain',
      description: data.description || '', image: data.image || '',
      status: data.status || 'active', staffSeats: '[]', staffSeatNames: '{}',
      createdAt: new Date().toISOString()
    };
    sheet.appendRow(headers.map(h => newTrip[h] !== undefined ? newTrip[h] : ''));
    addLogEntry('admin', data.adminUser || 'system', 'Trip Added', 'Added trip: ' + data.name);
    return { success: true, id };
  }
}

function deleteTrip(data) {
  const sheet = getSheet(SHEETS.TRIPS);
  const rowNum = findRowByField(sheet, 'id', data.id);
  if (rowNum < 0) return { success: false, error: 'Trip not found' };
  sheet.deleteRow(rowNum);
  addLogEntry('admin', data.adminUser || 'system', 'Trip Deleted', 'Deleted trip ID: ' + data.id);
  return { success: true };
}

// ── SEATS ──────────────────────────────────────────────────
function getSeats(data) {
  const tripId = data.tripId;
  if (!tripId) return { success: false, error: 'No tripId' };

  const bookingSheet = getSheet(SHEETS.BOOKINGS);
  const bookings = sheetToObjects(bookingSheet);

  const booked = [];
  // Build map: seatNumber → nickname (or name) for display
  const seatNameMap = {};
  bookings
    .filter(b => b.tripId === tripId && b.status !== 'cancelled')
    .forEach(b => {
      try {
        const seats = typeof b.seats === 'string' ? JSON.parse(b.seats) : (Array.isArray(b.seats) ? b.seats : []);
        const displayName = b.nickname ? String(b.nickname) : ((b.firstname || '') + ' ' + (b.lastname ? b.lastname.charAt(0) + '.' : '')).trim();
        seats.forEach((s, idx) => {
          booked.push(parseInt(s));
          seatNameMap[parseInt(s)] = idx === 0 ? displayName : displayName + ' (' + (idx+1) + ')';
        });
      } catch (e) {}
    });

  const tripSheet = getSheet(SHEETS.TRIPS);
  const trips = sheetToObjects(tripSheet);
  const trip = trips.find(t => t.id === tripId);
  let staff = [];
  let staffSeatNames = {};
  if (trip) {
    try { staff = JSON.parse(trip.staffSeats || '[]'); } catch(e) {}
    try { staffSeatNames = typeof trip.staffSeatNames === 'string' ? JSON.parse(trip.staffSeatNames || '{}') : {}; } catch(e) {}
  }

  return { success: true, booked: [...new Set(booked)], staff, seatNameMap, staffSeatNames };
}

// ── saveAdminSeats — now also saves custom seat names ───────
function saveAdminSeats(data) {
  const { tripId, staffSeats, staffSeatNames } = data;
  const sheet = getSheet(SHEETS.TRIPS);
  const rowNum = findRowByField(sheet, 'id', tripId);
  if (rowNum < 0) return { success: false, error: 'Trip not found' };
  const headers = getHeaders(sheet);

  const staffSeatsCol    = headers.indexOf('staffSeats') + 1;
  const staffNamesCol    = headers.indexOf('staffSeatNames') + 1;

  if (staffSeatsCol > 0) sheet.getRange(rowNum, staffSeatsCol).setValue(JSON.stringify(staffSeats || []));
  if (staffNamesCol > 0) sheet.getRange(rowNum, staffNamesCol).setValue(JSON.stringify(staffSeatNames || {}));

  addLogEntry('admin', data.adminUser || 'system', 'Seats Updated', 'Staff seats updated for trip ' + tripId);
  return { success: true };
}

// ── BOOKINGS CRUD ──────────────────────────────────────────
function createBooking(data) {
  const sheet = getSheet(SHEETS.BOOKINGS);

  // Validate seats
  const seatCheck = getSeats({ tripId: data.tripId });
  const requestedSeats = Array.isArray(data.seats) ? data.seats : JSON.parse(data.seats || '[]');
  const conflict = requestedSeats.filter(s => seatCheck.booked.includes(parseInt(s)));
  if (conflict.length > 0) {
    return { success: false, error: 'Seats ' + conflict.join(', ') + ' are already booked. Please choose different seats.' };
  }

  const bookingId = generateId('BK');
  const now = new Date().toISOString();
  const headers = getHeaders(sheet);
  const booking = {
    id: bookingId,
    tripId: data.tripId || '',
    tripName: data.tripName || '',
    nickname: data.nickname || '',   // ← ใหม่
    firstname: data.firstname || '',
    lastname: data.lastname || '',
    phone: data.phone || '',
    email: data.email || '',
    company: data.company || '',
    taxId: data.taxId || '',
    address: data.address || '',
    passengers: data.passengers || 1,
    seats: JSON.stringify(requestedSeats),
    slipData: data.slipData || '',
    totalAmount: data.totalAmount || 0,
    status: 'pending',
    note: '',
    createdAt: now,
    updatedAt: now
  };

  sheet.appendRow(headers.map(h => booking[h] !== undefined ? booking[h] : ''));

  // Update booked count on Trip
  const tripSheet = getSheet(SHEETS.TRIPS);
  const tripRow = findRowByField(tripSheet, 'id', data.tripId);
  if (tripRow > 0) {
    const tripHeaders = getHeaders(tripSheet);
    const bookedCol = tripHeaders.indexOf('booked') + 1;
    if (bookedCol > 0) {
      const currentBooked = parseInt(tripSheet.getRange(tripRow, bookedCol).getValue()) || 0;
      const newBooked = currentBooked + requestedSeats.length;
      tripSheet.getRange(tripRow, bookedCol).setValue(newBooked);
      // Auto-mark trip as full
      const seatsCol  = tripHeaders.indexOf('seats') + 1;
      const statusCol = tripHeaders.indexOf('status') + 1;
      if (seatsCol > 0 && statusCol > 0) {
        const totalSeats = parseInt(tripSheet.getRange(tripRow, seatsCol).getValue()) || 9;
        if (newBooked >= totalSeats) tripSheet.getRange(tripRow, statusCol).setValue('full');
      }
    }
  }

  // Send confirmation email
  if (data.email && data.email.includes('@')) {
    try { sendConfirmationEmail(booking); } catch (e) { Logger.log('Email failed: ' + e); }
  }

  addLogEntry('user', data.phone, 'Booking Created', 'New booking: ' + bookingId + ' for trip ' + data.tripId + (data.nickname ? ' (ชื่อเล่น: ' + data.nickname + ')' : ''));
  return { success: true, bookingId, booking: { ...booking, seats: requestedSeats } };
}

function getBookings(data) {
  const sheet = getSheet(SHEETS.BOOKINGS);
  const bookings = sheetToObjects(sheet);
  return {
    success: true,
    data: bookings.map(b => ({
      ...b,
      seats: (() => { try { return JSON.parse(b.seats || '[]'); } catch(e) { return []; } })()
    }))
  };
}

function getBookingById(data) {
  const sheet = getSheet(SHEETS.BOOKINGS);
  const bookings = sheetToObjects(sheet);
  const booking = bookings.find(b => b.id === data.id);
  if (!booking) return { success: false, error: 'Booking not found' };
  booking.seats = (() => { try { return JSON.parse(booking.seats || '[]'); } catch(e) { return []; } })();
  return { success: true, data: booking };
}

function updateBooking(data) {
  const sheet = getSheet(SHEETS.BOOKINGS);
  const rowNum = findRowByField(sheet, 'id', data.id);
  if (rowNum < 0) return { success: false, error: 'Booking not found' };

  const headers = getHeaders(sheet);
  const currentRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const current = {};
  headers.forEach((h, i) => current[h] = currentRow[i]);

  const updated = { ...current, ...data, updatedAt: new Date().toISOString() };
  if (data.seats && Array.isArray(data.seats)) updated.seats = JSON.stringify(data.seats);
  const rowValues = headers.map(h => updated[h] !== undefined ? updated[h] : '');
  sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowValues]);
  addLogEntry('admin', data.adminUser || 'system', 'Booking Updated', 'Updated booking: ' + data.id + ' status=' + (data.status || current.status));
  return { success: true };
}

function deleteBooking(data) {
  const sheet = getSheet(SHEETS.BOOKINGS);
  const rowNum = findRowByField(sheet, 'id', data.id);
  if (rowNum < 0) return { success: false, error: 'Not found' };
  sheet.deleteRow(rowNum);
  addLogEntry('admin', data.adminUser || 'system', 'Booking Deleted', 'Deleted booking: ' + data.id);
  return { success: true };
}

// ── SETTINGS ───────────────────────────────────────────────
function getPublicSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  const rows = sheetToObjects(sheet);
  const settings = {};
  rows.forEach(r => { settings[r.key] = r.value; });
  return {
    success: true,
    data: {
      bankName: settings.bankName || '',
      bankAccountName: settings.bankAccountName || '',
      bankAccountNumber: settings.bankAccountNumber || '',
      qrImage: settings.qrImage || '',
      companyName: settings.companyName || 'MPG Booking',
      companyTaxId: settings.companyTaxId || '',
      companyAddress: settings.companyAddress || '',
      companyPhone: settings.companyPhone || ''
    }
  };
}

function saveSettings(data) {
  const sheet = getSheet(SHEETS.SETTINGS);
  const allowedKeys = ['bankName','bankAccountName','bankAccountNumber','qrImage','companyName','companyTaxId','companyAddress','companyPhone'];
  allowedKeys.forEach(key => {
    if (data[key] !== undefined) {
      const rowNum = findRowByField(sheet, 'key', key);
      if (rowNum > 0) {
        sheet.getRange(rowNum, 2).setValue(data[key]);
      } else {
        sheet.appendRow([key, data[key]]);
      }
    }
  });
  addLogEntry('admin', data.adminUser || 'system', 'Settings Updated', 'Payment/company settings updated');
  return { success: true };
}

// ── ADMIN USERS ────────────────────────────────────────────
function getAdminUsers(data) {
  const sheet = getSheet(SHEETS.ADMINS);
  const admins = sheetToObjects(sheet);
  return {
    success: true,
    data: admins.map(a => ({
      username: a.username,
      role: a.role,
      email: a.email,
      lastLogin: a.lastLogin,
      active: a.active,
      createdAt: a.createdAt
    }))
  };
}

function addAdminUser(data) {
  const { username, passwordHash, role, email } = data;
  if (!username || !passwordHash) return { success: false, error: 'Username and password required' };
  const sheet = getSheet(SHEETS.ADMINS);
  const existing = findRowByField(sheet, 'username', username);
  if (existing > 0) return { success: false, error: 'Username already exists' };
  sheet.appendRow([username, passwordHash, role || 'staff', email || '', new Date().toISOString(), '', true]);
  addLogEntry('admin', data.adminUser || 'system', 'Admin Added', 'New admin user: ' + username + ' role: ' + (role || 'staff'));
  return { success: true };
}

function deleteAdminUser(data) {
  const sheet = getSheet(SHEETS.ADMINS);
  // Protect superadmin
  if (String(data.username) === ADMIN_DEFAULT_USER) return { success: false, error: 'Cannot delete the main superadmin account' };
  const rowNum = findRowByField(sheet, 'username', data.username);
  if (rowNum < 0) return { success: false, error: 'User not found' };
  sheet.deleteRow(rowNum);
  addLogEntry('admin', data.adminUser || 'system', 'Admin Deleted', 'Deleted admin user: ' + data.username);
  return { success: true };
}

// ── LOGS ───────────────────────────────────────────────────
function getLogs(data) {
  const sheet = getSheet(SHEETS.LOGS);
  const logs = sheetToObjects(sheet);
  return { success: true, data: logs.slice(-200) };
}

function addLog(data) {
  addLogEntry(data.type || 'admin', data.admin || 'system', data.action || '', data.detail || '');
  return { success: true };
}

function addLogEntry(type, user, action, detail) {
  try {
    const sheet = getSheet(SHEETS.LOGS);
    sheet.appendRow([new Date().toISOString(), user, action, detail, type]);
  } catch (e) { Logger.log('Log error: ' + e); }
}

// ── EMAIL: Booking Confirmation ─────────────────────────────
function sendConfirmationEmail(booking) {
  const settings = getPublicSettings().data;
  const seats = (() => { try { return JSON.parse(booking.seats || '[]'); } catch(e) { return []; } })();
  const subject = `✅ Booking Confirmed — ${booking.id} | MPG Booking`;
  const body = `
<!DOCTYPE html><html><head><style>
  body { font-family: Arial, sans-serif; background:#f5f5f5; margin:0; padding:20px; }
  .container { max-width:500px; margin:0 auto; background:white; border-radius:12px; overflow:hidden; box-shadow:0 4px 20px rgba(0,0,0,0.1); }
  .header { background:#0a0a0f; padding:24px; text-align:center; }
  .header h1 { color:#e8b84b; font-size:28px; margin:0; }
  .header p { color:#8a8a9a; font-size:12px; margin:4px 0 0; }
  .body { padding:24px; }
  .booking-id { background:#f9f5e8; border:1px solid #e8b84b; border-radius:8px; padding:16px; text-align:center; margin-bottom:20px; }
  .booking-id .label { font-size:11px; color:#888; text-transform:uppercase; letter-spacing:0.06em; }
  .booking-id .id { font-size:24px; color:#c49a35; font-weight:bold; font-family:monospace; }
  .row { display:flex; justify-content:space-between; padding:8px 0; border-bottom:1px solid #f0f0f0; font-size:13px; }
  .row .label { color:#888; }
  .row .value { font-weight:500; color:#333; }
  .total { background:#0a0a0f; padding:16px; display:flex; justify-content:space-between; margin-top:16px; border-radius:8px; }
  .total .label { color:#8a8a9a; font-size:13px; }
  .total .amount { color:#e8b84b; font-size:22px; font-weight:bold; }
  .footer { padding:16px 24px; background:#fafafa; text-align:center; font-size:11px; color:#aaa; }
</style></head><body>
<div class="container">
  <div class="header"><h1>MPG BOOKING</h1><p>Booking Confirmation</p></div>
  <div class="body">
    <p>สวัสดี <strong>${booking.nickname || booking.firstname}</strong> 👋</p>
    <p>ได้รับการจองของคุณแล้ว รอการยืนยันจากเจ้าหน้าที่</p>
    <div class="booking-id">
      <div class="label">Booking ID ของคุณ</div>
      <div class="id">${booking.id}</div>
    </div>
    <div class="row"><span class="label">ทริป</span><span class="value">${booking.tripName}</span></div>
    <div class="row"><span class="label">ชื่อเล่น</span><span class="value">${booking.nickname || '—'}</span></div>
    <div class="row"><span class="label">ชื่อ-นามสกุล</span><span class="value">${booking.firstname} ${booking.lastname}</span></div>
    <div class="row"><span class="label">โทรศัพท์</span><span class="value">${booking.phone}</span></div>
    <div class="row"><span class="label">จำนวนผู้โดยสาร</span><span class="value">${booking.passengers}</span></div>
    <div class="row"><span class="label">ที่นั่ง</span><span class="value">${seats.join(', ')}</span></div>
    <div class="row"><span class="label">สถานะ</span><span class="value">⏳ รอยืนยัน</span></div>
    <div class="total">
      <span class="label">ยอดชำระ</span>
      <span class="amount">฿${Number(booking.totalAmount).toLocaleString()}</span>
    </div>
  </div>
  <div class="footer">${settings.companyName}${settings.companyPhone ? ' · ' + settings.companyPhone : ''}<br>เก็บ Booking ID ไว้ใช้อ้างอิงในภายหลัง</div>
</div>
</body></html>`;
  MailApp.sendEmail({ to: booking.email, subject, htmlBody: body });
}

// ── MANUAL TESTS ────────────────────────────────────────────
function testSetup() {
  const result = setupSheets();
  Logger.log(JSON.stringify(result));
}

function testLogin() {
  const hash = hashPassword('adminmpg@2026#');
  Logger.log('Hash: ' + hash);
  const result = adminLogin({ username: 'admin.mpg2026!', passwordHash: hash });
  Logger.log(JSON.stringify(result));
}

function testHashGen() {
  // Run this to get the hash of a new password
  const pw = 'adminmpg@2026#';
  Logger.log('SHA-256 of "' + pw + '": ' + hashPassword(pw));
}
