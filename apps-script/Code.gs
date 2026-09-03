/**
 * MSpace-Time Counter_v2 · MSpace Mlang Google Sheets connector
 *
 * Deploy this file as a Web app from Extensions → Apps Script in the shared
 * MSpace sheet. The web app sends JSON POST requests for check-in/check-out
 * events and polls the snapshot with GET requests for monitoring.
 */
const SPREADSHEET_ID = '1Cs269YewMz6mPNK4ye7qzZMpVfg3rm6NWLsbPVfIS_I';
const FIRST_DATA_ROW = 4;
const LAST_DATA_ROW = 2000;

function doGet(e) {
  try {
    const result = snapshot_();
    result.ok = true;
    return json_(result);
  } catch (error) {
    return json_({ok: false, error: error.message || String(error)});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = payload.action;
    if (action === 'upsertSession') writeCheckIn_(payload.session || {});
    else if (action === 'checkout') writeCheckOut_(payload.session || {});
    else throw new Error('Unsupported action: ' + action);
    SpreadsheetApp.flush();
    const result = snapshot_();
    result.ok = true;
    return json_(result);
  } catch (error) {
    return json_({ok: false, error: error.message || String(error)});
  } finally {
    lock.releaseLock();
  }
}

function writeCheckIn_(session) {
  if (!session.sessionId || !session.uid || !session.customer) throw new Error('Check-in is missing session, RFID, or customer details.');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sessions = ss.getSheetByName('Sessions');
  const customer = session.customer;
  const row = findRow_(sessions, 1, session.sessionId) || firstEmptyRow_(sessions, 1);
  const checkIn = date_(session.checkin);
  sessions.getRange(row, 1, 1, 11).setValues([[
    session.sessionId, checkIn, session.uid, customer.name || '', customer.email || '', customer.address || '',
    customer.phone || '', titleCase_(customer.type || 'regular'), checkIn, '', ''
  ]]);
  sessions.getRange(row, 16, 1, 2).setValues([['Active', session.staff || 'Front desk']]);
  upsertCustomer_(ss.getSheetByName('Customers'), session.uid, customer, checkIn);
}

function writeCheckOut_(session) {
  if (!session.sessionId && !session.uid) throw new Error('Checkout is missing the session or RFID.');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sessions = ss.getSheetByName('Sessions');
  const row = (session.sessionId && findRow_(sessions, 1, session.sessionId)) || findActiveRow_(sessions, session.uid);
  if (!row) throw new Error('No active session found for checkout.');
  const checkout = date_(session.checkout || new Date().toISOString());
  const minutes = Math.max(1, Number(session.minutes) || Math.ceil((checkout - sessions.getRange(row, 9).getValue()) / 60000));
  sessions.getRange(row, 10, 1, 2).setValues([[checkout, minutes]]);
  sessions.getRange(row, 16, 1, 2).setValues([['Paid', session.staff || 'Front desk']]);
}

function upsertCustomer_(sheet, uid, customer, timestamp) {
  const row = findRow_(sheet, 1, uid) || firstEmptyRow_(sheet, 1);
  const existing = row >= FIRST_DATA_ROW ? sheet.getRange(row, 7).getValue() : '';
  sheet.getRange(row, 1, 1, 8).setValues([[
    uid, customer.name || '', customer.email || '', customer.address || '', customer.phone || '',
    titleCase_(customer.type || 'regular'), existing || timestamp, timestamp
  ]]);
}

function snapshot_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sessions = ss.getSheetByName('Sessions');
  const customers = ss.getSheetByName('Customers');
  const active = {};
  const completed = [];
  const customerMap = {};
  const sessionRows = sessions.getRange(FIRST_DATA_ROW, 1, LAST_DATA_ROW - FIRST_DATA_ROW + 1, 18).getValues();
  sessionRows.forEach(function(row) {
    if (!String(row[0] || '').trim()) return;
    const session = sessionFromRow_(row);
    if (String(row[15] || '').toUpperCase() === 'ACTIVE') active[session.uid] = session;
    else if (session.checkout || String(row[15] || '').toUpperCase() === 'PAID') completed.push(session);
  });
  const customerRows = customers.getRange(FIRST_DATA_ROW, 1, LAST_DATA_ROW - FIRST_DATA_ROW + 1, 8).getValues();
  customerRows.forEach(function(row) {
    if (!String(row[0] || '').trim()) return;
    customerMap[String(row[0]).toUpperCase()] = {
      name:String(row[1] || ''), email:String(row[2] || ''), address:String(row[3] || ''), phone:String(row[4] || ''), type:String(row[5] || 'Regular').toLowerCase()
    };
  });
  const settings = ss.getSheetByName('Settings');
  const baseRate = Number(settings.getRange('B4').getValue());
  const adjustment = Number(settings.getRange('B5').getValue());
  return {
    customers: customerMap,
    active: active,
    completed: completed,
    settings: {baseRate:baseRate > 0 ? baseRate : 30, adjustment:adjustment >= 0 ? adjustment : 0},
    fetchedAt: new Date().toISOString()
  };
}

function sessionFromRow_(row) {
  const baseRate = Number(row[11]) || 30;
  const discount = Number(row[12]) || 0;
  const appliedRate = Number(row[13]) || baseRate * (1 - discount);
  return {
    sessionId:String(row[0] || ''), uid:String(row[2] || '').toUpperCase(),
    customer:{name:String(row[3] || ''), email:String(row[4] || ''), address:String(row[5] || ''), phone:String(row[6] || ''), type:String(row[7] || 'Regular').toLowerCase()},
    checkin:iso_(row[8]), checkout:iso_(row[9]), minutes:Number(row[10]) || 0,
    rate:appliedRate, amount:Number(row[14]) || 0, staff:String(row[16] || 'Front desk')
  };
}

function findRow_(sheet, column, value) {
  const values = sheet.getRange(FIRST_DATA_ROW, column, LAST_DATA_ROW - FIRST_DATA_ROW + 1, 1).getValues();
  const target = String(value || '').trim().toUpperCase();
  for (let i = 0; i < values.length; i++) if (String(values[i][0] || '').trim().toUpperCase() === target) return FIRST_DATA_ROW + i;
  return 0;
}

function findActiveRow_(sheet, uid) {
  const rows = sheet.getRange(FIRST_DATA_ROW, 3, LAST_DATA_ROW - FIRST_DATA_ROW + 1, 14).getValues();
  const target = String(uid || '').trim().toUpperCase();
  for (let i = 0; i < rows.length; i++) if (String(rows[i][0] || '').trim().toUpperCase() === target && String(rows[i][13] || '').toUpperCase() === 'ACTIVE') return FIRST_DATA_ROW + i;
  return 0;
}

function firstEmptyRow_(sheet, column) {
  const values = sheet.getRange(FIRST_DATA_ROW, column, LAST_DATA_ROW - FIRST_DATA_ROW + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) if (!String(values[i][0] || '').trim()) return FIRST_DATA_ROW + i;
  throw new Error('The MSpace sheet has reached its 2,000-row operating limit.');
}

function date_(value) { return value instanceof Date ? value : new Date(value); }
function iso_(value) { return value instanceof Date ? value.toISOString() : (value ? new Date(value).toISOString() : ''); }
function titleCase_(value) { const text = String(value || 'regular').toLowerCase(); return text.charAt(0).toUpperCase() + text.slice(1); }
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
