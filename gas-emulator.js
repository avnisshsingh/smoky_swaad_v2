import fs from 'fs';
import path from 'path';

function colLetterToNum(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

function parseA1(a1, maxRows = 100, maxCols = 50) {
  if (typeof a1 !== 'string') return null;
  const rangeMatch = a1.match(/^([A-Z]+)(\d+)?:([A-Z]+)(\d+)?$/i);
  if (rangeMatch) {
    const col1 = colLetterToNum(rangeMatch[1].toUpperCase());
    const row1 = rangeMatch[2] ? parseInt(rangeMatch[2], 10) : 1;
    const col2 = colLetterToNum(rangeMatch[3].toUpperCase());
    const row2 = rangeMatch[4] ? parseInt(rangeMatch[4], 10) : maxRows;
    return {
      row: row1,
      col: col1,
      numRows: Math.max(1, row2 - row1 + 1),
      numCols: Math.max(1, col2 - col1 + 1)
    };
  }

  const singleMatch = a1.match(/^([A-Z]+)(\d+)$/i);
  if (singleMatch) {
    const col = colLetterToNum(singleMatch[1].toUpperCase());
    const row = parseInt(singleMatch[2], 10);
    return { row, col, numRows: 1, numCols: 1 };
  }

  return { row: 1, col: 1, numRows: maxRows, numCols: maxCols };
}

class TextFinder {
  constructor(range, queryText) {
    this.range = range;
    this.queryText = String(queryText || '');
    this.isMatchEntireCell = false;
    this.isMatchCase = false;
  }

  matchEntireCell(val) {
    this.isMatchEntireCell = !!val;
    return this;
  }

  matchCase(val) {
    this.isMatchCase = !!val;
    return this;
  }

  _matches(cellVal) {
    let text = String(cellVal ?? '');
    let search = this.queryText;
    if (!this.isMatchCase) {
      text = text.toLowerCase();
      search = search.toLowerCase();
    }
    if (this.isMatchEntireCell) {
      return text === search;
    }
    return text.includes(search);
  }

  findNext() {
    const vals = this.range.getValues();
    const startRow = this.range.startRow;
    const startCol = this.range.startCol;
    for (let r = 0; r < vals.length; r++) {
      for (let c = 0; c < vals[r].length; c++) {
        if (this._matches(vals[r][c])) {
          return new Range(this.range.sheet, startRow + r, startCol + c, 1, 1);
        }
      }
    }
    return null;
  }

  findAll() {
    const results = [];
    const vals = this.range.getValues();
    const startRow = this.range.startRow;
    const startCol = this.range.startCol;
    for (let r = 0; r < vals.length; r++) {
      for (let c = 0; c < vals[r].length; c++) {
        if (this._matches(vals[r][c])) {
          results.push(new Range(this.range.sheet, startRow + r, startCol + c, 1, 1));
        }
      }
    }
    return results;
  }
}

class Range {
  constructor(sheet, startRow, startCol, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.startRow = Math.max(1, startRow);
    this.startCol = Math.max(1, startCol);
    this.numRows = Math.max(1, numRows);
    this.numCols = Math.max(1, numCols);
  }

  getRow() {
    return this.startRow;
  }

  getColumn() {
    return this.startCol;
  }

  getValue() {
    const vals = this.getValues();
    return vals[0] ? vals[0][0] : '';
  }

  getDisplayValue() {
    const val = this.getValue();
    return val === null || val === undefined ? '' : String(val);
  }

  getValues() {
    const result = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.startRow - 1 + r;
      const rowArr = [];
      const sheetRow = this.sheet.data[rowIdx] || [];
      for (let c = 0; c < this.numCols; c++) {
        const colIdx = this.startCol - 1 + c;
        rowArr.push(sheetRow[colIdx] !== undefined ? sheetRow[colIdx] : '');
      }
      result.push(rowArr);
    }
    return result;
  }

  getDisplayValues() {
    return this.getValues().map(row => row.map(val => (val === null || val === undefined ? '' : String(val))));
  }

  setValue(val) {
    this.setValues([[val]]);
  }

  setValues(matrix) {
    for (let r = 0; r < matrix.length; r++) {
      const rowIdx = this.startRow - 1 + r;
      while (this.sheet.data.length <= rowIdx) {
        this.sheet.data.push([]);
      }
      for (let c = 0; c < matrix[r].length; c++) {
        const colIdx = this.startCol - 1 + c;
        this.sheet.data[rowIdx][colIdx] = matrix[r][c];
      }
    }
  }

  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.startRow - 1 + r;
      if (this.sheet.data[rowIdx]) {
        for (let c = 0; c < this.numCols; c++) {
          const colIdx = this.startCol - 1 + c;
          this.sheet.data[rowIdx][colIdx] = '';
        }
      }
    }
  }

  clear() {
    this.clearContent();
  }

  createTextFinder(text) {
    return new TextFinder(this, text);
  }
}

class InMemorySheet {
  constructor(name, initialData = []) {
    this.name = name;
    this.data = initialData.map(row => [...row]);
  }

  getLastRow() {
    for (let i = this.data.length - 1; i >= 0; i--) {
      const row = this.data[i];
      if (row && row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
        return i + 1;
      }
    }
    return 0;
  }

  getLastColumn() {
    let maxCol = 0;
    for (const row of this.data) {
      if (row) {
        for (let c = row.length - 1; c >= 0; c--) {
          if (row[c] !== '' && row[c] !== null && row[c] !== undefined) {
            maxCol = Math.max(maxCol, c + 1);
            break;
          }
        }
      }
    }
    return maxCol;
  }

  getRange(arg1, arg2, arg3, arg4) {
    if (typeof arg1 === 'string') {
      const coords = parseA1(arg1, Math.max(100, this.data.length + 10), Math.max(50, this.getLastColumn() + 5));
      return new Range(this, coords.row, coords.col, coords.numRows, coords.numCols);
    }
    return new Range(this, arg1, arg2, arg3 || 1, arg4 || 1);
  }

  getName() {
    return this.name;
  }

  setName(name) {
    this.name = name;
  }

  getDataRange() {
    const lastRow = Math.max(1, this.getLastRow());
    const lastCol = Math.max(1, this.getLastColumn());
    return new Range(this, 1, 1, lastRow, lastCol);
  }

  getMaxRows() {
    return Math.max(100, this.data.length);
  }

  getMaxColumns() {
    return Math.max(50, this.getLastColumn());
  }

  getCharts() {
    return [];
  }

  appendRow(rowArray) {
    this.data.push([...rowArray]);
  }

  deleteRow(rowNumber) {
    const idx = rowNumber - 1;
    if (idx >= 0 && idx < this.data.length) {
      this.data.splice(idx, 1);
    }
  }

  clear() {
    this.data = [];
  }

  clearContent() {
    this.data = [];
  }

  toast(msg, title, timeout) {
    console.log(`[Toast] ${title ? title + ': ' : ''}${msg}`);
  }
}

class Spreadsheet {
  constructor() {
    this.sheets = new Map();
    this.initDefaultSheets();
  }

  initDefaultSheets() {
    // SHEETS.MENU
    this.sheets.set('Menu', new InMemorySheet('Menu', [
      ['Item ID', 'Item Name', 'Category', 'Cost', 'Price', 'Tax %', 'Margin', 'Prep Time', 'Active'],
      ['MNU00001', 'Smoky Butter Chicken', 'Main Course', 120, 280, 5, '57%', '20 mins', 'Yes'],
      ['MNU00002', 'Smokey Paneer Tikka', 'Starters', 80, 220, 5, '63%', '15 mins', 'Yes'],
      ['MNU00003', 'Dal Makhani Special', 'Main Course', 60, 180, 5, '66%', '15 mins', 'Yes'],
      ['MNU00004', 'Chicken Dum Biryani', 'Biryani & Rice', 110, 260, 5, '57%', '25 mins', 'Yes'],
      ['MNU00005', 'Butter Naan (2 pcs)', 'Breads', 20, 60, 5, '66%', '10 mins', 'Yes'],
      ['MNU00006', 'Smokey Tandoori Roti', 'Breads', 10, 30, 5, '66%', '10 mins', 'Yes'],
      ['MNU00007', 'Gulab Jamun (2 pcs)', 'Desserts', 25, 70, 5, '64%', '5 mins', 'Yes'],
      ['MNU00008', 'Mango Lassi', 'Beverages', 30, 90, 5, '66%', '5 mins', 'Yes'],
      ['MNU00009', 'Kadhai Paneer', 'Main Course', 90, 240, 5, '62%', '20 mins', 'Yes'],
      ['MNU00010', 'Chicken Tikka Roll', 'Starters', 50, 150, 5, '66%', '12 mins', 'Yes']
    ]));

    // SHEETS.CUSTOMERS
    this.sheets.set('Customers', new InMemorySheet('Customers', [
      ['Customer ID', 'Customer Name', 'Mobile', 'Delivery Area', 'House Address', 'First Order Date', 'Last Order Date', 'Total Orders', 'Lifetime Spend', 'Last Order ID'],
      ['CUS00001', 'Rahul Sharma', '9876543210', 'Central City', 'Flat 402, Sunshine Apartments', '2026-08-01', '2026-08-09', 3, 1420, 'SS01004'],
      ['CUS00002', 'Priya Patel', '9812345678', 'North Suburbs', 'House 12, Green Avenue', '2026-08-03', '2026-08-08', 2, 850, 'SS01003'],
      ['CUS00003', 'Amit Kumar', '9988776655', 'East Zone', 'Plot 88, Metro Layout', '2026-08-05', '2026-08-05', 1, 520, 'SS01002']
    ]));

    // SHEETS.SETTINGS
    // Needs D: PaymentModes, G: PaymentStatus, M: OrderTypes, P: DeliverySlots, V:W DeliveryAreas
    const settingsData = Array.from({ length: 50 }, () => Array(30).fill(''));
    settingsData[0][3] = 'Payment Mode'; // D1
    settingsData[1][3] = 'UPI';
    settingsData[2][3] = 'Cash';
    settingsData[3][3] = 'Card';
    settingsData[4][3] = 'NetBanking';

    settingsData[0][6] = 'Payment Status'; // G1
    settingsData[1][6] = 'Paid';
    settingsData[2][6] = 'Pending';
    settingsData[3][6] = 'Partial';

    settingsData[0][12] = 'Order Type'; // M1
    settingsData[1][12] = 'Delivery';
    settingsData[2][12] = 'Takeaway';
    settingsData[3][12] = 'Dine-in';

    settingsData[0][15] = 'Delivery Slot'; // P1
    settingsData[1][15] = 'Immediate';
    settingsData[2][15] = '12:00 PM - 1:00 PM';
    settingsData[3][15] = '1:00 PM - 2:00 PM';
    settingsData[4][15] = '7:00 PM - 8:00 PM';

    settingsData[0][21] = 'Area'; // V1
    settingsData[0][22] = 'Charge'; // W1
    const areas = [
      ['Central City', 40],
      ['North Suburbs', 60],
      ['East Zone', 50],
      ['West End', 50],
      ['South Hills', 70]
    ];
    areas.forEach((a, idx) => {
      settingsData[idx + 1][21] = a[0];
      settingsData[idx + 1][22] = a[1];
    });

    this.sheets.set('Settings', new InMemorySheet('Settings', settingsData));

    // SHEETS.ORDERS
    this.sheets.set('Orders', new InMemorySheet('Orders', [
      ['Order ID', 'Order Date', 'Customer Name', 'Mobile', 'Delivery Area', 'House Address', 'Delivery Slot', 'Order Type', 'Payment Mode', 'Payment Status', 'Special Instructions', 'Subtotal', 'Discount', 'Delivery Charge', 'Grand Total', 'Created At'],
      ['SS01001', '01/08/2026', 'Rahul Sharma', '9876543210', 'Central City', 'Flat 402, Sunshine Apartments', 'Immediate', 'Delivery', 'UPI', 'Paid', 'Less Spicy', 500, 50, 40, 490, '2026-08-01 13:30:00'],
      ['SS01002', '05/08/2026', 'Amit Kumar', '9988776655', 'East Zone', 'Plot 88, Metro Layout', 'Immediate', 'Delivery', 'Cash', 'Paid', 'Extra Gravy', 520, 50, 50, 520, '2026-08-05 20:15:00'],
      ['SS01003', '08/08/2026', 'Priya Patel', '9812345678', 'North Suburbs', 'House 12, Green Avenue', '7:00 PM - 8:00 PM', 'Delivery', 'UPI', 'Paid', '', 800, 100, 60, 760, '2026-08-08 19:10:00'],
      ['SS01004', '09/08/2026', 'Rahul Sharma', '9876543210', 'Central City', 'Flat 402, Sunshine Apartments', 'Immediate', 'Delivery', 'UPI', 'Paid', 'Call Customer', 410, 40, 40, 410, '2026-08-09 14:00:00']
    ]));

    // SHEETS.ORDER_ITEMS
    this.sheets.set('OrderItems', new InMemorySheet('OrderItems', [
      ['Order ID', 'Item Name', 'Quantity', 'Price', 'Total'],
      ['SS01001', 'Smoky Butter Chicken', 1, 280, 280],
      ['SS01001', 'Butter Naan (2 pcs)', 2, 60, 120],
      ['SS01001', 'Mango Lassi', 1, 90, 90],
      ['SS01002', 'Chicken Dum Biryani', 2, 260, 520],
      ['SS01003', 'Smokey Paneer Tikka', 2, 220, 440],
      ['SS01003', 'Dal Makhani Special', 2, 180, 360],
      ['SS01004', 'Kadhai Paneer', 1, 240, 240],
      ['SS01004', 'Smokey Tandoori Roti', 3, 30, 90],
      ['SS01004', 'Gulab Jamun (2 pcs)', 1, 70, 70]
    ]));

    // SHEETS.CONFIG
    const configData = [['Setting', 'Value'], ['LastOrderID', 1004]];
    this.sheets.set('Config', new InMemorySheet('Config', configData));

    // SHEETS.POS
    this.sheets.set('POS', new InMemorySheet('POS', Array.from({ length: 35 }, () => ['', ''])));

    // SHEETS.BUSINESS_COSTS
    this.sheets.set('BusinessCosts', new InMemorySheet('BusinessCosts', [
      ['ID', 'Date', 'Category', 'Description', 'Amount', 'Payment Mode', 'Notes'],
      ['BC0001', '2026-08-01', 'Rent', 'Kitchen Premises Rent August', 25000, 'Bank Transfer', 'Monthly'],
      ['BC0002', '2026-08-02', 'Electricity', 'Commercial Power Bill July', 4500, 'UPI', 'Paid']
    ]));

    // SHEETS.PERSONAL_EXPENSES
    this.sheets.set('PersonalExpenses', new InMemorySheet('PersonalExpenses', [
      ['ID', 'Date', 'Category', 'Description', 'Amount', 'Payment Mode', 'For', 'Nature', 'Recurring'],
      ['PE0001', '2026-08-02', 'Groceries', 'Home Supplies', 1200, 'UPI', 'Family', 'Expense', 'No']
    ]));

    // SHEETS.PURCHASE_REGISTER
    this.sheets.set('Purchase Register', new InMemorySheet('Purchase Register', [
      ['Purchase ID', 'Date', 'Supplier', 'Item', 'Category', 'Quantity', 'Unit', 'Unit Price', 'Total Cost', 'Payment Status'],
      ['PUR0001', '2026-08-01', 'Metro Cash & Carry', 'Chicken (Fresh)', 'Raw Material', 20, 'kg', 180, 3600, 'Paid'],
      ['PUR0002', '2026-08-01', 'Verka Dairy', 'Paneer (Fresh)', 'Raw Material', 10, 'kg', 220, 2200, 'Paid']
    ]));

    // SHEETS.RECIPE
    this.sheets.set('Recipe Master', new InMemorySheet('Recipe Master', [
      ['Item ID', 'Item Name', 'Ingredient', 'Quantity', 'Unit'],
      ['MNU00001', 'Smoky Butter Chicken', 'Chicken', 200, 'g']
    ]));

    // SHEETS.DASHBOARD
    this.sheets.set('Dashboard', new InMemorySheet('Dashboard', [
      ['Metric', 'Value'],
      ['Total Sales', 2680]
    ]));
  }

  getSheetByName(name) {
    if (!this.sheets.has(name)) {
      this.sheets.set(name, new InMemorySheet(name, []));
    }
    return this.sheets.get(name);
  }

  getSheets() {
    return Array.from(this.sheets.values());
  }

  toast(msg, title, timeout) {
    console.log(`[Spreadsheet Toast] ${title ? title + ': ' : ''}${msg}`);
  }
}

const activeSS = new Spreadsheet();

export const SpreadsheetApp = {
  getActiveSpreadsheet: () => activeSS,
  flush: () => {}
};

class PropertyStore {
  constructor() {
    this.props = new Map();
  }
  getProperty(key) {
    return this.props.get(key) ?? null;
  }
  setProperty(key, value) {
    this.props.set(key, String(value));
  }
  getProperties() {
    const obj = {};
    for (const [k, v] of this.props.entries()) obj[k] = v;
    return obj;
  }
  setProperties(obj) {
    if (obj && typeof obj === 'object') {
      for (const k of Object.keys(obj)) this.props.set(k, String(obj[k]));
    }
  }
}

const scriptProps = new PropertyStore();

export const PropertiesService = {
  getScriptProperties: () => scriptProps,
  getUserProperties: () => scriptProps,
  getDocumentProperties: () => scriptProps
};

class DriveFolder {
  constructor(name = 'Root') {
    this.name = name;
  }
  createFile(blob) {
    return {
      getUrl: () => `https://drive.google.com/mock_file_${Date.now()}`,
      getName: () => (blob.getName ? blob.getName() : 'file.pdf')
    };
  }
}

export const DriveApp = {
  getFolderById: () => new DriveFolder(),
  getFoldersByName: () => ({
    hasNext: () => true,
    next: () => new DriveFolder()
  }),
  getFileById: (id) => ({
    getBlob: () => ({
      getBytes: () => Buffer.from('mock_file_content'),
      getContentType: () => 'image/png',
      getAsString: () => 'mock_file_content',
      getName: () => 'logo.png'
    }),
    getName: () => 'logo.png',
    getUrl: () => 'https://drive.google.com/mock_file',
    getDownloadUrl: () => 'https://drive.google.com/mock_download',
    getThumbnail: () => null,
    isTrashed: () => false,
    setTrashed: () => {}
  }),
  getFilesByName: (name) => ({
    hasNext: () => true,
    next: () => DriveApp.getFileById('mock_id')
  }),
  createFile: (blob) => ({
    getUrl: () => `https://drive.google.com/mock_file_${Date.now()}`,
    getName: () => (blob && blob.getName ? blob.getName() : 'file.pdf')
  }),
  createFolder: (name) => new DriveFolder(name)
};

export const Utilities = {
  base64Decode: (str) => Buffer.from(str, 'base64'),
  base64Encode: (data) => Buffer.from(data).toString('base64'),
  newBlob: (bytes, mime, name) => ({
    getBytes: () => bytes,
    getAsString: () => (Buffer.isBuffer(bytes) ? bytes.toString('utf-8') : String(bytes)),
    getName: () => name || 'blob',
    getContentType: () => mime || 'application/octet-stream'
  }),
  formatDate: (date, tz, fmt) => {
    const d = new Date(date);
    if (isNaN(d.getTime())) return '';
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    if (fmt === 'dd/MM/yyyy') return `${day}/${month}/${year}`;
    if (fmt === 'yyyy-MM-dd') return `${year}-${month}-${day}`;
    return d.toISOString();
  },
  sleep: (ms) => {}
};

export const Session = {
  getScriptTimeZone: () => 'Asia/Kolkata',
  getActiveUser: () => ({ getEmail: () => 'admin@smokyswaad.com' })
};

export const MimeType = {
  PDF: 'application/pdf',
  HTML: 'text/html',
  CSV: 'text/csv'
};

export const Logger = {
  log: (...args) => console.log('[GAS Logger]', ...args)
};

class Cache {
  constructor() {
    this.store = new Map();
  }
  get(key) {
    const item = this.store.get(key);
    if (!item) return null;
    if (item.expiresAt && Date.now() > item.expiresAt) {
      this.store.delete(key);
      return null;
    }
    return item.value;
  }
  getAll(keys) {
    const result = {};
    if (Array.isArray(keys)) {
      keys.forEach(k => {
        const val = this.get(k);
        if (val !== null) result[k] = val;
      });
    }
    return result;
  }
  put(key, value, expirationInSeconds = 600) {
    this.store.set(key, {
      value: String(value),
      expiresAt: Date.now() + (expirationInSeconds * 1000)
    });
  }
  putAll(values, expirationInSeconds = 600) {
    if (values && typeof values === 'object') {
      for (const [k, v] of Object.entries(values)) {
        this.put(k, v, expirationInSeconds);
      }
    }
  }
  remove(key) {
    this.store.delete(key);
  }
  removeAll(keys) {
    if (Array.isArray(keys)) {
      keys.forEach(k => this.store.delete(k));
    }
  }
}

const defaultCache = new Cache();

export const CacheService = {
  getScriptCache: () => defaultCache,
  getUserCache: () => defaultCache,
  getDocumentCache: () => defaultCache
};

class Lock {
  tryLock(timeoutInMillis) {
    return true;
  }
  waitLock(timeoutInMillis) {}
  releaseLock() {}
  hasLock() {
    return true;
  }
}

const defaultLock = new Lock();

export const LockService = {
  getScriptLock: () => defaultLock,
  getUserLock: () => defaultLock,
  getDocumentLock: () => defaultLock
};

export const MailApp = {
  sendEmail: (to, subject, body, options) => {
    console.log(`[MailApp.sendEmail] To: ${to}, Subject: ${subject}`);
  },
  getRemainingDailyQuota: () => 100
};

export const GmailApp = {
  sendEmail: (to, subject, body, options) => {
    console.log(`[GmailApp.sendEmail] To: ${to}, Subject: ${subject}`);
  }
};

export const UrlFetchApp = {
  fetch: (url, params = {}) => {
    console.log(`[UrlFetchApp.fetch] URL: ${url}`);
    return {
      getContentText: () => '{}',
      getResponseCode: () => 200,
      getHeaders: () => ({})
    };
  }
};

export const ScriptApp = {
  getOAuthToken: () => 'mock_oauth_token',
  getProjectTriggers: () => [],
  newTrigger: () => ({
    timeBased: () => ({
      everyDays: () => ({ create: () => {} }),
      everyHours: () => ({ create: () => {} })
    })
  }),
  deleteTrigger: () => {}
};

export const ContentService = {
  MimeType: {
    JSON: 'application/json',
    TEXT: 'text/plain',
    JAVASCRIPT: 'text/javascript'
  },
  createTextOutput: (content) => ({
    setContent: function (c) { return this; },
    setMimeType: function (m) { return this; },
    getContent: () => content
  })
};

export function getInvoiceFolder() {
  return new DriveFolder('Invoices');
}
