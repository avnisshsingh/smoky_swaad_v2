import express from 'express';
import fs from 'fs';
import path from 'path';
import vm from 'vm';
import {
  SpreadsheetApp,
  PropertiesService,
  DriveApp,
  Utilities,
  Session,
  MimeType,
  Logger,
  CacheService,
  LockService,
  MailApp,
  GmailApp,
  UrlFetchApp,
  ScriptApp,
  ContentService,
  getInvoiceFolder
} from './gas-emulator.js';

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Helper to render HTML with template includes
function renderHtmlFile(filename) {
  const filePath = path.join(process.cwd(), filename.endsWith('.html') ? filename : filename + '.html');
  if (!fs.existsSync(filePath)) {
    return `<!-- File not found: ${filename} -->`;
  }
  let content = fs.readFileSync(filePath, 'utf-8');

  let previous;
  let maxDepth = 20;
  do {
    previous = content;
    content = content.replace(/<\?!=?\s*include\(['"]([^'"]+)['"]\);?\s*\?>/gi, (_, childName) => {
      return renderHtmlFile(childName);
    });
    maxDepth--;
  } while (content !== previous && maxDepth > 0);

  return content;
}

// Client-side google.script.run Polyfill
const GOOGLE_SCRIPT_RUN_POLYFILL = `
<script>
class GoogleScriptRunner {
  constructor(successHandler, failureHandler, userObj) {
    this._success = successHandler;
    this._failure = failureHandler;
    this._userObj = userObj;
    return new Proxy(this, {
      get: (target, prop) => {
        if (prop === 'withSuccessHandler') {
          return (fn) => new GoogleScriptRunner(fn, target._failure, target._userObj);
        }
        if (prop === 'withFailureHandler') {
          return (fn) => new GoogleScriptRunner(target._success, fn, target._userObj);
        }
        if (prop === 'withUserObject') {
          return (u) => new GoogleScriptRunner(target._success, target._failure, u);
        }
        if (prop in target) {
          return target[prop];
        }
        return (...args) => {
          fetch('/api/rpc', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: prop, args: args })
          })
          .then(res => res.json())
          .then(data => {
            if (data.error) {
              if (target._failure) target._failure(new Error(data.error), target._userObj);
              else console.error('RPC Error [' + prop + ']:', data.error);
            } else {
              if (target._success) target._success(data.result, target._userObj);
            }
          })
          .catch(err => {
            if (target._failure) target._failure(err, target._userObj);
            else console.error('RPC Network Error [' + prop + ']:', err);
          });
        };
      }
    });
  }
}
window.google = window.google || {};
window.google.script = window.google.script || {
  run: new GoogleScriptRunner()
};
</script>
`;

// Global Google Apps Script Context
const gasContext = {
  console,
  Date,
  Math,
  Number,
  String,
  Array,
  Object,
  JSON,
  RegExp,
  parseInt,
  parseFloat,
  isNaN,
  isFinite,
  encodeURIComponent,
  decodeURIComponent,
  setTimeout,
  clearTimeout,
  setInterval,
  clearInterval,

  // Apps Script Globals
  SpreadsheetApp,
  PropertiesService,
  DriveApp,
  Utilities,
  Session,
  MimeType,
  Logger,
  CacheService,
  LockService,
  MailApp,
  GmailApp,
  UrlFetchApp,
  ScriptApp,
  ContentService,
  getInvoiceFolder,

  include: (filename) => {
    return renderHtmlFile(filename);
  },

  HtmlService: {
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
    createHtmlOutput: (content) => ({
      getContent: () => content,
      addMetaTag: function () { return this; },
      setTitle: function () { return this; },
      setXFrameOptionsMode: function () { return this; }
    }),
    createHtmlOutputFromFile: (filename) => {
      const content = renderHtmlFile(filename);
      return {
        getContent: () => content,
        evaluate: function () { return this; }
      };
    },
    createTemplateFromFile: (filename) => {
      return {
        evaluate: function () {
          const content = renderHtmlFile(filename);
          return {
            getContent: () => content,
            addMetaTag: function () { return this; },
            setTitle: function () { return this; },
            setXFrameOptionsMode: function () { return this; }
          };
        }
      };
    }
  }
};

vm.createContext(gasContext);

// Load all Apps Script JS files in logical dependency order
const scriptFiles = [
  'Constants.js',
  'PurchaseConstants.js',
  'Utilities.js',
  'Validation.js',
  'PurchaseValidation.js',
  'Storage.js',
  'UnitConverter.js',
  'Drive.js',
  'BackupService.js',
  'MenuService.js',
  'CustomerService.js',
  'SettingsService.js',
  'BusinessCostsService.js',
  'PersonalExpensesService.js',
  'PricingCalculatorService.js',
  'ProfitService.js',
  'PurchaseService.js',
  'ReportsService.js',
  'WebOrderService.js',
  'SaveOrder.js',
  'PurchaseSave.js',
  'Invoices.js',
  'Code.js'
];

for (const file of scriptFiles) {
  const filePath = path.join(process.cwd(), file);
  if (fs.existsSync(filePath)) {
    const code = fs.readFileSync(filePath, 'utf-8');
    try {
      vm.runInContext(code, gasContext, { filename: file });
    } catch (err) {
      console.error(`Error loading Apps Script file [${file}]:`, err);
    }
  }
}

// RPC Route
app.post('/api/rpc', (req, res) => {
  const { action, args = [] } = req.body;
  if (!action || typeof gasContext[action] !== 'function') {
    console.error(`RPC method '${action}' not found in backend context.`);
    return res.status(400).json({ error: `Function '${action}' not found.` });
  }
  try {
    const result = gasContext[action].apply(gasContext, args);
    res.json({ success: true, result });
  } catch (err) {
    console.error(`RPC execution error in '${action}':`, err);
    res.status(500).json({ error: err.message || String(err) });
  }
});

// Main Web App Handler
app.get('*', (req, res) => {
  try {
    let indexHtml = renderHtmlFile('Index');
    // Inject google.script.run polyfill before </head>
    if (indexHtml.includes('</head>')) {
      indexHtml = indexHtml.replace('</head>', `${GOOGLE_SCRIPT_RUN_POLYFILL}\n</head>`);
    } else {
      indexHtml = GOOGLE_SCRIPT_RUN_POLYFILL + indexHtml;
    }
    res.setHeader('Content-Type', 'text/html');
    res.send(indexHtml);
  } catch (err) {
    console.error('Error serving index page:', err);
    res.status(500).send('Server Error: ' + err.message);
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Smoky Swaad ERP running on http://0.0.0.0:${PORT}`);
});
