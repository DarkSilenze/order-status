// server.js
const express = require('express');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const db = require('./db');

const app = express();
const PORT = process.env.PORT || 2026;

app.use(express.json());
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.use(express.static(__dirname));

/* =========================================
   Multer — store uploaded files in memory
   ========================================= */
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB
});

const ALLOWED_COLS = ['Order', 'Customer PO', '_1', 'Planned Delivery Date', 'Carrier/LSP', 'Status'];
const VALID_STATUS = ['OPEN', 'CHECKED IN', 'SHIPPED', 'LATE', 'CANCEL'];

/* =========================================
   GET /api/orders
   ========================================= */
app.get('/api/orders', (req, res) => {
  try {
    const rows = db.prepare(`
      SELECT "Order", "Customer PO", "_1", "Planned Delivery Date", "Carrier/LSP", "Status"
      FROM orders ORDER BY rowid_pk DESC
    `).all();
    res.json(rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

/* =========================================
   POST /api/orders — create one
   ========================================= */
app.post('/api/orders', (req, res) => {
  try {
    const body = req.body || {};
    const cols = [];
    const placeholders = [];
    const values = [];

    for (const [key, val] of Object.entries(body)) {
      if (!ALLOWED_COLS.includes(key)) continue;
      cols.push(`"${key}"`);
      placeholders.push('?');
      values.push(val);
    }
    if (cols.length === 0) return res.status(400).json({ error: 'No valid fields provided' });

    db.prepare(`INSERT INTO orders (${cols.join(', ')}) VALUES (${placeholders.join(', ')})`).run(...values);
    const created = db.prepare(`
      SELECT "Order", "Customer PO", "_1", "Planned Delivery Date", "Carrier/LSP", "Status"
      FROM orders ORDER BY rowid_pk DESC LIMIT 1
    `).get();
    res.status(201).json(created);
  } catch (err) {
    if (String(err.message).includes('UNIQUE')) return res.status(409).json({ error: 'Order already exists' });
    res.status(500).json({ error: err.message });
  }
});

/* =========================================
   PATCH /api/orders/:id — update
   ========================================= */
app.patch('/api/orders/:id', (req, res) => {
  try {
    const idValue = req.params.id;
    const body = req.body || {};
    const matchColumn = body.column || 'Order';
    if (!ALLOWED_COLS.includes(matchColumn)) return res.status(400).json({ error: 'Invalid match column' });

    const updates = [];
    const values = [];
    for (const [key, val] of Object.entries(body)) {
      if (key === 'column') continue;
      if (!ALLOWED_COLS.includes(key)) continue;
      updates.push(`"${key}" = ?`);
      values.push(val);
    }
    if (updates.length === 0) return res.status(400).json({ error: 'No valid fields to update' });

    values.push(idValue);
    const info = db.prepare(`UPDATE orders SET ${updates.join(', ')} WHERE "${matchColumn}" = ?`).run(...values);
    if (info.changes === 0) return res.status(404).json({ error: 'Order not found' });
    res.json({ success: true, changes: info.changes });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/* =========================================
   DELETE /api/orders/:id
   ========================================= */
app.delete('/api/orders/:id', (req, res) => {
  try {
    const info = db.prepare('DELETE FROM orders WHERE "Order" = ?').run(req.params.id);
    if (info.changes === 0) return res.status(404).json({ error: 'Order not found' });
    res.json({ success: true, deletedCount: info.changes });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/* =========================================
   DELETE /api/orders — all
   ========================================= */
app.delete('/api/orders', (req, res) => {
  try {
    const info = db.prepare('DELETE FROM orders').run();
    db.prepare("DELETE FROM sqlite_sequence WHERE name = 'orders'").run();
    res.json({ success: true, deletedCount: info.changes });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

/* =========================================
   POST /api/orders/upload — Excel / CSV bulk import
   ========================================= */
app.post('/api/orders/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const originalName = req.file.originalname || 'upload';
    const ext = path.extname(originalName).toLowerCase();
    if (!['.xlsx', '.xls', '.csv'].includes(ext)) {
      return res.status(400).json({ error: 'Only .xlsx, .xls, or .csv files are supported' });
    }

    // Parse workbook
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) return res.status(400).json({ error: 'File contains no sheets' });

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
    if (!rows.length) return res.status(400).json({ error: 'File is empty' });

    // Build column-name lookup that tolerates small naming variations
    const headerAliases = {
      'Order': ['Order', 'order', 'ORDER', 'Order Number', 'order_number', 'OrderNumber'],
      'Customer PO': ['Customer PO', 'customer po', 'CustomerPO', 'customer_po', 'PO', 'po_number'],
      '_1': ['_1', 'Customer', 'customer', 'Customer Name', 'customer_name'],
      'Planned Delivery Date': ['Planned Delivery Date', 'planned delivery date', 'PlannedDeliveryDate',
                                 'Delivery Date', 'delivery_date', 'Planned Delivery Time'],
      'Carrier/LSP': ['Carrier/LSP', 'Carrier', 'carrier', 'carrier_lsp', 'LSP', 'lsp'],
      'Status': ['Status', 'status', 'STATUS']
    };

    const headerMap = {};
    for (const [canonical, aliases] of Object.entries(headerAliases)) {
      for (const key of Object.keys(rows[0])) {
        if (aliases.includes(key)) {
          headerMap[canonical] = key;
          break;
        }
      }
    }

    if (!headerMap['Order']) {
      return res.status(400).json({ error: 'Could not find an "Order" column in the file' });
    }

    // Normalize rows to the DB schema
    const normalized = rows.map((r) => {
      const out = {};
      for (const [canonical, srcKey] of Object.entries(headerMap)) {
        let val = r[srcKey];
        if (val === undefined || val === null) val = '';
        val = String(val).trim();
        out[canonical] = val;
      }
      // Default status to OPEN if empty or invalid
      const statusUpper = String(out['Status'] || '').toUpperCase().trim();
      out['Status'] = VALID_STATUS.includes(statusUpper) ? statusUpper : 'OPEN';
      return out;
    }).filter((r) => r['Order']); // Skip rows without an Order

    if (normalized.length === 0) {
      return res.status(400).json({ error: 'No valid rows found (missing "Order" values)' });
    }

    // Insert with conflict handling: update if "Order" already exists
    const insertStmt = db.prepare(`
      INSERT INTO orders ("Order", "Customer PO", "_1", "Planned Delivery Date", "Carrier/LSP", "Status")
      VALUES (@Order, @CustomerPO, @customer, @delivery, @carrier, @status)
      ON CONFLICT("Order") DO UPDATE SET
        "Customer PO" = excluded."Customer PO",
        "_1" = excluded."_1",
        "Planned Delivery Date" = excluded."Planned Delivery Date",
        "Carrier/LSP" = excluded."Carrier/LSP",
        "Status" = excluded."Status"
    `);

    const runAll = db.transaction((list) => {
      for (const r of list) {
        insertStmt.run({
          Order: r['Order'],
          CustomerPO: r['Customer PO'] || '',
          customer: r['_1'] || '',
          delivery: r['Planned Delivery Date'] || '',
          carrier: r['Carrier/LSP'] || '',
          status: r['Status'] || 'OPEN'
        });
      }
    });

    runAll(normalized);

    res.json({
      success: true,
      imported: normalized.length,
      columns: Object.keys(headerMap),
      preview: normalized.slice(0, 5),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to process file: ' + err.message });
  }
});

/* =========================================
   Fallback
   ========================================= */
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

app.listen(PORT, () => {
  console.log(`\n🚀 Order Status API running at http://localhost:${PORT}`);
});