/* Minimal Google Apps Script / Sheets mock, faithful on the one behaviour that matters
   here: writing a value into a cell DESTROYS any formula in it. Used to prove the app
   never writes over the workbook's formula columns. */
const fs = require('fs');

function makeSheet(name, header, formulaCols, initialRows, maxRows) {
  // cells[r][c] = {v: value, f: formula or ''}   (r,c are 0-based; row 0 = header)
  const cells = [];
  const cols = header.length;
  const rows = Math.max(maxRows || 500, (initialRows ? initialRows.length : 0) + 1);
  for (let r = 0; r < rows; r++) {
    cells.push(Array.from({ length: cols }, () => ({ v: '', f: '' })));
  }
  header.forEach((h, c) => { cells[0][c] = { v: h, f: '' }; });
  (initialRows || []).forEach((row, i) => {
    row.forEach((val, c) => { cells[i + 1][c] = { v: val, f: '' }; });
    // give this data row the sheet's formulas
    (formulaCols || []).forEach(c => {
      cells[i + 1][c] = { v: '(computed)', f: `=FORMULA_COL_${c}(R${i + 2})` };
    });
  });

  const sh = {
    _cells: cells,
    _name: name,
    getName: () => name,
    getMaxRows: () => cells.length,
    getLastRow() {
      for (let r = cells.length - 1; r >= 0; r--) {
        if (cells[r].some(c => c.v !== '' || c.f !== '')) return r + 1;
      }
      return 0;
    },
    // like the real API: grows when a new header is written past the current width
    getLastColumn() {
      let last = 0;
      for (let r = 0; r < cells.length; r++) {
        for (let c = cells[r].length - 1; c >= last; c--) {
          const cell = cells[r][c];
          if (cell && (cell.v !== '' || cell.f !== '')) { last = Math.max(last, c + 1); break; }
        }
      }
      return last;
    },
    getDataRange() { return sh.getRange(1, 1, sh.getLastRow() || 1, cols); },
    getRange(row, col, numRows = 1, numCols = 1) {
      if (col < 1) throw new Error('The starting column of the range is too small.');
      if (row < 1) throw new Error('The starting row of the range is too small.');
      return {
        _row: row, _col: col, _n: numRows, _m: numCols,
        getValues() {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const line = [];
            for (let c = 0; c < numCols; c++) {
              const cell = cells[row - 1 + r] && cells[row - 1 + r][col - 1 + c];
              line.push(cell ? cell.v : '');
            }
            out.push(line);
          }
          return out;
        },
        getValue() { return this.getValues()[0][0]; },
        getFormulas() {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const line = [];
            for (let c = 0; c < numCols; c++) {
              const cell = cells[row - 1 + r] && cells[row - 1 + r][col - 1 + c];
              line.push(cell ? cell.f : '');
            }
            out.push(line);
          }
          return out;
        },
        setValues(vals) {
          for (let r = 0; r < vals.length; r++) {
            for (let c = 0; c < vals[r].length; c++) {
              // THE key behaviour: setting a value wipes the formula
              cells[row - 1 + r][col - 1 + c] = { v: vals[r][c], f: '' };
            }
          }
          return this;
        },
        setValue(v) { return this.setValues([[v]]); },
        copyTo(dest, type, transpose) {
          const srcCell = cells[row - 1][col - 1];
          for (let r = 0; r < dest._n; r++) {
            for (let c = 0; c < dest._m; c++) {
              const tr = dest._row - 1 + r, tc = dest._col - 1 + c;
              if (!cells[tr]) continue;
              const rebased = srcCell.f.replace(/R\d+/, 'R' + (tr + 1));
              cells[tr][tc] = { v: cells[tr][tc].v, f: rebased };
            }
          }
          return this;
        }
      };
    }
  };
  return sh;
}

function buildEnv(sheets) {
  const env = {};
  env.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getName: () => 'QSNCC_Disinfection_Machine_Registry',
      getUrl: () => 'https://docs.google.com/spreadsheets/d/FAKE/edit',
      getId: () => 'FAKE',
      getSheetByName: (n) => sheets[n] || null
    }),
    flush: () => {},
    CopyPasteType: { PASTE_FORMULA: 'PASTE_FORMULA' }
  };
  env.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  env.Session = { getEffectiveUser: () => ({ getEmail: () => 'tester@example.com' }) };
  env.ContentService = { createTextOutput: (s) => ({ setMimeType: () => s }), MimeType: { JSON: 'json' } };
  env.HtmlService = { createHtmlOutputFromFile: () => ({ setTitle(){return this}, addMetaTag(){return this}, setXFrameOptionsMode(){return this} }), XFrameOptionsMode: { ALLOWALL: 1 } };
  return env;
}

function loadCode(env) {
  const src = fs.readFileSync('/home/user/QNCC-reccord/Code.gs', 'utf8');
  const names = Object.keys(env);
  const fn = new Function(...names, src + '\n; return {saveRecords, updateRecord, fillFormulasDown, diagnose, queryHistory, formulaTemplate, lastRowByColumn};');
  return fn(...names.map(n => env[n]));
}

module.exports = { makeSheet, buildEnv, loadCode };
