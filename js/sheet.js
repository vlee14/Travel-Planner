// Spreadsheet and CSV parsing.
export function installSheet(app) {
  app.parseCSV = function parseCSV(text) {
    // Detect delimiter properly by finding the first *unquoted* newline
    let firstRowEnd = text.length;
    let q = false;
    for (let i = 0; i < text.length; i++) {
      if (text[i] === '"') q = !q;
      else if (!q && (text[i] === '\n' || text[i] === '\r')) {
        firstRowEnd = i;
        break;
      }
    }

    const firstLine = text.slice(0, firstRowEnd);
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    const delimiter = commaCount > tabCount ? ',' : '\t';

    const rows = [];
    let current = '';
    let inQuotes = false;
    const row = [];
    const pushRow = () => {
      rows.push(row.slice());
      row.length = 0;
    };
    const pushCell = () => {
      row.push(current.trim());
      current = '';
    };
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      const next = text[i + 1];

      if (inQuotes) {
        if (c === '"') {
          if (next === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += c;
        }
        continue;
      }

      if (c === '"') {
        inQuotes = true;
        continue;
      }

      if (c === delimiter) {
        pushCell();
        continue;
      }

      if (c === '\n' || c === '\r') {
        pushCell();
        pushRow();
        if (c === '\r' && text[i + 1] === '\n') i++;
        continue;
      }
      current += c;
    }

    if (current || row.length > 0) {
      pushCell();
      pushRow();
    }

    return rows;
  };

  app.looksLikeHotelHeader = function looksLikeHotelHeader(str) {
    const s = (str || '').trim();
    return /\bhotel\b/i.test(s);
  };

  app.parseSheetRows = function parseSheetRows(rows) {
    if (!rows.length) return null;
    const first = rows[0];
    const locations = first.map((c, i) => (i === 0 ? app.getTextContent(c || 'Time') : app.getTextContent(c || `Location ${i}`)));
    const dates = rows[1] ? rows[1].map((c, i) => (i === 0 ? '' : app.getTextContent(c || ''))) : first.map(() => '');
    const dataRows = rows.slice(2);
    let hotelCells = null;
    const data = [];
    dataRows.forEach(row => {
      const firstCell = app.getTextContent(row[0] || '');
      if (app.looksLikeHotelHeader(firstCell)) {
        hotelCells = row.slice(0, locations.length);
        while (hotelCells.length < locations.length) hotelCells.push('');
      } else {
        // Keep row if it's a time or potential note (empty time)
        const time = firstCell.trim();
        const cells = row.slice(1);
        data.push({ time, cells });
      }
    });
    return { locations, dates, rows: data, hotelCells };
  };

  app.rowsToCsv = function rowsToCsv(rows) {
    const toCsvValue = (val) => {
      const v = (val || '').toString();
      if (/[\t\n\r"]/.test(v)) return '"' + v.replace(/"/g, '""') + '"';
      return v;
    };
    return rows.map(row => row.map(toCsvValue).join('\t')).join('\n');
  };

  app.parsedToGrid = function parsedToGrid(model) {
    const grid = [
      model.locations.map(cell => app.getTextContent(cell || '')),
      model.dates.map(cell => app.getTextContent(cell || ''))
    ];
    if (model.hotelCells) grid.push(model.hotelCells.map(cell => app.getTextContent(cell || '')));
    model.rows.forEach(row => {
      const line = [row.time || ''];
      for (let i = 1; i < model.locations.length; i++) line.push(app.getTextContent(row.cells[i - 1] || ''));
      grid.push(line);
    });
    return grid;
  };

  app.findDayColumn = function findDayColumn(location, dateLabel) {
    if (!app.parsed) return -1;
    for (let i = 1; i < app.parsed.locations.length; i++) {
      const loc = app.getTextContent(app.parsed.locations[i] || '').trim();
      const date = app.getTextContent(app.parsed.dates[i] || '').trim();
      if (loc === location && date === dateLabel) return i;
    }
    return -1;
  };
}
