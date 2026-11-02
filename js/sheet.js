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
    let hotelGridRow = null;
    dataRows.forEach((row, offset) => {
      const gridRow = offset + 2;
      const firstCell = app.getTextContent(row[0] || '');
      if (app.looksLikeHotelHeader(firstCell)) {
        hotelCells = row.slice(0, locations.length);
        while (hotelCells.length < locations.length) hotelCells.push('');
        hotelGridRow = gridRow;
      } else {
        // Keep row if it's a time or potential note (empty time)
        const time = firstCell.trim();
        const cells = row.slice(1);
        data.push({ time, cells, gridRow });
      }
    });
    return { locations, dates, rows: data, hotelCells, hotelGridRow };
  };

  app.planLinkedSheetEdits = function planLinkedSheetEdits(grid, days, mergesIn) {
    const requests = [];
    let merges = (mergesIn || []).map(merge => ({
      startRowIndex: merge.startRowIndex || 0,
      endRowIndex: merge.endRowIndex || 0,
      startColumnIndex: merge.startColumnIndex || 0,
      endColumnIndex: merge.endColumnIndex || 0
    })).filter(merge => merge.endRowIndex > merge.startRowIndex && merge.endColumnIndex > merge.startColumnIndex);

    const clone = (source) => source.map(row => (row || []).map(cell => (cell == null ? '' : String(cell))));
    const work = clone(grid || []);
    const server = clone(grid || []);
    if (!work.length) work.push(['Time']);
    if (!server.length) server.push(['Time']);
    const headerWidth = Math.max(work[0].length, 1);
    const pad = (rows) => {
      rows.forEach(row => { while (row.length < headerWidth) row.push(''); });
    };
    pad(work);
    pad(server);

    const rowTrans = [];
    const overlaps = (start, end, index, stop) => start < stop && index < end;

    const unmergeWhere = (hit) => {
      const keep = [];
      merges.forEach(merge => {
        if (!hit(merge)) {
          keep.push(merge);
          return;
        }
        requests.push({ unmergeCells: { range: { sheetId: 0, ...merge } } });
        for (let r = merge.startRowIndex; r < merge.endRowIndex; r++) {
          for (let c = merge.startColumnIndex; c < merge.endColumnIndex; c++) {
            if (r === merge.startRowIndex && c === merge.startColumnIndex) continue;
            if (server[r]) server[r][c] = '';
          }
        }
      });
      merges = keep;
    };

    const shiftMerges = (dimension, index, delta) => {
      merges.forEach(merge => {
        const startKey = dimension === 'COLUMNS' ? 'startColumnIndex' : 'startRowIndex';
        const endKey = dimension === 'COLUMNS' ? 'endColumnIndex' : 'endRowIndex';
        if (merge[startKey] >= index) {
          merge[startKey] += delta;
          merge[endKey] += delta;
        } else if (merge[endKey] > index) {
          merge[endKey] += delta;
        }
      });
    };

    const mapMovedIndex = (index, from, toFinal) => {
      if (index === from) return toFinal;
      if (from < toFinal && index > from && index <= toFinal) return index - 1;
      if (toFinal < from && index >= toFinal && index < from) return index + 1;
      return index;
    };

    const eachRow = (fn) => {
      const rows = Math.max(work.length, server.length);
      for (let r = 0; r < rows; r++) {
        if (!work[r]) work[r] = [];
        if (!server[r]) server[r] = [];
        fn(work[r], server[r]);
      }
    };

    const applyColumnDelete = (index) => {
      unmergeWhere(merge => overlaps(merge.startColumnIndex, merge.endColumnIndex, index, index + 1));
      requests.push({ deleteDimension: { range: { sheetId: 0, dimension: 'COLUMNS', startIndex: index, endIndex: index + 1 } } });
      eachRow((left, right) => { left.splice(index, 1); right.splice(index, 1); });
      shiftMerges('COLUMNS', index + 1, -1);
    };

    const applyColumnInsert = (index) => {
      unmergeWhere(merge => merge.startColumnIndex < index && index < merge.endColumnIndex);
      requests.push({
        insertDimension: {
          range: { sheetId: 0, dimension: 'COLUMNS', startIndex: index, endIndex: index + 1 },
          inheritFromBefore: index > 0
        }
      });
      eachRow((left, right) => {
        while (left.length < index) left.push('');
        while (right.length < index) right.push('');
        left.splice(index, 0, '');
        right.splice(index, 0, '');
      });
      shiftMerges('COLUMNS', index, 1);
    };

    const applyColumnMove = (from, toFinal) => {
      if (from === toFinal) return;
      unmergeWhere(merge => overlaps(merge.startColumnIndex, merge.endColumnIndex, from, from + 1));
      const destinationIndex = toFinal > from ? toFinal + 1 : toFinal;
      requests.push({
        moveDimension: {
          source: { sheetId: 0, dimension: 'COLUMNS', startIndex: from, endIndex: from + 1 },
          destinationIndex
        }
      });
      eachRow((left, right) => {
        const width = Math.max(left.length, right.length, from + 1, toFinal + 1);
        while (left.length < width) left.push('');
        while (right.length < width) right.push('');
        const [leftCell] = left.splice(from, 1);
        const [rightCell] = right.splice(from, 1);
        left.splice(toFinal, 0, leftCell || '');
        right.splice(toFinal, 0, rightCell || '');
      });
      merges.forEach(merge => {
        const cols = [];
        for (let c = merge.startColumnIndex; c < merge.endColumnIndex; c++) cols.push(mapMovedIndex(c, from, toFinal));
        cols.sort((a, b) => a - b);
        merge.startColumnIndex = cols[0];
        merge.endColumnIndex = cols[cols.length - 1] + 1;
      });
    };

    const applyRowInsert = (index) => {
      unmergeWhere(merge => merge.startRowIndex < index && index < merge.endRowIndex);
      requests.push({
        insertDimension: {
          range: { sheetId: 0, dimension: 'ROWS', startIndex: index, endIndex: index + 1 },
          inheritFromBefore: index > 0
        }
      });
      work.splice(index, 0, []);
      server.splice(index, 0, []);
      shiftMerges('ROWS', index, 1);
      rowTrans.push({ type: 'insert', index });
    };

    const applyRowDelete = (index) => {
      unmergeWhere(merge => overlaps(merge.startRowIndex, merge.endRowIndex, index, index + 1));
      requests.push({ deleteDimension: { range: { sheetId: 0, dimension: 'ROWS', startIndex: index, endIndex: index + 1 } } });
      work.splice(index, 1);
      server.splice(index, 1);
      shiftMerges('ROWS', index + 1, -1);
      rowTrans.push({ type: 'delete', index });
    };

    const mapRow = (row) => {
      let next = row;
      for (let i = 0; i < rowTrans.length; i++) {
        const op = rowTrans[i];
        if (op.type === 'insert' && next >= op.index) next++;
        else if (op.type === 'delete') {
          if (next === op.index) return -1;
          if (next > op.index) next--;
        }
      }
      return next;
    };

    const text = (row, col) => String((row && row[col]) || '').trim();
    const ensure = (rows, r, c) => {
      while (rows.length <= r) rows.push([]);
      while (rows[r].length <= c) rows[r].push('');
    };
    const put = (r, c, value) => {
      ensure(work, r, c);
      work[r][c] = value;
    };

    let tokens = [];
    for (let c = 1; c < headerWidth; c++) tokens.push(c);
    const used = new Set();
    const desired = (days || []).map(day => {
      const col = day.sourceCol == null || day.sourceCol === '' ? NaN : Number(day.sourceCol);
      if (Number.isInteger(col) && col >= 1 && col < headerWidth && !used.has(col)) {
        used.add(col);
        return col;
      }
      return null;
    });

    for (let i = tokens.length - 1; i >= 0; i--) {
      if (used.has(tokens[i])) continue;
      applyColumnDelete(i + 1);
      tokens.splice(i, 1);
    }
    for (let i = 0; i < desired.length; i++) {
      const want = desired[i];
      if (want == null) {
        applyColumnInsert(i + 1);
        tokens.splice(i, 0, null);
        continue;
      }
      const j = tokens.indexOf(want);
      if (j !== i) {
        applyColumnMove(j + 1, i + 1);
        const [item] = tokens.splice(j, 1);
        tokens.splice(i, 0, item);
      }
    }

    const clock = (label) => app.timeToInputValue(String(label || '').trim());
    const isHotel = (rows, r) => app.looksLikeHotelHeader(text(rows[r], 0));
    const classify = (rows, col) => {
      let lastTimed = -1;
      for (let r = 2; r < rows.length; r++) {
        if (isHotel(rows, r)) continue;
        if (text(rows[r], col) && text(rows[r], 0) && !app.looksLikeHotelHeader(text(rows[r], 0))) lastTimed = r;
      }
      const activityRows = [];
      const noteRows = [];
      for (let r = 2; r < rows.length; r++) {
        if (isHotel(rows, r)) continue;
        const content = text(rows[r], col);
        if (!content) continue;
        const time = text(rows[r], 0);
        if (r <= lastTimed && time && !app.looksLikeHotelHeader(time)) activityRows.push(r);
        else noteRows.push(r);
      }
      return { activityRows, noteRows };
    };

    const snapshots = desired.map((_, index) => classify(work, index + 1));
    const sameDate = (cellText, ymd) => {
      const current = String(cellText || '').trim();
      if (!current && !ymd) return true;
      const parsed = app.parseSmartDate(current);
      if (parsed && ymd && parsed === ymd) return true;
      return current === app.formatTripDate(ymd);
    };

    desired.forEach((_, index) => {
      const day = days[index];
      const col = index + 1;
      if (!sameDate(text(work[1], col), day.date || '')) put(1, col, app.formatTripDate(day.date || ''));
      if (text(work[0], col) !== String(day.location || '').trim()) put(0, col, String(day.location || '').trim());
    });

    let hotelRow = -1;
    for (let r = 2; r < work.length; r++) if (isHotel(work, r)) hotelRow = r;
    const anyHotel = (days || []).some(day => String(day.hotel || '').trim());
    if (hotelRow < 0 && anyHotel) {
      applyRowInsert(2);
      put(2, 0, 'Hotel');
      hotelRow = 2;
    } else if (hotelRow >= 0) {
      hotelRow = mapRow(hotelRow);
    }
    if (hotelRow >= 0) {
      desired.forEach((_, index) => {
        const next = String(days[index].hotel || '').trim();
        if (text(work[hotelRow], index + 1) !== next) put(hotelRow, index + 1, next);
      });
    }

    const findTimeRow = (col, hhmm) => {
      for (let r = 2; r < work.length; r++) {
        if (isHotel(work, r)) continue;
        if (clock(text(work[r], 0)) !== hhmm) continue;
        if (!text(work[r], col)) return r;
      }
      return -1;
    };

    const insertTimeRow = (hhmm) => {
      const minutes = app.timeToMinutes(hhmm);
      const timed = [];
      for (let r = 2; r < work.length; r++) {
        if (isHotel(work, r)) continue;
        if (clock(text(work[r], 0))) timed.push(r);
      }
      let insertAt = null;
      for (let i = 0; i < timed.length; i++) {
        const rowMinutes = app.timeToMinutes(clock(text(work[timed[i]], 0)));
        if (rowMinutes <= minutes) insertAt = timed[i] + 1;
        else if (insertAt == null) {
          insertAt = timed[i];
          break;
        } else break;
      }
      if (insertAt == null) insertAt = hotelRow >= 0 ? hotelRow + 1 : 2;
      applyRowInsert(insertAt);
      if (hotelRow >= insertAt) hotelRow++;
      put(insertAt, 0, app.formatTimeLabel(hhmm));
      return insertAt;
    };

    desired.forEach((_, index) => {
      const day = days[index];
      const col = index + 1;
      const snapshot = snapshots[index];
      const usedRows = new Set();
      const pending = [];
      const activities = [];
      const noteLines = String(day.notes || '').split('\n').map(line => line.trim()).filter(Boolean);
      (day.activities || []).forEach(act => {
        const content = String(act.content || '').trim();
        if (!content) return;
        if (!act.time) {
          noteLines.push(content);
          return;
        }
        activities.push({ ...act, content, time: act.time });
      });

      activities.forEach(act => {
        const source = act.sourceRow == null || act.sourceRow === '' ? NaN : Number(act.sourceRow);
        const known = Number.isInteger(source) && snapshot.activityRows.includes(source);
        const mapped = known ? mapRow(source) : -1;
        const sameTime = mapped >= 0 && clock(text(work[mapped], 0)) === act.time;
        if (known && sameTime) {
          usedRows.add(source);
          if (text(work[mapped], col) !== act.content) put(mapped, col, act.content);
          return;
        }
        if (!known) {
          const exact = snapshot.activityRows.find(row => {
            if (usedRows.has(row)) return false;
            const at = mapRow(row);
            return at >= 0 && clock(text(work[at], 0)) === act.time && text(work[at], col) === act.content;
          });
          if (exact != null) {
            usedRows.add(exact);
            return;
          }
        }
        pending.push(act);
      });

      snapshot.activityRows.forEach(row => {
        if (usedRows.has(row)) return;
        const at = mapRow(row);
        if (at >= 0) put(at, col, '');
      });

      pending.forEach(act => {
        let row = findTimeRow(col, act.time);
        if (row < 0) row = insertTimeRow(act.time);
        put(row, col, act.content);
      });

      const originalNotes = snapshot.noteRows.map(row => text(work[mapRow(row)], col)).filter(Boolean).join('\n');
      if (originalNotes === noteLines.join('\n')) return;
      const targets = snapshot.noteRows.map(row => mapRow(row)).filter(row => row >= 0);
      const claim = (line) => {
        for (let r = 2; r < work.length; r++) {
          if (isHotel(work, r) || text(work[r], 0) || text(work[r], col)) continue;
          if (targets.includes(r)) continue;
          put(r, col, line);
          targets.push(r);
          return;
        }
        const at = work.length;
        applyRowInsert(at);
        put(at, col, line);
        targets.push(at);
      };
      noteLines.forEach((line, lineIndex) => {
        if (lineIndex < targets.length) {
          if (text(work[targets[lineIndex]], col) !== line) put(targets[lineIndex], col, line);
        } else claim(line);
      });
      for (let lineIndex = noteLines.length; lineIndex < targets.length; lineIndex++) put(targets[lineIndex], col, '');
    });

    const body = (rows, r) => (rows[r] || []).slice(1).some(cell => String(cell || '').trim());
    const meaningful = (rows, r) => (rows[r] || []).some(cell => String(cell || '').trim());
    for (let r = work.length - 1; r >= 2; r--) {
      if (body(work, r) || !meaningful(server, r)) continue;
      applyRowDelete(r);
    }

    merges.slice().forEach(merge => {
      let sample = null;
      let uniform = true;
      for (let r = merge.startRowIndex; r < merge.endRowIndex && uniform; r++) {
        for (let c = merge.startColumnIndex; c < merge.endColumnIndex; c++) {
          const value = text(work[r], c);
          if (sample == null) sample = value;
          else if (sample !== value) uniform = false;
        }
      }
      if (uniform) {
        for (let r = merge.startRowIndex; r < merge.endRowIndex; r++) {
          for (let c = merge.startColumnIndex; c < merge.endColumnIndex; c++) {
            if (r === merge.startRowIndex && c === merge.startColumnIndex) continue;
            ensure(server, r, c);
            server[r][c] = text(work[r], c);
          }
        }
        return;
      }
      requests.push({ unmergeCells: { range: { sheetId: 0, ...merge } } });
      merges = merges.filter(item => item !== merge);
      for (let r = merge.startRowIndex; r < merge.endRowIndex; r++) {
        for (let c = merge.startColumnIndex; c < merge.endColumnIndex; c++) {
          if (r === merge.startRowIndex && c === merge.startColumnIndex) continue;
          ensure(server, r, c);
          server[r][c] = '';
        }
      }
    });

    const cells = [];
    const rowCount = Math.max(server.length, work.length);
    for (let r = 0; r < rowCount; r++) {
      const colCount = Math.max((server[r] || []).length, (work[r] || []).length);
      for (let c = 0; c < colCount; c++) {
        if (text(server[r], c) === text(work[r], c)) continue;
        cells.push({ row: r, col: c, value: text(work[r], c) });
      }
    }

    while (work.length && work[work.length - 1].every(cell => !String(cell || '').trim())) work.pop();
    return { requests, cells, grid: work, merges };
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

  app.bookendColumns = function bookendColumns(entries) {
    const runs = [];
    (entries || []).forEach(entry => {
      const loc = app.getTextContent(entry && entry.location || '').trim();
      if (!loc || entry.id == null) return;
      const key = loc.toLowerCase();
      const prev = runs[runs.length - 1];
      if (prev && prev.key === key) prev.ids.push(entry.id);
      else runs.push({ key, ids: [entry.id] });
    });
    if (runs.length < 3) return new Set();
    const first = runs[0];
    const last = runs[runs.length - 1];
    if (first.ids.length !== 1 || last.ids.length !== 1 || first.key !== last.key) return new Set();
    return new Set([first.ids[0], last.ids[0]]);
  };

  app.destinationNames = function destinationNames(days) {
    const entries = (days || []).map((day, index) => ({ location: day && day.location, id: index }));
    const bookends = app.bookendColumns(entries);
    const seen = new Set();
    const out = [];
    entries.forEach(entry => {
      if (bookends.has(entry.id)) return;
      const loc = app.getTextContent(entry.location || '').trim();
      if (!loc) return;
      const key = loc.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      out.push(loc);
    });
    return out;
  };

  app.parsedDestinations = function parsedDestinations() {
    if (!app.parsed) return [];
    const entries = [];
    for (let i = 1; i < app.parsed.locations.length; i++) {
      const location = app.getTextContent(app.parsed.locations[i] || '').trim();
      const date = app.getTextContent(app.parsed.dates[i] || '').trim();
      if (!location || !date) continue;
      entries.push({ location, id: i });
    }
    const bookends = app.bookendColumns(entries);
    const seen = new Set();
    const out = [];
    entries.forEach(entry => {
      if (bookends.has(entry.id)) return;
      const key = entry.location.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      out.push(entry.location);
    });
    return out;
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
