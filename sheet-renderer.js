/**
 * Content Brief Renderer — Google Sheets API v4
 * Builds a batchUpdate request list that matches the "roulette online" tab
 * layout from NewExample_betonline.ag - Content Briefs _ Internal.xlsx.
 *
 * Entry point: buildBriefRequests(sheetId, job) -> { requests, postRequests }
 *
 * `requests` is applied in one batchUpdate call. `postRequests` contains things
 * that must run *after* values exist (currently nothing, but reserved).
 *
 * The `sheetId` passed in is the gid of the target sheet inside the spreadsheet.
 */

/* ============================== CONSTANTS ============================== */
const C = {
  blue:       '#182F7A',
  altLight:   '#F8FAFC',
  white:      '#FFFFFF',
  labelGrey:  '#334155',
  textDark:   '#0F172A',
  greyHdr:    '#E2E8F0',
  trust:      '#6AA84F',
  benefits:   '#3D85C6',
  pain:       '#CC0000',
  csRed:      '#F4CCCC',
  csLabel:    '#64748B',
  cyanLabel:  '#04DBF5',
  cyanHdr:    '#DEFBFF',
  questBlue:  '#E0F2FE',
  questText:  '#0C4A6E',
  link:       '#0000FF',
  black:      '#000000'
};
const TOTAL_COLS = 11;   // A..K
const OUTLINE_COLS = 10; // B..K
const MAIN_COLS = 6;     // B..G
const FONT_SECTION = 'Lexend';
const FONT_BODY = 'Arial';

/* ============================== HELPERS ============================== */
function hexToRgb(hex) {
  const h = hex.replace('#', '');
  return {
    red:   parseInt(h.substring(0, 2), 16) / 255,
    green: parseInt(h.substring(2, 4), 16) / 255,
    blue:  parseInt(h.substring(4, 6), 16) / 255
  };
}
function nl_(v) {
  if (v == null) return '';
  if (typeof v !== 'string') return String(v);
  return v.replace(/\\n\\n/g, '\n\n').replace(/\\n/g, '\n');
}
function bullets_(v) {
  if (v == null) return '';
  if (typeof v !== 'string') return String(v);
  const lines = v.split('\n');
  const re = /^\s*[•\-—\*]\s*/;
  const has = lines.some(l => re.test(l));
  return lines.map(l => {
    if (!l.trim()) return l;
    if (re.test(l)) return l.replace(re, '• ');
    return has ? l : '• ' + l;
  }).join('\n');
}
function toBool_(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string') return /^true$/i.test(v.trim());
  return false;
}
function get_(it, ...keys) {
  for (const k of keys) if (it && it[k] != null) return it[k];
  return '';
}

/* ============================== REQUEST BUILDERS ============================== */

/**
 * Range helper: rows/cols are 1-indexed in our API surface, but Sheets API
 * uses 0-indexed half-open ranges. We convert here.
 */
function gridRange(sheetId, r1, c1, numRows, numCols) {
  return {
    sheetId,
    startRowIndex: r1 - 1,
    endRowIndex: r1 - 1 + numRows,
    startColumnIndex: c1 - 1,
    endColumnIndex: c1 - 1 + numCols
  };
}

/** Build a userEnteredFormat object. */
function fmt(opts = {}) {
  const f = {};
  if (opts.bg) f.backgroundColor = hexToRgb(opts.bg);
  const tf = {};
  if (opts.fontFamily) tf.fontFamily = opts.fontFamily;
  if (opts.fontSize != null) tf.fontSize = opts.fontSize;
  if (opts.bold != null) tf.bold = !!opts.bold;
  if (opts.italic != null) tf.italic = !!opts.italic;
  if (opts.color) tf.foregroundColor = hexToRgb(opts.color);
  if (Object.keys(tf).length) f.textFormat = tf;
  if (opts.hAlign) f.horizontalAlignment = opts.hAlign;
  if (opts.vAlign) f.verticalAlignment = opts.vAlign;
  if (opts.wrap) f.wrapStrategy = opts.wrap;
  if (opts.numberFormat) f.numberFormat = opts.numberFormat;
  return f;
}

/** repeatCell: apply the same userEnteredFormat to an entire range. */
function repeatFormat(sheetId, r1, c1, numRows, numCols, format) {
  // Build fields mask from format keys
  const fields = Object.keys(format)
    .map(k => k === 'textFormat'
      ? 'userEnteredFormat.textFormat'
      : 'userEnteredFormat.' + k)
    .join(',');
  return {
    repeatCell: {
      range: gridRange(sheetId, r1, c1, numRows, numCols),
      cell: { userEnteredFormat: format },
      fields: fields || 'userEnteredFormat'
    }
  };
}

/** Merge a range. */
function merge(sheetId, r1, c1, numRows, numCols, type = 'MERGE_ALL') {
  return {
    mergeCells: {
      range: gridRange(sheetId, r1, c1, numRows, numCols),
      mergeType: type
    }
  };
}

/** Unmerge a range. */
function unmerge(sheetId, r1, c1, numRows, numCols) {
  return {
    unmergeCells: { range: gridRange(sheetId, r1, c1, numRows, numCols) }
  };
}

/** Write a single cell with userEnteredValue + optional format. */
function writeCell(sheetId, row, col, value, format) {
  const uv = {};
  if (typeof value === 'number') uv.numberValue = value;
  else if (typeof value === 'boolean') uv.boolValue = value;
  else if (typeof value === 'string' && value.startsWith('=')) uv.formulaValue = value;
  else uv.stringValue = value == null ? '' : String(value);

  const cell = { userEnteredValue: uv };
  let fields = 'userEnteredValue';
  if (format) {
    cell.userEnteredFormat = format;
    fields += ',userEnteredFormat';
  }
  return {
    updateCells: {
      range: gridRange(sheetId, row, col, 1, 1),
      rows: [{ values: [cell] }],
      fields
    }
  };
}

/** Write a cell with a hyperlink (renders as clickable blue text). */
function writeLinkCell(sheetId, row, col, text, url, format) {
  if (!url) return writeCell(sheetId, row, col, text, format);
  const formulaValue = `=HYPERLINK("${url.replace(/"/g, '\\"')}","${String(text || url).replace(/"/g, '\\"')}")`;
  return writeCell(sheetId, row, col, formulaValue, format);
}

/** Row height. */
function setRowHeight(sheetId, row, px) {
  return {
    updateDimensionProperties: {
      range: { sheetId, dimension: 'ROWS', startIndex: row - 1, endIndex: row },
      properties: { pixelSize: px },
      fields: 'pixelSize'
    }
  };
}

/** Column width. */
function setColWidth(sheetId, col, px) {
  return {
    updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: col - 1, endIndex: col },
      properties: { pixelSize: px },
      fields: 'pixelSize'
    }
  };
}

/** Borders — updateBorders request. */
function setBorder(sheetId, r1, c1, numRows, numCols, sides, style = 'SOLID', color = C.black) {
  const b = { style, color: hexToRgb(color) };
  const req = { updateBorders: { range: gridRange(sheetId, r1, c1, numRows, numCols) } };
  sides.split(',').forEach(s => { req.updateBorders[s.trim()] = b; });
  return req;
}

/** Data validation (checkbox). */
function checkboxRule(sheetId, r1, c1, numRows, numCols) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, c1, numRows, numCols),
      rule: {
        condition: { type: 'BOOLEAN' },
        strict: true
      }
    }
  };
}

/** Alternate fill helper: append requests that repeat bg color per row. */
function altFillRequests(sheetId, r1, r2, c1, c2, startLight) {
  const out = [];
  for (let r = r1; r <= r2; r++) {
    const isLight = ((r - r1) % 2 === 0) ? startLight : !startLight;
    out.push(repeatFormat(sheetId, r, c1, 1, c2 - c1 + 1, { backgroundColor: hexToRgb(isLight ? C.altLight : C.white) }));
  }
  return out;
}

/* ============================== RESET / GRID PRIME ============================== */
/** Clear formatting, unmerge everything, clear values (keep sheet). */
function resetSheetRequests(sheetId) {
  const req = [];
  // Unmerge all
  req.push({ unmergeCells: { range: { sheetId } } });
  // Clear formatting + values across a wide range
  req.push({
    updateCells: {
      range: { sheetId, startRowIndex: 0, endRowIndex: 300, startColumnIndex: 0, endColumnIndex: TOTAL_COLS },
      fields: 'userEnteredValue,userEnteredFormat,dataValidation,textFormatRuns'
    }
  });
  return req;
}

function primeGridRequests(sheetId) {
  const widths = [
    [1, 33], [2, 184], [3, 70], [4, 370], [5, 184], [6, 70],
    [7, 70], [8, 218], [9, 184], [10, 80], [11, 80]
  ];
  const out = widths.map(([col, px]) => setColWidth(sheetId, col, px));
  out.push({
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount'
    }
  });
  return out;
}

/* ============================== RENDERERS ============================== */

/** Title bar rows 1–2 — A1:K2 blue, B1:G2 merged title, logo via =IMAGE() in H1:K2. */
function renderTitleBar(sheetId, meta) {
  const out = [];
  // Blue bg A1:K2
  out.push(repeatFormat(sheetId, 1, 1, 2, TOTAL_COLS, { backgroundColor: hexToRgb(C.blue) }));

  // Merge title B1:G2
  out.push(merge(sheetId, 1, 2, 2, MAIN_COLS));
  const home = (meta && meta.homeUrl) ? String(meta.homeUrl).replace(/^https?:\/\//i, '') : '';
  out.push(writeCell(sheetId, 1, 2, 'Content Brief: ' + home, fmt({
    bg: C.blue, fontFamily: FONT_SECTION, fontSize: 24, bold: true, color: C.white,
    hAlign: 'LEFT', vAlign: 'MIDDLE'
  })));

  // Logo in H1:K2 via IMAGE formula (mode 1 = fit to cell preserving aspect)
  out.push(merge(sheetId, 1, 8, 2, 4));
  const logoUrl = 'https://cdn-amehi.nitrocdn.com/ldNPGLQtVWaqliEfWebnqecfYajgRCdk/assets/images/optimized/rev-7ef5dc9/www.digitalspotlight.com/wp-content/uploads/2017/09/logo-new.png';
  out.push(writeCell(sheetId, 1, 8, `=IMAGE("${logoUrl}", 1)`, fmt({
    bg: C.blue, hAlign: 'CENTER', vAlign: 'MIDDLE'
  })));

  out.push(setRowHeight(sheetId, 1, 33));
  out.push(setRowHeight(sheetId, 2, 33));
  return { requests: out, nextRow: 3 };
}

/** Overview — 3 rows, labels in B/E, values in C/F, alt bg rows. */
function renderOverview(sheetId, startRow, meta) {
  const out = [];
  const r = startRow;
  const labels = [
    ['MAIN KEYWORD:', 'KEYWORD VOLUME:'],
    ['RECOMMENDED URL:', 'TARGET WORD COUNT:'],
    ['NEW/EXISTING PAGE:', 'CONTENT TYPE:']
  ];
  const values = [
    [meta.mainKeyword || '', meta.keywordVolume || ''],
    [meta.recommendedUrl || '', meta.minWordCount || ''],
    [meta.existingOrNew || '', meta.pageType || '']
  ];

  // Alt row backgrounds across A..K
  out.push(repeatFormat(sheetId, r, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.altLight) }));
  out.push(repeatFormat(sheetId, r + 1, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.white) }));
  out.push(repeatFormat(sheetId, r + 2, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.altLight) }));

  const labelFmt = fmt({ fontFamily: FONT_BODY, fontSize: 12, bold: true, color: C.labelGrey, vAlign: 'MIDDLE' });
  const valueFmt = fmt({ fontFamily: FONT_BODY, fontSize: 12, color: C.textDark, vAlign: 'MIDDLE' });

  for (let i = 0; i < 3; i++) {
    out.push(writeCell(sheetId, r + i, 2, labels[i][0], labelFmt));
    out.push(writeCell(sheetId, r + i, 5, labels[i][1], labelFmt));
    if (i === 1) {
      // Recommended URL — hyperlink
      out.push(writeLinkCell(sheetId, r + i, 3, nl_(values[i][0]), values[i][0], valueFmt));
    } else {
      out.push(writeCell(sheetId, r + i, 3, nl_(values[i][0]), valueFmt));
    }
    out.push(writeCell(sheetId, r + i, 6, nl_(values[i][1]), valueFmt));
  }

  for (let i = 0; i < 3; i++) out.push(setRowHeight(sheetId, r + i, 49));
  return { requests: out, nextRow: r + 3 + 1 };
}

/** Big blue section header (2 rows, B:G merged). */
function renderBlueHeader(sheetId, row, text, height = 35) {
  const out = [];
  out.push(repeatFormat(sheetId, row, 1, 2, TOTAL_COLS, { backgroundColor: hexToRgb(C.blue) }));
  out.push(merge(sheetId, row, 2, 2, MAIN_COLS));
  out.push(writeCell(sheetId, row, 2, text || '', fmt({
    bg: C.blue, fontFamily: FONT_SECTION, fontSize: 24, bold: true, color: C.white,
    hAlign: 'LEFT', vAlign: 'MIDDLE'
  })));
  out.push(setRowHeight(sheetId, row, height));
  out.push(setRowHeight(sheetId, row + 1, height));
  return { requests: out, nextRow: row + 2 };
}

/** Compact 1-row blue header (used by Top 10). */
function renderBlueHeaderSmall(sheetId, row, text) {
  const out = [];
  out.push(repeatFormat(sheetId, row, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.blue) }));
  out.push(merge(sheetId, row, 2, 1, MAIN_COLS));
  out.push(writeCell(sheetId, row, 2, text || '', fmt({
    bg: C.blue, fontFamily: FONT_BODY, fontSize: 11, bold: true, color: C.white,
    hAlign: 'LEFT', vAlign: 'MIDDLE'
  })));
  out.push(setRowHeight(sheetId, row, 31));
  return { requests: out, nextRow: row + 1 };
}

/** Content Outline — blue header, grey column headers, 10 data columns. */
function renderContentOutline(sheetId, startRow, items) {
  const out = [];
  const bh = renderBlueHeader(sheetId, startRow, 'Content Outline');
  out.push(...bh.requests);
  let r = bh.nextRow;

  // Header row
  const headers = ['Section', 'Heading', 'Writer Instructions', 'Type', 'Capsule?',
                   'Word Target', 'Required Elements', 'Entities / Terms', 'WRITER ✓', 'EDITOR ✓'];
  out.push(repeatFormat(sheetId, r, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.greyHdr) }));
  const hdrFmt = fmt({
    bg: C.greyHdr, fontFamily: FONT_BODY, fontSize: 11, bold: true, color: C.labelGrey,
    hAlign: 'CENTER', vAlign: 'MIDDLE', wrap: 'WRAP'
  });
  for (let i = 0; i < headers.length; i++) {
    out.push(writeCell(sheetId, r, 2 + i, headers[i], hdrFmt));
  }
  out.push(setBorder(sheetId, r, 1, 1, TOTAL_COLS, 'bottom', 'SOLID_MEDIUM'));
  out.push(setRowHeight(sheetId, r, 35));
  r++;

  const rows = Array.isArray(items) ? items : [];
  if (!rows.length) return { requests: out, nextRow: r + 1 };

  const bodyStart = r;
  const bodyFmtBase = {
    fontFamily: FONT_BODY, fontSize: 12, color: C.textDark, vAlign: 'MIDDLE', wrap: 'WRAP'
  };

  for (let i = 0; i < rows.length; i++, r++) {
    const it = rows[i] || {};
    const section  = nl_(get_(it, 'Section', 'section'));
    const heading  = nl_(get_(it, 'Heading', 'heading'));
    const writer   = bullets_(nl_(get_(it, 'Writer Instructions', 'writerInstructions', 'reqs')));
    const type     = nl_(get_(it, 'Type', 'type'));
    const capsule  = nl_(get_(it, 'Capsule?', 'capsule'));
    const wordT    = nl_(get_(it, 'Word Target', 'wordTarget'));
    const reqElem  = bullets_(nl_(get_(it, 'Required Elements', 'requiredElements')));
    const entities = nl_(get_(it, 'Entities / Terms', 'entities'));
    const w        = toBool_(get_(it, 'WRITER ✓', 'writer'));
    const e        = toBool_(get_(it, 'EDITOR ✓', 'editor'));

    out.push(writeCell(sheetId, r, 2, section, fmt({ ...bodyFmtBase, bold: true, hAlign: 'CENTER' })));
    out.push(writeCell(sheetId, r, 3, heading, fmt({ ...bodyFmtBase, hAlign: 'LEFT' })));
    out.push(writeCell(sheetId, r, 4, writer, fmt({ ...bodyFmtBase, hAlign: 'LEFT' })));
    out.push(writeCell(sheetId, r, 5, type, fmt({ ...bodyFmtBase, hAlign: 'CENTER' })));
    out.push(writeCell(sheetId, r, 6, capsule, fmt({ ...bodyFmtBase, hAlign: 'CENTER' })));
    out.push(writeCell(sheetId, r, 7, wordT, fmt({ ...bodyFmtBase, hAlign: 'CENTER' })));
    out.push(writeCell(sheetId, r, 8, reqElem, fmt({ ...bodyFmtBase, hAlign: 'LEFT' })));
    out.push(writeCell(sheetId, r, 9, entities, fmt({ ...bodyFmtBase, hAlign: 'LEFT' })));

    // Checkboxes
    out.push(checkboxRule(sheetId, r, 10, 1, 2));
    out.push(writeCell(sheetId, r, 10, w, fmt({ ...bodyFmtBase, hAlign: 'CENTER' })));
    out.push(writeCell(sheetId, r, 11, e, fmt({ ...bodyFmtBase, hAlign: 'CENTER' })));
  }

  // Alternate fill across A..K for body
  out.push(...altFillRequests(sheetId, bodyStart, r - 1, 1, TOTAL_COLS, true));

  return { requests: out, nextRow: r + 1 };
}

/** Trust / Benefits / Pain Points. */
function renderTrustBenefits(sheetId, startRow, trust, benefits, pain) {
  const out = [];
  let r = startRow;

  // Label row — 3 merged cells B:C, D:E, F:G
  const labelDefs = [
    [2, 'TRUST ELEMENTS', C.trust],
    [4, 'Benefits, Offers & CTAs', C.benefits],
    [6, 'PAIN POINTS', C.pain]
  ];
  labelDefs.forEach(([col, text, bg]) => {
    out.push(merge(sheetId, r, col, 1, 2));
    out.push(writeCell(sheetId, r, col, text, fmt({
      bg, fontFamily: FONT_BODY, fontSize: 14, bold: true, color: C.white,
      hAlign: 'CENTER', vAlign: 'MIDDLE'
    })));
  });
  out.push(setRowHeight(sheetId, r, 35));
  r++;

  // Body: 6 merged rows per column group
  const bodyFmt = fmt({
    bg: C.white, fontFamily: FONT_BODY, fontSize: 10, color: C.labelGrey,
    vAlign: 'TOP', hAlign: 'LEFT', wrap: 'WRAP'
  });
  const bodyDefs = [
    [2, bullets_(nl_(trust || ''))],
    [4, bullets_(nl_(benefits || ''))],
    [6, bullets_(nl_(pain || ''))]
  ];
  bodyDefs.forEach(([col, val]) => {
    out.push(merge(sheetId, r, col, 6, 2));
    out.push(writeCell(sheetId, r, col, val, bodyFmt));
  });
  for (let i = 0; i < 6; i++) out.push(setRowHeight(sheetId, r + i, 35));

  // Thin outer border around each of the 3 box column groups
  // (label row + 6 body rows). The label row is at r - 1 since we've
  // already incremented past it.
  const boxTop = r - 1;
  const boxHeight = 7; // 1 label row + 6 body rows
  [2, 4, 6].forEach(col => {
    out.push(setBorder(sheetId, boxTop, col, boxHeight, 2,
      'top,bottom,left,right', 'SOLID'));
  });

  return { requests: out, nextRow: r + 6 + 1 };
}

/** Resources & Technical Details. */
function renderResourcesTech(sheetId, startRow, meta, tables) {
  const out = [];
  const bh = renderBlueHeader(sheetId, startRow, 'RESOURCES & TECHNICAL DETAILS');
  out.push(...bh.requests);
  let r = bh.nextRow;

  // SEO Terms Details label
  out.push(merge(sheetId, r, 2, 1, MAIN_COLS));
  out.push(writeCell(sheetId, r, 2, 'SEO Terms Details', fmt({
    fontFamily: FONT_BODY, fontSize: 11, bold: true, color: C.textDark,
    hAlign: 'LEFT', vAlign: 'BOTTOM'
  })));
  out.push(setRowHeight(sheetId, r, 31));
  r++;

  // Clearscope strip (3 rows red bg)
  out.push(repeatFormat(sheetId, r, 1, 3, TOTAL_COLS, { backgroundColor: hexToRgb(C.csRed) }));

  const csLabelFmt = fmt({ bg: C.csRed, fontFamily: FONT_BODY, fontSize: 10, color: C.csLabel, vAlign: 'BOTTOM', wrap: 'WRAP' });
  const csValueFmt = fmt({ bg: C.csRed, fontFamily: FONT_BODY, fontSize: 10, color: C.textDark, vAlign: 'BOTTOM' });

  // Row 1
  out.push(writeCell(sheetId, r, 2, 'Clearscope Report Link:', csLabelFmt));
  out.push(writeLinkCell(sheetId, r, 3, meta.clearScopeLink || '', meta.clearScopeLink, fmt({
    bg: C.csRed, fontFamily: FONT_BODY, fontSize: 10, color: C.link, vAlign: 'BOTTOM'
  })));
  out.push(writeCell(sheetId, r, 4, 'Clearscope Grade (writer):', csLabelFmt));
  out.push(writeCell(sheetId, r, 5, (tables && tables.clearscopeWriterGrade) || '', csValueFmt));
  out.push(writeCell(sheetId, r, 6, 'Clearscope Grade (editor/QA):', csLabelFmt));
  out.push(writeCell(sheetId, r, 7, (tables && tables.clearscopeEditorGrade) || '', csValueFmt));
  out.push(setRowHeight(sheetId, r, 27));
  r++;

  // Row 2 — readability
  out.push(writeCell(sheetId, r, 4, 'Clearscope Readability Grade (writer):', csLabelFmt));
  out.push(writeCell(sheetId, r, 5, (tables && tables.clearscopeWriterReadability) || '', csValueFmt));
  out.push(writeCell(sheetId, r, 6, 'Clearscope Readability Grade (editor/QA):', csLabelFmt));
  out.push(writeCell(sheetId, r, 7, (tables && tables.clearscopeEditorReadability) || '', csValueFmt));
  out.push(setRowHeight(sheetId, r, 31));
  r++;

  // Row 3 — spacer red row
  out.push(setRowHeight(sheetId, r, 22));
  r++;

  // Spacer row
  r++;

  // Target GEO row
  out.push(writeCell(sheetId, r, 2, 'TARGET GEO:', fmt({
    fontFamily: FONT_BODY, fontSize: 11, color: C.csLabel, vAlign: 'MIDDLE'
  })));
  out.push(writeCell(sheetId, r, 3, nl_(meta.targetGeo || ''), fmt({
    fontFamily: FONT_BODY, fontSize: 11, color: C.labelGrey, vAlign: 'MIDDLE'
  })));
  r++;

  return { requests: out, nextRow: r + 1 };
}

/** SEO Terms. */
function renderSEOTerms(sheetId, startRow, terms) {
  const out = [];
  let r = startRow;
  const labelRow = r;

  // Label row
  out.push(merge(sheetId, r, 2, 1, MAIN_COLS));
  out.push(writeCell(sheetId, r, 2, 'SEO Terms', fmt({
    bg: C.cyanLabel, fontFamily: FONT_BODY, fontSize: 14, bold: true, color: '#000020',
    hAlign: 'CENTER', vAlign: 'MIDDLE'
  })));
  out.push(setRowHeight(sheetId, r, 35));
  r++;

  // Header row
  const headers = ['Primary Variant', 'Secondary Variants', 'Typical Uses Min',
                   'Typical Uses Max', 'Current Uses', 'Add (+)/Remove (-)'];
  out.push(repeatFormat(sheetId, r, 2, 1, MAIN_COLS, { backgroundColor: hexToRgb(C.cyanHdr) }));
  const hdrFmt = fmt({
    bg: C.cyanHdr, fontFamily: FONT_BODY, fontSize: 12, bold: true, color: C.textDark,
    hAlign: 'CENTER', vAlign: 'MIDDLE', wrap: 'WRAP'
  });
  for (let i = 0; i < headers.length; i++) {
    out.push(writeCell(sheetId, r, 2 + i, headers[i], hdrFmt));
  }
  out.push(setRowHeight(sheetId, r, 35));
  r++;

  const rows = Array.isArray(terms) ? terms : [];
  if (!rows.length) {
    // Still frame the label + header with a medium outer border
    out.push(setBorder(sheetId, labelRow, 2, 2, MAIN_COLS,
      'top,bottom,left,right', 'SOLID_MEDIUM'));
    return { requests: out, nextRow: r + 1 };
  }

  const bodyStart = r;
  const cellFmtLeft = fmt({
    fontFamily: FONT_BODY, fontSize: 10, color: C.textDark, vAlign: 'MIDDLE', hAlign: 'LEFT', wrap: 'WRAP'
  });
  const cellFmtCenter = fmt({
    fontFamily: FONT_BODY, fontSize: 10, color: C.textDark, vAlign: 'MIDDLE', hAlign: 'CENTER',
    numberFormat: { type: 'NUMBER', pattern: '0' }
  });
  const adjFmt = fmt({
    fontFamily: FONT_BODY, fontSize: 10, color: C.labelGrey, vAlign: 'MIDDLE', hAlign: 'CENTER'
  });

  for (let i = 0; i < rows.length; i++, r++) {
    const t = rows[i] || {};
    out.push(writeCell(sheetId, r, 2, nl_(t.primary || ''), cellFmtLeft));
    out.push(writeCell(sheetId, r, 3, nl_(t.secondary || ''), cellFmtLeft));
    out.push(writeCell(sheetId, r, 4, Number(t.min || 0), cellFmtCenter));
    out.push(writeCell(sheetId, r, 5, Number(t.max || 0), cellFmtCenter));
    out.push(writeCell(sheetId, r, 6, Number(t.current || 0), cellFmtCenter));
    out.push(writeCell(sheetId, r, 7,
      `=if(F${r}<D${r},(D${r}-F${r}), if(F${r}>E${r},(E${r}-F${r}),"OK"))`,
      adjFmt));
    out.push(setRowHeight(sheetId, r, 28));
  }

  // Thin inner grid across header row + all body rows
  const gridTop = labelRow + 1;
  const gridHeight = r - gridTop;
  out.push(setBorder(sheetId, gridTop, 2, gridHeight, MAIN_COLS,
    'innerHorizontal,innerVertical', 'SOLID'));

  // Medium outer border around the whole table (label + header + body)
  out.push(setBorder(sheetId, labelRow, 2, r - labelRow, MAIN_COLS,
    'top,bottom,left,right', 'SOLID_MEDIUM'));

  return { requests: out, nextRow: r + 1 };
}

/** Notes — Writer / Uploader / Features (3 blocks of label + 6-row merged body). */
function renderNotes(sheetId, startRow, notes) {
  const out = [];
  let r = startRow;
  const blocks = [
    ['Writer Notes:', notes.writerNotes, true],
    ['Notes for the Uploader:', notes.notesForUploader, false],
    ['Features, Designs and/or Elements to Include:', notes.otherFeatures, true]
  ];

  blocks.forEach(([label, value, lightBg]) => {
    const bg = lightBg ? C.altLight : C.white;
    // Label row
    out.push(repeatFormat(sheetId, r, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(bg) }));
    out.push(writeCell(sheetId, r, 2, label, fmt({
      bg, fontFamily: FONT_BODY, fontSize: 11, bold: true, color: C.textDark, vAlign: 'MIDDLE'
    })));
    out.push(setRowHeight(sheetId, r, 28));
    r++;

    // Body 6 rows
    out.push(repeatFormat(sheetId, r, 1, 6, TOTAL_COLS, { backgroundColor: hexToRgb(bg) }));
    out.push(merge(sheetId, r, 2, 6, MAIN_COLS));
    out.push(writeCell(sheetId, r, 2, nl_(value || ''), fmt({
      bg, fontFamily: FONT_BODY, fontSize: 10, color: C.labelGrey,
      vAlign: 'TOP', hAlign: 'LEFT', wrap: 'WRAP'
    })));
    out.push(setBorder(sheetId, r, 2, 1, MAIN_COLS, 'top', 'SOLID'));
    out.push(setBorder(sheetId, r + 5, 2, 1, MAIN_COLS, 'bottom', 'SOLID'));
    for (let i = 0; i < 6; i++) out.push(setRowHeight(sheetId, r + i, 31));
    r += 6;
  });

  return { requests: out, nextRow: r };
}

/** QA Checklist. */
function renderQA(sheetId, startRow) {
  const out = [];
  let r = startRow;

  out.push(repeatFormat(sheetId, r, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.greyHdr) }));
  out.push(writeCell(sheetId, r, 2, 'QA Checklist', fmt({
    bg: C.greyHdr, fontFamily: FONT_BODY, fontSize: 11, bold: true, color: C.textDark, vAlign: 'MIDDLE'
  })));
  const qaHdrFmt = fmt({
    bg: C.greyHdr, fontFamily: FONT_BODY, fontSize: 11, bold: true, color: C.textDark,
    hAlign: 'CENTER', vAlign: 'MIDDLE'
  });
  out.push(writeCell(sheetId, r, 5, 'Writer Self QA', qaHdrFmt));
  out.push(writeCell(sheetId, r, 6, 'Editors/QA', qaHdrFmt));
  out.push(writeCell(sheetId, r, 7, 'Notes', qaHdrFmt));
  out.push(setBorder(sheetId, r, 1, 1, TOTAL_COLS, 'bottom', 'SOLID_MEDIUM'));
  out.push(setRowHeight(sheetId, r, 31));
  r++;

  const qaLabel = fmt({ fontFamily: FONT_BODY, fontSize: 11, color: C.csLabel, vAlign: 'MIDDLE' });
  const qaNote = fmt({ fontFamily: FONT_BODY, fontSize: 8, italic: true, color: C.labelGrey, hAlign: 'CENTER', wrap: 'WRAP', vAlign: 'MIDDLE' });

  // Row 1
  out.push(writeCell(sheetId, r, 2, 'Article Link', qaLabel));
  out.push(writeCell(sheetId, r, 3, '(LINK)', fmt({ fontFamily: FONT_BODY, fontSize: 11, color: C.textDark })));
  out.push(writeCell(sheetId, r, 4, 'Plagiarism Score\n(via Grammarly)', fmt({
    fontFamily: FONT_BODY, fontSize: 11, color: C.csLabel, hAlign: 'CENTER', vAlign: 'MIDDLE', wrap: 'WRAP'
  })));
  out.push(writeCell(sheetId, r, 7, 'Should not be more than 2%', qaNote));
  out.push(setRowHeight(sheetId, r, 31));
  r++;

  // Row 2
  out.push(writeCell(sheetId, r, 2, 'Actual Word Count', qaLabel));
  out.push(writeCell(sheetId, r, 4, 'Grammarly Score', fmt({
    fontFamily: FONT_BODY, fontSize: 11, color: C.csLabel, hAlign: 'CENTER', vAlign: 'MIDDLE', wrap: 'WRAP'
  })));
  out.push(writeLinkCell(sheetId, r, 7,
    'Should be 98% and above. Please use these Grammarly Settings.',
    'https://docs.google.com/spreadsheets/d/1KZ2oF1ERtnbHrB_idRw5Ou4AB6BOaak_aRoSLA2pMvA/edit#gid=665386058',
    qaNote));
  out.push(setRowHeight(sheetId, r, 31));
  r++;

  return { requests: out, nextRow: r + 1 };
}

/** RESEARCH DATA + Questions. */
function renderResearchAndQuestions(sheetId, startRow, questions) {
  const out = [];
  const bh = renderBlueHeader(sheetId, startRow, 'RESEARCH DATA');
  out.push(...bh.requests);
  let r = bh.nextRow;

  // Questions label row
  out.push(repeatFormat(sheetId, r, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.questBlue) }));
  out.push(merge(sheetId, r, 2, 1, MAIN_COLS));
  out.push(writeCell(sheetId, r, 2, 'Questions', fmt({
    bg: C.questBlue, fontFamily: FONT_BODY, fontSize: 11, bold: true, color: C.questText,
    hAlign: 'LEFT', vAlign: 'MIDDLE'
  })));
  out.push(setRowHeight(sheetId, r, 31));
  r++;

  const qs = Array.isArray(questions) ? questions : [];
  if (!qs.length) return { requests: out, nextRow: r + 1 };

  const bodyStart = r;
  const qFmt = fmt({
    fontFamily: FONT_BODY, fontSize: 10, color: C.csLabel,
    hAlign: 'LEFT', vAlign: 'MIDDLE', wrap: 'WRAP'
  });
  for (let i = 0; i < qs.length; i += 2, r++) {
    const left = nl_(qs[i] || '');
    const right = nl_(qs[i + 1] || '');
    out.push(merge(sheetId, r, 2, 1, 3));
    out.push(writeCell(sheetId, r, 2, left, qFmt));
    out.push(merge(sheetId, r, 5, 1, 3));
    out.push(writeCell(sheetId, r, 5, right, qFmt));
    out.push(setRowHeight(sheetId, r, 31));
  }

  out.push(...altFillRequests(sheetId, bodyStart, r - 1, 1, TOTAL_COLS, false));
  return { requests: out, nextRow: r + 1 };
}

/** Top 10 Mobile Rankings. */
function renderTop10(sheetId, startRow, t10) {
  const out = [];
  const bh = renderBlueHeaderSmall(sheetId, startRow, 'TOP 10 MOBILE RANKINGS');
  out.push(...bh.requests);
  let r = bh.nextRow;

  // Column header row
  out.push(repeatFormat(sheetId, r, 1, 1, TOTAL_COLS, { backgroundColor: hexToRgb(C.questBlue) }));
  const tHdrFmt = fmt({
    bg: C.questBlue, fontFamily: FONT_BODY, fontSize: 10, bold: true, color: C.questText,
    hAlign: 'CENTER', vAlign: 'MIDDLE'
  });
  out.push(writeCell(sheetId, r, 1, '#', tHdrFmt));
  out.push(writeCell(sheetId, r, 2, 'PAGE TITLE', tHdrFmt));
  out.push(writeCell(sheetId, r, 3, 'URL', tHdrFmt));
  out.push(merge(sheetId, r, 4, 1, 4));
  out.push(writeCell(sheetId, r, 4, 'HEADER OUTLINE', tHdrFmt));
  out.push(setRowHeight(sheetId, r, 31));
  r++;

  const rows = Array.isArray(t10) ? t10.slice(0, 50) : [];
  if (!rows.length) return { requests: out, nextRow: r + 1 };

  const bodyStart = r;
  const cellFmt = fmt({
    fontFamily: FONT_BODY, fontSize: 10, color: C.textDark, vAlign: 'MIDDLE', wrap: 'WRAP'
  });
  const outlineFmt = fmt({
    fontFamily: FONT_BODY, fontSize: 10, color: C.textDark, vAlign: 'TOP', hAlign: 'LEFT', wrap: 'WRAP'
  });
  const numFmt = fmt({
    fontFamily: FONT_BODY, fontSize: 10, color: C.textDark, vAlign: 'MIDDLE', hAlign: 'CENTER',
    numberFormat: { type: 'NUMBER', pattern: '0' }
  });

  for (let i = 0; i < rows.length; i++, r++) {
    const item = rows[i] || {};
    const rank = item.rank != null ? Number(item.rank) : (i + 1);
    out.push(writeCell(sheetId, r, 1, rank, numFmt));
    out.push(writeCell(sheetId, r, 2, nl_(item.pageTitle || ''), cellFmt));
    out.push(writeLinkCell(sheetId, r, 3, nl_(item.url || ''), item.url, cellFmt));
    out.push(merge(sheetId, r, 4, 1, 4));
    out.push(writeCell(sheetId, r, 4, bullets_(nl_(item.headerOutline || '')), outlineFmt));
    out.push(setRowHeight(sheetId, r, 60));
  }

  out.push(...altFillRequests(sheetId, bodyStart, r - 1, 1, TOTAL_COLS, false));
  out.push(setBorder(sheetId, bodyStart, 1, r - bodyStart, 7,
    'top,bottom,left,right,innerHorizontal,innerVertical', 'SOLID'));
  return { requests: out, nextRow: r + 1 };
}

/* ============================== PUBLIC ENTRY ============================== */
/**
 * Build the full requests array for a brief.
 * @param {number} sheetId - gid of target sheet
 * @param {object} job - { meta, tables, notes }
 * @returns {{ requests: Array<object> }}
 */
function buildBriefRequests(sheetId, job) {
  if (!job || !job.meta || !job.tables) {
    throw new Error('job must include {meta, tables}');
  }
  const meta = job.meta || {};
  const tables = job.tables || {};
  const notes = job.notes || {};

  const all = [];
  all.push(...resetSheetRequests(sheetId));
  all.push(...primeGridRequests(sheetId));

  let row;
  let section;

  section = renderTitleBar(sheetId, meta);     all.push(...section.requests); row = section.nextRow;
  section = renderOverview(sheetId, row, meta); all.push(...section.requests); row = section.nextRow;
  section = renderContentOutline(sheetId, row, tables.contentOutline || []);
    all.push(...section.requests); row = section.nextRow;
  section = renderTrustBenefits(sheetId, row, tables.trustElements, tables.benefitsCta, tables.painPoints);
    all.push(...section.requests); row = section.nextRow;
  section = renderResourcesTech(sheetId, row, meta, tables);
    all.push(...section.requests); row = section.nextRow;
  section = renderSEOTerms(sheetId, row, tables.seoTerms || []);
    all.push(...section.requests); row = section.nextRow;
  section = renderNotes(sheetId, row, notes);
    all.push(...section.requests); row = section.nextRow;
  section = renderQA(sheetId, row);
    all.push(...section.requests); row = section.nextRow;
  section = renderResearchAndQuestions(sheetId, row, tables.questions || []);
    all.push(...section.requests); row = section.nextRow;
  section = renderTop10(sheetId, row, tables.top10Rankings || []);
    all.push(...section.requests); row = section.nextRow;

  return { requests: all };
}

module.exports = { buildBriefRequests };
