/**
 * Google Apps Script Web App — Content Brief Renderer (v2)
 * Matches the "roulette online" tab layout from
 * NewExample_betonline.ag - Content Briefs _ Internal.xlsx
 *
 * EXPECTED POST BODY (array or single object):
 * {
 *   spreadsheetId: string,
 *   sheetId?: number|string,
 *   sheetName: string,
 *   meta: {
 *     homeUrl, mainKeyword, keywordVolume,
 *     recommendedUrl, minWordCount, existingOrNew, pageType,
 *     targetGeo, clearScopeLink
 *   },
 *   tables: {
 *     contentOutline: Array<{
 *       Section, Heading, "Writer Instructions", Type,
 *       "Capsule?", "Word Target", "Required Elements",
 *       "Entities / Terms", "WRITER ✓", "EDITOR ✓"
 *     }>,
 *     seoTerms: Array<{ primary, secondary, min, max, current, adj }>,
 *     top10Rankings: Array<{ rank, pageTitle, url, headerOutline }>,
 *     questions: string[],
 *     trustElements: string,
 *     benefitsCta: string,
 *     painPoints?: string,
 *     clearscopeWriterGrade?: string,
 *     clearscopeEditorGrade?: string,
 *     clearscopeWriterReadability?: string,
 *     clearscopeEditorReadability?: string
 *   },
 *   notes: {
 *     writerNotes, notesForUploader, otherFeatures
 *   }
 * }
 */

/* =================== CONSTANTS =================== */
const C = {
  blue:       '#182F7A',
  altLight:   '#f8fafc',
  white:      '#ffffff',
  labelGrey:  '#334155',
  textDark:   '#0F172A',
  greyHdr:    '#e2e8f0',
  trust:      '#6aa84f',
  benefits:   '#3d85c6',
  pain:       '#cc0000',
  csRed:      '#f4cccc',
  csLabel:    '#64748b',
  cyanLabel:  '#04dbf5',
  cyanHdr:    '#defbff',
  questBlue:  '#e0f2fe',
  questText:  '#0c4a6e',
  link:       '#0000ff'
};
const TOTAL_COLS = 11;       // A..K
const OUTLINE_COLS = 10;     // B..K
const MAIN_COLS = 6;         // B..G
const FONT_SECTION = 'Lexend';
const FONT_BODY = 'Arial';

/* =================== UTILITIES =================== */
function nl_(v) {
  if (v == null) return v;
  if (typeof v !== 'string') return v;
  return v.replace(/\\n\\n/g, '\n\n').replace(/\\n/g, '\n');
}
function bullets_(v) {
  if (v == null || typeof v !== 'string') return v;
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
function letter_(n) {
  let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function a1_(r, c, nr, nc) {
  return `${letter_(c)}${r}:${letter_(c + nc - 1)}${r + nr - 1}`;
}
function safeUnmerge_(sheet, a1) {
  const r = sheet.getRange(a1);
  const m = r.getMergedRanges();
  if (m && m.length) m.forEach(x => x.breakApart());
  r.breakApart();
}
function setRowH_(sheet, row, h) { sheet.setRowHeight(row, h); }
function clearAllBorders_(sheet) {
  sheet.getDataRange().setBorder(false, false, false, false, false, false, null, null);
}
function hyperlinkify_(rng) {
  const vals = rng.getValues();
  const rich = vals.map(row => row.map(cell => {
    const txt = (cell == null) ? '' : String(cell);
    const b = SpreadsheetApp.newRichTextValue().setText(txt);
    if (/^https?:\/\/\S+$/i.test(txt)) return b.setLinkUrl(txt).build();
    return b.build();
  }));
  rng.setRichTextValues(rich);
}
function setLink_(rng, text, url) {
  rng.setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText(text || '').setLinkUrl(url || null).build()
  );
}
function placeLogo_(sheet, url, anchorRow, anchorCol, rowsSpan, colsSpan) {
  const r = sheet.getRange(anchorRow, anchorCol, rowsSpan, colsSpan);
  const w = r.getWidth(), h = r.getHeight();
  try {
    const blob = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getBlob();
    sheet.getImages().forEach(img => {
      const pos = img.getAnchorCell && img.getAnchorCell();
      if (pos && pos.getRow() === anchorRow && pos.getColumn() === anchorCol) img.remove();
    });
    const img = sheet.insertImage(blob, anchorCol, anchorRow);
    const nW = img.getWidth(), nH = img.getHeight();
    const scale = Math.min(w / nW, h / nH);
    const newW = Math.round(nW * scale), newH = Math.round(nH * scale);
    img.setWidth(newW).setHeight(newH);
    if (img.setAnchorCell) {
      img.setAnchorCell(r);
      if (img.setAnchorCellXOffset) img.setAnchorCellXOffset(Math.round((w - newW) / 2));
      if (img.setAnchorCellYOffset) img.setAnchorCellYOffset(Math.round((h - newH) / 2));
    }
  } catch (_) {}
}
function fillAlt_(sheet, r1, r2, c1, c2, startLight) {
  if (!r1 || !r2 || r2 < r1) return;
  for (let r = r1; r <= r2; r++) {
    const isLight = ((r - r1) % 2 === 0) ? startLight : !startLight;
    sheet.getRange(r, c1, 1, c2 - c1 + 1).setBackground(isLight ? C.altLight : C.white);
  }
}
function deleteTrailingRows_(sheet) {
  const last = sheet.getLastRow(), max = sheet.getMaxRows();
  if (max > last && last > 0) sheet.deleteRows(last + 1, max - last);
}

/* =================== GRID PRIME =================== */
function primeGrid_(sheet) {
  // Ensure A..K
  if (sheet.getMaxColumns() < TOTAL_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), TOTAL_COLS - sheet.getMaxColumns());
  } else if (sheet.getMaxColumns() > TOTAL_COLS) {
    sheet.deleteColumns(TOTAL_COLS + 1, sheet.getMaxColumns() - TOTAL_COLS);
  }
  // Column widths (matched to xlsx)
  sheet.setColumnWidth(1, 33);    // A gutter
  sheet.setColumnWidth(2, 184);   // B
  sheet.setColumnWidth(3, 70);    // C
  sheet.setColumnWidth(4, 370);   // D (Writer Instructions)
  sheet.setColumnWidth(5, 184);   // E
  sheet.setColumnWidth(6, 70);    // F
  sheet.setColumnWidth(7, 70);    // G
  sheet.setColumnWidth(8, 218);   // H (Required Elements)
  sheet.setColumnWidth(9, 184);   // I (Entities)
  sheet.setColumnWidth(10, 80);   // J (Writer ✓)
  sheet.setColumnWidth(11, 80);   // K (Editor ✓)

  sheet.getDataRange().setWrap(true);
  sheet.setFrozenRows(2);
}

/* =================== RENDERERS =================== */

/** Title bar rows 1-2 — A1:K2 blue, B1:G2 merged title, logo H1:K2. */
function renderTitleBar_(sheet, meta) {
  sheet.clear();
  sheet.getImages().forEach(img => img.remove());
  primeGrid_(sheet);
  clearAllBorders_(sheet);

  sheet.getRange(1, 1, 2, TOTAL_COLS).setBackground(C.blue);
  safeUnmerge_(sheet, 'B1:G2');
  const title = sheet.getRange('B1:G2').merge();
  const home = (meta && meta.homeUrl) ? String(meta.homeUrl).replace(/^https?:\/\//i, '') : '';
  title.setValue('Content Brief: ' + home)
       .setFontFamily(FONT_SECTION).setFontSize(24)
       .setFontColor(C.white).setFontWeight('bold')
       .setHorizontalAlignment('left').setVerticalAlignment('middle');

  safeUnmerge_(sheet, 'H1:K2');
  sheet.getRange('H1:K2').merge();
  placeLogo_(sheet,
    'https://cdn-amehi.nitrocdn.com/ldNPGLQtVWaqliEfWebnqecfYajgRCdk/assets/images/optimized/rev-7ef5dc9/www.digitalspotlight.com/wp-content/uploads/2017/09/logo-new.png',
    1, 8, 2, 4);
  setRowH_(sheet, 1, 33); setRowH_(sheet, 2, 33);
  return 3;
}

/** Overview (3 rows): labels in B/E, values in C and F. f8fafc bg on rows 3 & 5. */
function renderOverview_(sheet, startRow, meta) {
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
  for (let i = 0; i < 3; i++) {
    sheet.getRange(r + i, 2).setValue(labels[i][0]);
    sheet.getRange(r + i, 5).setValue(labels[i][1]);
    sheet.getRange(r + i, 3).setValue(nl_(values[i][0]));
    sheet.getRange(r + i, 6).setValue(nl_(values[i][1]));
  }
  // Label styles
  const labelRng = sheet.getRangeList([
    `${letter_(2)}${r}:${letter_(2)}${r+2}`,
    `${letter_(5)}${r}:${letter_(5)}${r+2}`
  ]);
  labelRng.setFontFamily(FONT_BODY).setFontSize(12).setFontWeight('bold').setFontColor(C.labelGrey);

  // Value styles
  sheet.getRange(r, 3, 3, 1)
    .setFontFamily(FONT_BODY).setFontSize(12).setFontColor(C.textDark).setVerticalAlignment('middle');
  sheet.getRange(r, 6, 3, 1)
    .setFontFamily(FONT_BODY).setFontSize(12).setFontColor(C.textDark).setVerticalAlignment('middle');

  // Alternating bg rows: r=light, r+1=white, r+2=light (across full A..K)
  sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(C.altLight);
  sheet.getRange(r + 1, 1, 1, TOTAL_COLS).setBackground(C.white);
  sheet.getRange(r + 2, 1, 1, TOTAL_COLS).setBackground(C.altLight);

  // Recommended URL hyperlink
  setLink_(sheet.getRange(r + 1, 3), meta.recommendedUrl || '', meta.recommendedUrl || null);

  for (let i = 0; i < 3; i++) setRowH_(sheet, r + i, 49);
  return r + 3 + 1; // spacer
}

/** Big blue section header (A:K with B:G merged). */
function renderBlueHeader_(sheet, row, text, height) {
  sheet.getRange(row, 1, 2, TOTAL_COLS).setBackground(C.blue);
  safeUnmerge_(sheet, a1_(row, 2, 2, MAIN_COLS));
  sheet.getRange(row, 2, 2, MAIN_COLS).merge()
    .setValue(text || '')
    .setFontFamily(FONT_SECTION).setFontSize(24).setFontWeight('bold').setFontColor(C.white)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  setRowH_(sheet, row, height || 35);
  setRowH_(sheet, row + 1, height || 35);
  return row + 2;
}

/** Single-row narrower section header (used by TOP 10). */
function renderBlueHeaderSmall_(sheet, row, text) {
  sheet.getRange(row, 1, 1, TOTAL_COLS).setBackground(C.blue);
  safeUnmerge_(sheet, a1_(row, 2, 1, MAIN_COLS));
  sheet.getRange(row, 2, 1, MAIN_COLS).merge()
    .setValue(text || '')
    .setFontFamily(FONT_BODY).setFontSize(11).setFontWeight('bold').setFontColor(C.white)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  setRowH_(sheet, row, 31);
  return row + 1;
}

/** Content Outline — 10 cols (B..K). */
function renderContentOutline_(sheet, startRow, items) {
  let r = renderBlueHeader_(sheet, startRow, 'Content Outline', 35);

  // Header row
  const hdr = r;
  const headers = ['Section', 'Heading', 'Writer Instructions', 'Type', 'Capsule?',
                   'Word Target', 'Required Elements', 'Entities / Terms', 'WRITER ✓', 'EDITOR ✓'];
  sheet.getRange(hdr, 1, 1, TOTAL_COLS).setBackground(C.greyHdr);
  for (let i = 0; i < headers.length; i++) sheet.getRange(hdr, 2 + i).setValue(headers[i]);
  sheet.getRange(hdr, 2, 1, OUTLINE_COLS)
    .setFontFamily(FONT_BODY).setFontSize(11).setFontWeight('bold').setFontColor(C.labelGrey)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.getRange(hdr, 1, 1, TOTAL_COLS).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  setRowH_(sheet, hdr, 35);
  r++;

  const rows = Array.isArray(items) ? items : [];
  if (!rows.length) return r + 1;

  const startBody = r;
  const get = (it, ...keys) => { for (const k of keys) if (it[k] != null) return it[k]; return ''; };

  for (let i = 0; i < rows.length; i++, r++) {
    const it = rows[i] || {};
    const section  = nl_(get(it, 'Section', 'section'));
    const heading  = nl_(get(it, 'Heading', 'heading'));
    const writer   = bullets_(nl_(get(it, 'Writer Instructions', 'writerInstructions', 'reqs')));
    const type     = nl_(get(it, 'Type', 'type'));
    const capsule  = nl_(get(it, 'Capsule?', 'capsule'));
    const wordT    = nl_(get(it, 'Word Target', 'wordTarget'));
    const reqElem  = bullets_(nl_(get(it, 'Required Elements', 'requiredElements')));
    const entities = nl_(get(it, 'Entities / Terms', 'entities'));
    const w        = toBool_(get(it, 'WRITER ✓', 'writer'));
    const e        = toBool_(get(it, 'EDITOR ✓', 'editor'));

    sheet.getRange(r, 2).setValue(section).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(r, 3).setValue(heading).setVerticalAlignment('middle');
    sheet.getRange(r, 4).setValue(writer).setVerticalAlignment('middle');
    sheet.getRange(r, 5).setValue(type).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(r, 6).setValue(capsule).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(r, 7).setValue(wordT).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(r, 8).setValue(reqElem).setVerticalAlignment('middle');
    sheet.getRange(r, 9).setValue(entities).setVerticalAlignment('middle');

    const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(r, 10).setDataValidation(cb).setValue(w);
    sheet.getRange(r, 11).setDataValidation(cb).setValue(e);
    sheet.getRange(r, 10, 1, 2).setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.getRange(r, 2, 1, OUTLINE_COLS)
      .setFontFamily(FONT_BODY).setFontSize(12).setFontColor(C.textDark).setWrap(true);

    sheet.autoResizeRows(r, 1);
  }
  fillAlt_(sheet, startBody, r - 1, 1, TOTAL_COLS, true);
  return r + 1;
}

/** Trust / Benefits / Pain Points — labels in row, body in 6 merged rows below. */
function renderTrustBenefits_(sheet, startRow, trust, benefits, pain) {
  let r = startRow;

  safeUnmerge_(sheet, a1_(r, 2, 1, 2)); sheet.getRange(r, 2, 1, 2).merge().setValue('TRUST ELEMENTS')
    .setBackground(C.trust).setFontColor(C.white).setFontFamily(FONT_BODY).setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  safeUnmerge_(sheet, a1_(r, 4, 1, 2)); sheet.getRange(r, 4, 1, 2).merge().setValue('Benefits, Offers & CTAs')
    .setBackground(C.benefits).setFontColor(C.white).setFontFamily(FONT_BODY).setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  safeUnmerge_(sheet, a1_(r, 6, 1, 2)); sheet.getRange(r, 6, 1, 2).merge().setValue('PAIN POINTS')
    .setBackground(C.pain).setFontColor(C.white).setFontFamily(FONT_BODY).setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  setRowH_(sheet, r, 35);
  r++;

  // 6-row merged body blocks
  const blocks = [
    [2, bullets_(nl_(trust || ''))],
    [4, bullets_(nl_(benefits || ''))],
    [6, bullets_(nl_(pain || ''))]
  ];
  blocks.forEach(([col, val]) => {
    safeUnmerge_(sheet, a1_(r, col, 6, 2));
    sheet.getRange(r, col, 6, 2).merge().setValue(val)
      .setBackground(C.white).setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.labelGrey)
      .setVerticalAlignment('top').setWrap(true);
  });
  for (let i = 0; i < 6; i++) setRowH_(sheet, r + i, 35);
  return r + 6 + 1;
}

/** Resources & Technical Details (incl. Clearscope row + Target GEO). */
function renderResourcesTech_(sheet, startRow, meta, tables) {
  let r = renderBlueHeader_(sheet, startRow, 'RESOURCES & TECHNICAL DETAILS', 35);

  // SEO Terms Details label
  safeUnmerge_(sheet, a1_(r, 2, 1, MAIN_COLS));
  sheet.getRange(r, 2, 1, MAIN_COLS).merge()
    .setValue('SEO Terms Details')
    .setFontFamily(FONT_BODY).setFontSize(11).setFontWeight('bold').setFontColor(C.textDark)
    .setVerticalAlignment('bottom').setHorizontalAlignment('left');
  setRowH_(sheet, r, 31);
  r++;

  // Clearscope strip — 3 rows on red bg
  const csRows = 3;
  sheet.getRange(r, 1, csRows, TOTAL_COLS).setBackground(C.csRed);

  // Row 1: Clearscope link + writer grade + editor grade
  sheet.getRange(r, 2).setValue('Clearscope Report Link:')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.csLabel).setVerticalAlignment('bottom');
  setLink_(sheet.getRange(r, 3), meta.clearScopeLink || '', meta.clearScopeLink || null);
  sheet.getRange(r, 3).setFontColor(C.link);
  sheet.getRange(r, 4).setValue('Clearscope Grade (writer):')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.csLabel).setVerticalAlignment('bottom');
  sheet.getRange(r, 5).setValue(tables && tables.clearscopeWriterGrade || '')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.textDark).setVerticalAlignment('bottom');
  sheet.getRange(r, 6).setValue('Clearscope Grade (editor/QA):')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.csLabel).setVerticalAlignment('bottom');
  sheet.getRange(r, 7).setValue(tables && tables.clearscopeEditorGrade || '')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.textDark).setVerticalAlignment('bottom');
  setRowH_(sheet, r, 27);
  r++;

  // Row 2: readability (writer/editor)
  sheet.getRange(r, 4).setValue('Clearscope Readability Grade (writer):')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.csLabel).setVerticalAlignment('bottom').setWrap(true);
  sheet.getRange(r, 5).setValue(tables && tables.clearscopeWriterReadability || '')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.textDark).setVerticalAlignment('bottom');
  sheet.getRange(r, 6).setValue('Clearscope Readability Grade (editor/QA):')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.csLabel).setVerticalAlignment('bottom').setWrap(true);
  sheet.getRange(r, 7).setValue(tables && tables.clearscopeEditorReadability || '')
    .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.textDark).setVerticalAlignment('bottom');
  setRowH_(sheet, r, 31);
  r++;

  // Row 3: spacer red row
  setRowH_(sheet, r, 22);
  r++;

  // Spacer + Target GEO
  r++;
  sheet.getRange(r, 2).setValue('TARGET GEO:')
    .setFontFamily(FONT_BODY).setFontSize(11).setFontColor(C.csLabel);
  sheet.getRange(r, 3).setValue(nl_(meta.targetGeo || ''))
    .setFontFamily(FONT_BODY).setFontSize(11).setFontColor(C.labelGrey);
  r++;

  return r + 1;
}

/** SEO Terms — cyan label, light cyan header, body with adj formula. */
function renderSEOTerms_(sheet, startRow, terms) {
  let r = startRow;

  // Label row B:G cyan, merged
  safeUnmerge_(sheet, a1_(r, 2, 1, MAIN_COLS));
  sheet.getRange(r, 2, 1, MAIN_COLS).merge()
    .setValue('SEO Terms')
    .setBackground(C.cyanLabel).setFontFamily(FONT_BODY).setFontSize(14).setFontWeight('bold').setFontColor('#000020')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  setRowH_(sheet, r, 35);
  r++;

  // Header row
  const headers = ['Primary Variant', 'Secondary Variants', 'Typical Uses Min', 'Typical Uses Max', 'Current Uses', 'Add (+)/Remove (-)'];
  for (let i = 0; i < headers.length; i++) sheet.getRange(r, 2 + i).setValue(headers[i]);
  sheet.getRange(r, 2, 1, MAIN_COLS)
    .setBackground(C.cyanHdr).setFontFamily(FONT_BODY).setFontSize(12).setFontWeight('bold').setFontColor(C.textDark)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  setRowH_(sheet, r, 35);
  r++;

  const rows = Array.isArray(terms) ? terms : [];
  if (!rows.length) return r + 1;

  const startBody = r;
  for (let i = 0; i < rows.length; i++, r++) {
    const t = rows[i] || {};
    sheet.getRange(r, 2).setValue(nl_(t.primary || '')).setHorizontalAlignment('left');
    sheet.getRange(r, 3).setValue(nl_(t.secondary || '')).setHorizontalAlignment('left');
    sheet.getRange(r, 4).setValue(Number(t.min || 0));
    sheet.getRange(r, 5).setValue(Number(t.max || 0));
    sheet.getRange(r, 6).setValue(Number(t.current || 0));
    sheet.getRange(r, 7).setFormula(`=if(F${r}<D${r},(D${r}-F${r}), if(F${r}>E${r},(E${r}-F${r}),"OK"))`);
    sheet.getRange(r, 2, 1, MAIN_COLS)
      .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.textDark).setVerticalAlignment('middle').setWrap(true);
    sheet.getRange(r, 4, 1, 3).setHorizontalAlignment('center');
    sheet.getRange(r, 7).setFontColor(C.labelGrey);
    setRowH_(sheet, r, 28);
  }
  sheet.getRange(startBody, 4, r - startBody, 3).setNumberFormat('0');
  return r + 1;
}

/** Notes — Writer / Uploader / Features. */
function renderNotes_(sheet, startRow, notes) {
  let r = startRow;
  const block = (label, value, lightBg) => {
    sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(lightBg ? C.altLight : C.white);
    sheet.getRange(r, 2).setValue(label)
      .setFontFamily(FONT_BODY).setFontSize(11).setFontWeight('bold').setFontColor(C.textDark);
    setRowH_(sheet, r, 28);
    r++;
    safeUnmerge_(sheet, a1_(r, 2, 6, MAIN_COLS));
    sheet.getRange(r, 1, 6, TOTAL_COLS).setBackground(lightBg ? C.altLight : C.white);
    sheet.getRange(r, 2, 6, MAIN_COLS).merge()
      .setValue(nl_(value || ''))
      .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.labelGrey)
      .setVerticalAlignment('top').setWrap(true).setHorizontalAlignment('left');
    sheet.getRange(r, 2, 1, MAIN_COLS).setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r + 5, 2, 1, MAIN_COLS).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    for (let i = 0; i < 6; i++) setRowH_(sheet, r + i, 31);
    r += 6;
  };
  block('Writer Notes:', notes.writerNotes, true);
  block('Notes for the Uploader:', notes.notesForUploader, false);
  block('Features, Designs and/or Elements to Include:', notes.otherFeatures, true);
  return r;
}

/** QA Checklist — 1 header + 2 prefilled rows. */
function renderQA_(sheet, startRow) {
  let r = startRow;
  sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(C.greyHdr);
  sheet.getRange(r, 2).setValue('QA Checklist')
    .setFontFamily(FONT_BODY).setFontSize(11).setFontWeight('bold').setFontColor(C.textDark);
  sheet.getRange(r, 5).setValue('Writer Self QA');
  sheet.getRange(r, 6).setValue('Editors/QA');
  sheet.getRange(r, 7).setValue('Notes');
  sheet.getRange(r, 5, 1, 3)
    .setFontFamily(FONT_BODY).setFontSize(11).setFontWeight('bold').setFontColor(C.textDark)
    .setHorizontalAlignment('center');
  sheet.getRange(r, 1, 1, TOTAL_COLS).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  setRowH_(sheet, r, 31);
  r++;

  sheet.getRange(r, 2).setValue('Article Link').setFontColor(C.csLabel);
  sheet.getRange(r, 3).setValue('(LINK)');
  sheet.getRange(r, 4).setValue('Plagiarism Score\n(via Grammarly)').setFontColor(C.csLabel).setHorizontalAlignment('center').setWrap(true);
  sheet.getRange(r, 7).setValue('Should not be more than 2%').setFontSize(8).setFontStyle('italic').setFontColor(C.labelGrey).setHorizontalAlignment('center');
  setRowH_(sheet, r, 31);
  r++;

  sheet.getRange(r, 2).setValue('Actual Word Count').setFontColor(C.csLabel);
  sheet.getRange(r, 4).setValue('Grammarly Score').setFontColor(C.csLabel).setHorizontalAlignment('center').setWrap(true);
  setLink_(sheet.getRange(r, 7), 'Should be 98% and above. Please use these Grammarly Settings.', 'https://docs.google.com/spreadsheets/d/1KZ2oF1ERtnbHrB_idRw5Ou4AB6BOaak_aRoSLA2pMvA/edit#gid=665386058');
  sheet.getRange(r, 7).setFontSize(8).setFontStyle('italic').setHorizontalAlignment('center').setWrap(true);
  setRowH_(sheet, r, 31);
  r++;

  return r + 1;
}

/** RESEARCH DATA + Questions (immediately after, no spacer). */
function renderResearchAndQuestions_(sheet, startRow, questions) {
  let r = renderBlueHeader_(sheet, startRow, 'RESEARCH DATA', 35);

  // Questions label row (light blue)
  sheet.getRange(r, 1, 1, TOTAL_COLS).setBackground(C.questBlue);
  safeUnmerge_(sheet, a1_(r, 2, 1, MAIN_COLS));
  sheet.getRange(r, 2, 1, MAIN_COLS).merge()
    .setValue('Questions')
    .setFontFamily(FONT_BODY).setFontSize(11).setFontWeight('bold').setFontColor(C.questText)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  setRowH_(sheet, r, 31);
  r++;

  const qs = Array.isArray(questions) ? questions : [];
  if (!qs.length) return r + 1;

  const startBody = r;
  for (let i = 0; i < qs.length; i += 2, r++) {
    const left = nl_(qs[i] || '');
    const right = nl_(qs[i + 1] || '');
    safeUnmerge_(sheet, a1_(r, 2, 1, 3));
    sheet.getRange(r, 2, 1, 3).merge().setValue(left)
      .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.csLabel)
      .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('left');
    safeUnmerge_(sheet, a1_(r, 5, 1, 3));
    sheet.getRange(r, 5, 1, 3).merge().setValue(right)
      .setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.csLabel)
      .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('left');
    setRowH_(sheet, r, 31);
  }
  fillAlt_(sheet, startBody, r - 1, 1, TOTAL_COLS, false);
  return r + 1;
}

/** Top 10 Mobile Rankings. */
function renderTop10_(sheet, startRow, t10) {
  let r = renderBlueHeaderSmall_(sheet, startRow, 'TOP 10 MOBILE RANKINGS');

  // Column header row (light blue)
  const hdr = r;
  sheet.getRange(hdr, 1, 1, TOTAL_COLS).setBackground(C.questBlue);
  sheet.getRange(hdr, 1).setValue('#');
  sheet.getRange(hdr, 2).setValue('PAGE TITLE');
  sheet.getRange(hdr, 3).setValue('URL');
  safeUnmerge_(sheet, a1_(hdr, 4, 1, 4));
  sheet.getRange(hdr, 4, 1, 4).merge().setValue('HEADER OUTLINE');
  sheet.getRange(hdr, 1, 1, 7)
    .setFontFamily(FONT_BODY).setFontSize(10).setFontWeight('bold').setFontColor(C.questText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  setRowH_(sheet, hdr, 31);
  r++;

  const rows = Array.isArray(t10) ? t10.slice(0, 50) : [];
  if (!rows.length) return r + 1;

  const startBody = r;
  for (let i = 0; i < rows.length; i++, r++) {
    const item = rows[i] || {};
    const rank = item.rank != null ? item.rank : (i + 1);
    sheet.getRange(r, 1).setValue(rank).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(r, 2).setValue(nl_(item.pageTitle || '')).setVerticalAlignment('middle');
    setLink_(sheet.getRange(r, 3), nl_(item.url || ''), item.url || null);
    sheet.getRange(r, 3).setVerticalAlignment('middle');
    safeUnmerge_(sheet, a1_(r, 4, 1, 4));
    sheet.getRange(r, 4, 1, 4).merge()
      .setValue(bullets_(nl_(item.headerOutline || '')))
      .setWrap(true).setVerticalAlignment('top').setHorizontalAlignment('left');
    sheet.getRange(r, 1, 1, 7).setFontFamily(FONT_BODY).setFontSize(10).setFontColor(C.textDark);
    sheet.autoResizeRows(r, 1);
  }
  fillAlt_(sheet, startBody, r - 1, 1, TOTAL_COLS, false);
  sheet.getRange(startBody, 1, r - startBody, 7)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  return r + 1;
}

/* =================== MAIN =================== */
function validateJob_(job) {
  if (!job || typeof job !== 'object') throw new Error('Missing job payload.');
  if (!job.spreadsheetId) throw new Error('spreadsheetId is required.');
  if (!job.sheetName) throw new Error('sheetName is required.');
  if (!job.meta || !job.tables) throw new Error('meta and tables are required.');
}
function ensureSheet_(ss, sheetName, sheetId) {
  if (sheetId != null && sheetId !== '') {
    const byId = ss.getSheets().find(s => String(s.getSheetId()) === String(sheetId));
    if (byId) return byId;
  }
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  return sh;
}

function renderBrief_(job) {
  validateJob_(job);
  const ss = SpreadsheetApp.openById(job.spreadsheetId);
  const sheet = ensureSheet_(ss, job.sheetName, job.sheetId);
  const meta = job.meta || {};
  const tables = job.tables || {};
  const notes = job.notes || {};

  let row = renderTitleBar_(sheet, meta);
  row = renderOverview_(sheet, row, meta);
  row = renderContentOutline_(sheet, row, tables.contentOutline || []);
  row = renderTrustBenefits_(sheet, row, tables.trustElements, tables.benefitsCta, tables.painPoints);
  row = renderResourcesTech_(sheet, row, meta, tables);
  row = renderSEOTerms_(sheet, row, tables.seoTerms || []);
  row = renderNotes_(sheet, row, notes);
  row = renderQA_(sheet, row);
  row = renderResearchAndQuestions_(sheet, row, tables.questions || []);
  row = renderTop10_(sheet, row, tables.top10Rankings || []);

  deleteTrailingRows_(sheet);
  sheet.setFrozenRows(2);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'No POST body received' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    let payload;
    try { payload = JSON.parse(e.postData.contents); }
    catch (err) { payload = e.parameter || {}; }

    const jobs = Array.isArray(payload) ? payload : [payload];
    jobs.forEach(job => {
      if (job && job.meta && job.tables) renderBrief_(job);
      else throw new Error('Unsupported payload shape. Provide {meta, tables, notes}.');
    });
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('doPost error:', err);
    return ContentService.createTextOutput(JSON.stringify({ error: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
