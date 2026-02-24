/**
 * Unit tests for the quarterly amounts table fallback (detectQuarterTable).
 *
 * Usage:  node scripts/test-quarter-table.mjs
 *
 * Functions are mirrored from app.js so this runs in Node without a browser.
 */

// ── Shared helpers (mirrored from app.js) ─────────────────────
const LINE_GAP = 2;
const EUROPEAN_NUMBER_REGEX = /^-?\d{1,3}(?:\.\d{3})*(?:,\d+)?$/;
function isEuropeanNumber(s) { return EUROPEAN_NUMBER_REGEX.test(s.trim()); }

function normalizePdfText(str) {
  return String(str || '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildLineText(items) {
  if (!items.length) return '';
  const sorted = [...items].sort((a, b) => a.x - b.x);
  let text = '';
  let prev = null;
  for (const item of sorted) {
    const value = normalizePdfText(item.str);
    if (!value) continue;
    if (text) {
      const gap = prev ? item.x - (prev.x + prev.w) : LINE_GAP + 1;
      if (gap > LINE_GAP) text += ' ';
    }
    text += value;
    prev = item;
  }
  return text.replace(/\s+/g, ' ').trim();
}

// ── detectQuarterTable (mirrored from app.js) ─────────────────
const QUARTER_LABELS = ['1T', '2T', '3T', '4T', 'Total'];
function detectQuarterTable(lines) {
  const QUARTER_HEADER = /\b1T\b.*\b2T\b.*\b3T\b.*\b4T\b.*\bTotal\b/i;
  for (let i = 0; i < lines.length; i++) {
    const headerText = buildLineText(lines[i].items);
    if (!QUARTER_HEADER.test(headerText)) continue;

    const headerItems = [...lines[i].items].sort((a, b) => a.x - b.x);
    const colMidX = QUARTER_LABELS.map(label => {
      const item = headerItems.find(
        it => normalizePdfText(it.str).toUpperCase() === label.toUpperCase()
      );
      return item ? item.x + item.w / 2 : null;
    });

    if (colMidX.filter(x => x !== null).length < 2) continue;

    let j = i + 1;
    while (j < lines.length && !buildLineText(lines[j].items)) j++;
    if (j >= lines.length) continue;

    const importesText = buildLineText(lines[j].items);
    if (!/^importes\b/i.test(importesText)) continue;

    const numericItems = lines[j].items.filter(
      it => isEuropeanNumber(normalizePdfText(it.str))
    );
    if (!numericItems.length) continue;

    const values = ['', '', '', '', ''];
    for (const numItem of numericItems) {
      const numMidX = numItem.x + numItem.w / 2;
      let bestIdx = -1;
      let bestDist = Infinity;
      for (let c = 0; c < QUARTER_LABELS.length; c++) {
        if (colMidX[c] === null) continue;
        const dist = Math.abs(numMidX - colMidX[c]);
        if (dist < bestDist) { bestDist = dist; bestIdx = c; }
      }
      if (bestIdx >= 0) values[bestIdx] = normalizePdfText(numItem.str);
    }

    return { headerIdx: i, importesIdx: j, values, colMidX };
  }
  return null;
}

// ── Test helpers ──────────────────────────────────────────────
// Default column midX positions matching realistic PDF layout
const DEFAULT_COL_MIDS = [100, 160, 220, 280, 360];

// Create a single item centred at a given midX
function makeItemAt(str, midX, y = 20) {
  const w = str.length * 6;
  return { str, x: midX - w / 2, y, w, h: 12 };
}

// Build a header line with each quarter label at its column midX
function makeHeaderLine(y, labelMidXs = DEFAULT_COL_MIDS) {
  const items = QUARTER_LABELS.map((label, i) => makeItemAt(label, labelMidXs[i], y));
  return { y, items };
}

// Build an Importes line with values placed at matching column midXs.
// Pass '' for any column that should be blank (no item created).
function makeImportesLine(y, values, labelMidXs = DEFAULT_COL_MIDS) {
  const labelItem = makeItemAt('Importes', 40, y);
  const numItems = values
    .map((v, i) => v ? makeItemAt(v, labelMidXs[i], y) : null)
    .filter(Boolean);
  return { y, items: [labelItem, ...numItems] };
}

// Legacy helper for simple non-coordinate tests
function makeLine(y, text) {
  const tokens = text.split(' ');
  let x = 10;
  const items = tokens.map(token => {
    const item = { str: token, x, y, w: token.length * 6, h: 12 };
    x += token.length * 6 + 10;
    return item;
  });
  return { y, items };
}

// ── Test runner ───────────────────────────────────────────────
let passed = 0;
let failed = 0;

function assert(condition, msg) {
  if (condition) {
    console.log(`  ✅ PASS: ${msg}`);
    passed++;
  } else {
    console.log(`  ❌ FAIL: ${msg}`);
    failed++;
  }
}

// Test 1: nominal match – all 5 values present
console.log('\n── Test 1: nominal quarterly table detection (all 5 values) ──');
{
  const lines = [
    makeLine(10, 'Some other text'),
    makeHeaderLine(20),
    makeImportesLine(30, ['6.357,63', '6.402,12', '6.816,40', '6.626,67', '26.202,82']),
    makeLine(40, 'More text'),
  ];
  const result = detectQuarterTable(lines);
  assert(result !== null, 'should detect the quarterly table');
  assert(result?.headerIdx === 1, 'headerIdx should be 1');
  assert(result?.importesIdx === 2, 'importesIdx should be 2');
  assert(result?.values?.length === 5, 'should have 5 values');
  assert(result?.values?.[0] === '6.357,63', '1T value should be 6.357,63');
  assert(result?.values?.[1] === '6.402,12', '2T value should be 6.402,12');
  assert(result?.values?.[4] === '26.202,82', 'Total value should be 26.202,82');
}

// Test 2: case-insensitive "Importes"
console.log('\n── Test 2: case-insensitive Importes ──');
{
  const lines = [
    makeHeaderLine(10),
    makeImportesLine(20, ['1.000,00', '2.000,00', '3.000,00', '4.000,00', '10.000,00']),
  ];
  // Override label item text to lowercase
  lines[1].items[0].str = 'importes';
  const result = detectQuarterTable(lines);
  assert(result !== null, 'should detect with lowercase importes');
}

// Test 3: missing one quarter value (blank 2T) – must NOT shift other values
console.log('\n── Test 3: missing 2T – values must not shift left ──');
{
  const lines = [
    makeHeaderLine(10),
    makeImportesLine(20, ['6.357,63', '', '6.816,40', '6.626,67', '26.202,82']),
  ];
  const result = detectQuarterTable(lines);
  assert(result !== null, 'should detect even with blank 2T');
  assert(result?.values?.[0] === '6.357,63', '1T should be 6.357,63');
  assert(result?.values?.[1] === '', '2T should be blank');
  assert(result?.values?.[2] === '6.816,40', '3T should be 6.816,40 (not shifted)');
  assert(result?.values?.[3] === '6.626,67', '4T should be 6.626,67');
  assert(result?.values?.[4] === '26.202,82','Total should be 26.202,82');
}

// Test 4: no match – missing header
console.log('\n── Test 4: no quarterly header present ──');
{
  const lines = [
    makeLine(10, 'January February March Total'),
    makeImportesLine(20, ['1.000,00', '2.000,00', '3.000,00', '4.000,00', '10.000,00']),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect without 1T/2T/3T/4T header');
}

// Test 5: no match – Importes line does not follow header
console.log('\n── Test 5: Importes line missing after header ──');
{
  const lines = [
    makeHeaderLine(10),
    makeLine(20, 'SomethingElse 1.000,00 2.000,00 3.000,00 4.000,00 10.000,00'),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect when next line does not start with Importes');
}

// Test 6: no match – no numeric values in Importes line
console.log('\n── Test 6: non-numeric values in Importes line ──');
{
  const lines = [
    makeHeaderLine(10),
    makeLine(20, 'Importes abc def ghi jkl mno'),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect with non-numeric values');
}

// Test 7: empty line between header and Importes is skipped
console.log('\n── Test 7: empty line between header and Importes ──');
{
  const lines = [
    makeHeaderLine(10),
    { y: 20, items: [] },  // empty line
    makeImportesLine(30, ['6.357,63', '6.402,12', '6.816,40', '6.626,67', '26.202,82']),
  ];
  const result = detectQuarterTable(lines);
  assert(result !== null, 'should detect even with empty line between header and Importes');
  assert(result?.importesIdx === 2, 'importesIdx should be 2 (after the empty line)');
}

// Test 8: only 1 valid column midpoint – should not detect (< 2 threshold)
console.log('\n── Test 8: only 1 header label item findable – should not detect ──');
{
  // Header line text passes the regex but only "1T" is an exact-match item
  const items = [
    makeItemAt('1T 2T 3T 4T Total', 150, 10),  // all labels in ONE item
  ];
  const lines = [
    { y: 10, items },
    makeImportesLine(20, ['6.357,63', '6.402,12', '6.816,40', '6.626,67', '26.202,82']),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect when < 2 column midpoints are findable');
}

// ── Summary ───────────────────────────────────────────────────
console.log(`\n══ Results: ${passed} passed, ${failed} failed ══`);
process.exit(failed > 0 ? 1 : 0);
