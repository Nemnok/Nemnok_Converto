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
function detectQuarterTable(lines) {
  const QUARTER_HEADER = /\b1T\b.*\b2T\b.*\b3T\b.*\b4T\b.*\bTotal\b/i;
  for (let i = 0; i < lines.length; i++) {
    const headerText = buildLineText(lines[i].items);
    if (!QUARTER_HEADER.test(headerText)) continue;

    let j = i + 1;
    while (j < lines.length && !buildLineText(lines[j].items)) j++;
    if (j >= lines.length) continue;

    const importesText = buildLineText(lines[j].items);
    if (!/^importes\b/i.test(importesText)) continue;

    const tokens = importesText.replace(/^importes\s*/i, '').trim().split(/\s+/);
    if (tokens.length !== 5) continue;
    if (!tokens.every(t => isEuropeanNumber(t))) continue;

    return { headerIdx: i, importesIdx: j, values: tokens };
  }
  return null;
}

// ── Helper: build synthetic line items ────────────────────────
function makeLine(y, text) {
  // Produce one item per space-separated token at increasing x positions
  const tokens = text.split(' ');
  let x = 10;
  const items = tokens.map(token => {
    const item = { str: token, x, y, w: token.length * 6, h: 12 };
    x += token.length * 6 + 10;
    return item;
  });
  return { y, items };
}

// ── Test cases ────────────────────────────────────────────────
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

// Test 1: nominal match
console.log('\n── Test 1: nominal quarterly table detection ──');
{
  const lines = [
    makeLine(10, 'Some other text'),
    makeLine(20, '1T 2T 3T 4T Total'),
    makeLine(30, 'Importes 6.357,63 6.402,12 6.816,40 6.626,67 26.202,82'),
    makeLine(40, 'More text'),
  ];
  const result = detectQuarterTable(lines);
  assert(result !== null, 'should detect the quarterly table');
  assert(result?.headerIdx === 1, 'headerIdx should be 1');
  assert(result?.importesIdx === 2, 'importesIdx should be 2');
  assert(result?.values?.length === 5, 'should have 5 values');
  assert(result?.values?.[0] === '6.357,63', 'first value should be 6.357,63');
  assert(result?.values?.[4] === '26.202,82', 'last value should be 26.202,82');
}

// Test 2: case-insensitive "Importes"
console.log('\n── Test 2: case-insensitive Importes ──');
{
  const lines = [
    makeLine(10, '1T 2T 3T 4T Total'),
    makeLine(20, 'importes 1.000,00 2.000,00 3.000,00 4.000,00 10.000,00'),
  ];
  const result = detectQuarterTable(lines);
  assert(result !== null, 'should detect with lowercase importes');
}

// Test 3: no match – wrong number of values
console.log('\n── Test 3: wrong number of values (4 instead of 5) ──');
{
  const lines = [
    makeLine(10, '1T 2T 3T 4T Total'),
    makeLine(20, 'Importes 1.000,00 2.000,00 3.000,00 4.000,00'),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect with only 4 values');
}

// Test 4: no match – missing header
console.log('\n── Test 4: no quarterly header present ──');
{
  const lines = [
    makeLine(10, 'January February March Total'),
    makeLine(20, 'Importes 1.000,00 2.000,00 3.000,00 4.000,00 10.000,00'),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect without 1T/2T/3T/4T header');
}

// Test 5: no match – Importes line does not follow header
console.log('\n── Test 5: Importes line missing after header ──');
{
  const lines = [
    makeLine(10, '1T 2T 3T 4T Total'),
    makeLine(20, 'SomethingElse 1.000,00 2.000,00 3.000,00 4.000,00 10.000,00'),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect when next line does not start with Importes');
}

// Test 6: no match – non-numeric values
console.log('\n── Test 6: non-numeric values in Importes line ──');
{
  const lines = [
    makeLine(10, '1T 2T 3T 4T Total'),
    makeLine(20, 'Importes abc def ghi jkl mno'),
  ];
  const result = detectQuarterTable(lines);
  assert(result === null, 'should NOT detect with non-numeric values');
}

// Test 7: empty line between header and importes is skipped
console.log('\n── Test 7: empty line between header and Importes ──');
{
  const lines = [
    makeLine(10, '1T 2T 3T 4T Total'),
    { y: 20, items: [] },  // empty line
    makeLine(30, 'Importes 6.357,63 6.402,12 6.816,40 6.626,67 26.202,82'),
  ];
  const result = detectQuarterTable(lines);
  assert(result !== null, 'should detect even with empty line between header and Importes');
  assert(result?.importesIdx === 2, 'importesIdx should be 2 (after the empty line)');
}

// ── Summary ───────────────────────────────────────────────────
console.log(`\n══ Results: ${passed} passed, ${failed} failed ══`);
process.exit(failed > 0 ? 1 : 0);
