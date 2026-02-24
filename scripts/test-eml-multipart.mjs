/**
 * Test harness for buildEml multipart/related CID image embedding.
 *
 * Usage:  node scripts/test-eml-multipart.mjs
 *
 * Validates:
 *  1) EML without images still produces single-part output.
 *  2) EML with local <img src="..."> produces multipart/related with CID parts.
 *  3) Spanish diacritics survive UTF-8 encoding round-trip.
 *  4) Signature links (mailto / https) are preserved in the HTML part.
 */

// ── Shared helpers (mirrored from app.js for Node) ────────────

function base64FromBytes(bytes) {
  let binary = '';
  for (const b of bytes) {
    binary += String.fromCharCode(b);
  }
  return btoa(binary);
}

function emlBodyBase64EncodeWithWrap(html) {
  const bytes = new TextEncoder().encode(String(html || ''));
  const base64 = base64FromBytes(bytes);
  return base64.match(/.{1,76}/g)?.join('\r\n') || '';
}

function rfc2047EncodeHeaderValue(value) {
  const bytes = new TextEncoder().encode(String(value || ''));
  const base64 = base64FromBytes(bytes);
  return `=?UTF-8?B?${base64}?=`;
}

function formatRfc2822Date(date) {
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const dayName = days[date.getUTCDay()];
  const day = String(date.getUTCDate()).padStart(2, '0');
  const month = months[date.getUTCMonth()];
  const year = date.getUTCFullYear();
  const hour = String(date.getUTCHours()).padStart(2, '0');
  const minute = String(date.getUTCMinutes()).padStart(2, '0');
  const second = String(date.getUTCSeconds()).padStart(2, '0');
  return `${dayName}, ${day} ${month} ${year} ${hour}:${minute}:${second} +0000`;
}

// ── Standalone buildEml (Node version with sync image injection) ──

function buildEmlSync(bodyHtml, { toEmail, subject }, inlineImages) {
  const date = formatRfc2822Date(new Date());
  const toHeaderValue = toEmail ? toEmail : 'undisclosed-recipients:;';

  let processedHtml = bodyHtml;
  for (const img of inlineImages) {
    processedHtml = processedHtml.split(img.originalSrc).join('cid:' + img.cid);
  }

  const fullHtml = `<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head>\n<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">\n${processedHtml}\n</body></html>`;

  if (inlineImages.length === 0) {
    const lines = [
      'MIME-Version: 1.0',
      `Date: ${date}`,
      `To: ${toHeaderValue}`,
      `Subject: ${rfc2047EncodeHeaderValue(subject)}`,
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      emlBodyBase64EncodeWithWrap(fullHtml),
      ''
    ];
    return lines.join('\r\n');
  }

  const boundary = '----=_ConvertoBoundary_001';
  const lines = [
    'MIME-Version: 1.0',
    `Date: ${date}`,
    `To: ${toHeaderValue}`,
    `Subject: ${rfc2047EncodeHeaderValue(subject)}`,
    `Content-Type: multipart/related; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    emlBodyBase64EncodeWithWrap(fullHtml),
  ];

  for (const img of inlineImages) {
    lines.push(`--${boundary}`);
    lines.push(`Content-Type: ${img.mimeType}`);
    lines.push('Content-Transfer-Encoding: base64');
    lines.push(`Content-ID: <${img.cid}>`);
    lines.push(`Content-Disposition: inline; filename="${img.filename}"`);
    lines.push('');
    lines.push(img.base64Data);
  }

  lines.push(`--${boundary}--`);
  lines.push('');
  return lines.join('\r\n');
}

// ── collectLocalImages regex extraction (Node version) ────────

function extractImageSrcs(html) {
  const imgRegex = /<img\s[^>]*src\s*=\s*["']([^"']+)["'][^>]*>/gi;
  const srcs = [];
  const seen = new Set();
  let match;
  while ((match = imgRegex.exec(html)) !== null) {
    const src = match[1];
    if (/^(https?:|data:|cid:)/i.test(src)) continue;
    if (seen.has(src)) continue;
    seen.add(src);
    srcs.push(src);
  }
  return srcs;
}

// ── Tests ─────────────────────────────────────────────────────

let passed = 0;
let failed = 0;

function assert(condition, msg) {
  if (condition) {
    passed++;
    console.log(`  ✓ ${msg}`);
  } else {
    failed++;
    console.error(`  ✗ ${msg}`);
  }
}

console.log('── Test 1: Single-part EML (no images) ──');
{
  const html = '<p>Hello world</p>';
  const eml = buildEmlSync(html, { toEmail: 'test@example.com', subject: 'Test' }, []);
  assert(eml.includes('Content-Type: text/html; charset=UTF-8'), 'Has text/html content-type');
  assert(!eml.includes('multipart/related'), 'No multipart/related header');
  assert(!eml.includes('Content-ID'), 'No Content-ID');
  assert(eml.includes('MIME-Version: 1.0'), 'Has MIME-Version');
}

console.log('── Test 2: Multipart/related EML (with images) ──');
{
  const html = '<p>Hello</p><img src="signatures/assets/logo_nm.jpg" alt="logo">';
  const fakeImage = {
    originalSrc: 'signatures/assets/logo_nm.jpg',
    cid: 'sigimg-1@converto',
    mimeType: 'image/jpeg',
    filename: 'logo_nm.jpg',
    base64Data: '/9j/4AAQSkZJRgABAQ=='  // truncated fake JPEG data
  };
  const eml = buildEmlSync(html, { toEmail: 'test@example.com', subject: 'Test' }, [fakeImage]);
  assert(eml.includes('multipart/related; boundary="----=_ConvertoBoundary_001"'), 'Has multipart/related header');
  assert(eml.includes('Content-ID: <sigimg-1@converto>'), 'Has Content-ID for image');
  assert(eml.includes('Content-Disposition: inline; filename="logo_nm.jpg"'), 'Has inline disposition');
  assert(eml.includes('Content-Type: image/jpeg'), 'Image has correct MIME type');
  assert(eml.includes('------=_ConvertoBoundary_001--'), 'Has closing boundary');

  // Decode the HTML part and check CID replacement
  const htmlPartMatch = eml.match(/Content-Type: text\/html; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n([\s\S]*?)\r\n------=_ConvertoBoundary_001/);
  assert(htmlPartMatch !== null, 'HTML part found in multipart');
  if (htmlPartMatch) {
    const decoded = Buffer.from(htmlPartMatch[1].replace(/\r\n/g, ''), 'base64').toString('utf-8');
    assert(decoded.includes('cid:sigimg-1@converto'), 'HTML references cid: for image');
    assert(!decoded.includes('signatures/assets/logo_nm.jpg'), 'Original src replaced');
  }
}

console.log('── Test 3: Spanish diacritics survive encoding ──');
{
  const html = '<p>Lukyán Tazzít – Dpto. de Contabilidad</p>';
  const eml = buildEmlSync(html, { toEmail: 'test@example.com', subject: 'Notificación' }, []);
  const separatorIndex = eml.indexOf('\r\n\r\n');
  const body64 = eml.slice(separatorIndex + 4).trim();
  const decoded = Buffer.from(body64, 'base64').toString('utf-8');
  assert(decoded.includes('Lukyán'), 'Name with á preserved');
  assert(decoded.includes('Tazzít'), 'Name with í preserved');
}

console.log('── Test 4: Signature links preserved ──');
{
  const html = '<a href="mailto:conta2@nmeconomista.com">conta2@nmeconomista.com</a><br><a href="https://www.nmeconomista.com">www.nmeconomista.com</a>';
  const eml = buildEmlSync(html, { toEmail: 'test@example.com', subject: 'Test' }, []);
  const separatorIndex = eml.indexOf('\r\n\r\n');
  const body64 = eml.slice(separatorIndex + 4).trim();
  const decoded = Buffer.from(body64, 'base64').toString('utf-8');
  assert(decoded.includes('mailto:conta2@nmeconomista.com'), 'mailto link preserved');
  assert(decoded.includes('https://www.nmeconomista.com'), 'https link preserved');
}

console.log('── Test 5: extractImageSrcs only picks local paths ──');
{
  const html = `
    <img src="signatures/assets/logo.jpg" alt="local">
    <img src="https://example.com/img.png" alt="remote">
    <img src="data:image/png;base64,abc" alt="data">
    <img src="cid:existing@test" alt="cid">
    <img src="other/path.png" alt="local2">
  `;
  const srcs = extractImageSrcs(html);
  assert(srcs.length === 2, `Found 2 local images (got ${srcs.length})`);
  assert(srcs[0] === 'signatures/assets/logo.jpg', 'First local src correct');
  assert(srcs[1] === 'other/path.png', 'Second local src correct');
}

console.log('── Test 6: Duplicate images handled ──');
{
  const html = '<img src="a.jpg"><img src="a.jpg"><img src="b.png">';
  const srcs = extractImageSrcs(html);
  assert(srcs.length === 2, `Deduplicated to 2 images (got ${srcs.length})`);
}

console.log('── Test 7: Subject with diacritics is RFC2047 encoded ──');
{
  const subject = 'Notificación MODELO 347';
  const encoded = rfc2047EncodeHeaderValue(subject);
  assert(encoded.startsWith('=?UTF-8?B?'), 'Starts with UTF-8 B-encoding prefix');
  assert(encoded.endsWith('?='), 'Ends with ?= suffix');
  const inner = encoded.slice('=?UTF-8?B?'.length, -2);
  const decoded = Buffer.from(inner, 'base64').toString('utf-8');
  assert(decoded === subject, 'Subject round-trips correctly');
}

// ── Summary ───────────────────────────────────────────────────

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed > 0 ? 1 : 0);
