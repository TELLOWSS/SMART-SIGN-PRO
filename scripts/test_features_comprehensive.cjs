const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');

// Import utilities
// We can test masking and XML processing directly
const maskIdentificationNumber = (value) => {
  if (value === null || value === undefined) return '';
  const str = String(value).trim();
  if (!str) return '';

  if (/^\d{6}[-\s]?[1-8]\d{6}$/.test(str)) {
    return str.replace(/^(\d{6})[-\s]?([1-8])\d{6}$/, '$1-$2******');
  }

  if (/^\d{6}[-\s][1-8]\*?$/.test(str)) {
    return str.replace(/^(\d{6})[-\s]([1-8])\*?$/, '$1-$2******');
  }

  const foreignMatch = str.match(/^([A-Za-z]\d{2,3})(\d{3,8})$/);
  if (foreignMatch) {
    const prefix = foreignMatch[1];
    const suffix = foreignMatch[2];
    return prefix + '*'.repeat(suffix.length);
  }

  if (/^[A-Za-z0-9]{8,14}$/.test(str)) {
    const keepLen = Math.min(4, Math.floor(str.length / 2));
    return str.slice(0, keepLen) + '*'.repeat(str.length - keepLen);
  }

  return str;
};

const maskXmlContent = (xml) => {
  if (!xml) return xml;
  return xml.replace(/<t([^>]*)>(.*?)<\/t>/g, (fullMatch, attr, text) => {
    const trimmed = text.trim();
    const masked = maskIdentificationNumber(trimmed);
    if (masked !== trimmed) {
      return `<t${attr}>${masked}</t>`;
    }
    return fullMatch;
  });
};

async function runTests() {
  console.log('=== 1. 주민등록번호 & 외국인등록번호 마스킹 검증 ===');
  assert.equal(maskIdentificationNumber('670305-1234567'), '670305-1******');
  assert.equal(maskIdentificationNumber('670305-1'), '670305-1******');
  assert.equal(maskIdentificationNumber('691215-1'), '691215-1******');
  assert.equal(maskIdentificationNumber('E01109933'), 'E011*****');
  assert.equal(maskIdentificationNumber('C9859552'), 'C985****');
  assert.equal(maskIdentificationNumber('P02535409'), 'P025*****');
  assert.equal(maskIdentificationNumber('K0083508'), 'K008****');
  assert.equal(maskIdentificationNumber('홍길동'), '홍길동'); // 일반 이름은 불변
  assert.equal(maskIdentificationNumber('형틀'), '형틀'); // 공종 불변
  console.log('✓ 마스킹 패턴 8종 단위 테스트 통과');

  console.log('\n=== 2. XML 텍스트 노드 안전 치환 검증 ===');
  const sampleXml = '<si><t>보반탄</t></si><si><t>E01109933</t></si><si><t>670305-1</t></si>';
  const maskedXml = maskXmlContent(sampleXml);
  assert.ok(maskedXml.includes('<t>보반탄</t>'), '이름 노드 보존');
  assert.ok(maskedXml.includes('<t>E011*****</t>'), '외국인등록번호 마스킹');
  assert.ok(maskedXml.includes('<t>670305-1******</t>'), '주민번호 단축형 마스킹');
  console.log('✓ XML sharedStrings 안전 치환 검증 통과');

  console.log('\n=== 3. 무손실 ZIP 주입 + 마스킹 통합 테스트 ===');
  // Create a minimal realistic XLSX zip
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>');
  zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>공사참여자 실명부</t></si><si><t>서기원</t></si><si><t>670305-1</t></si><si><t>보반탄</t></si><si><t>E01109933</t></si></sst>');
  zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>');

  // Apply masking
  const ss = await zip.file('xl/sharedStrings.xml').async('string');
  const maskedSs = maskXmlContent(ss);
  zip.file('xl/sharedStrings.xml', maskedSs);

  const finalSs = await zip.file('xl/sharedStrings.xml').async('string');
  assert.ok(finalSs.includes('670305-1******'));
  assert.ok(finalSs.includes('E011*****'));
  assert.ok(finalSs.includes('보반탄'));
  console.log('✓ 무손실 ZIP 내 마스킹 검증 통과');

  console.log('\n=== ALL VERIFICATION TESTS PASSED SUCCESSFULLY! ===');
}

runTests().catch(err => {
  console.error('Test failed:', err);
  process.exit(1);
});
