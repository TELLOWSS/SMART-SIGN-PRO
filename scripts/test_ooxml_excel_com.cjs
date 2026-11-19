const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execSync } = require('node:child_process');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');

// Load compiled ooxmlUtils & excelService logic directly via esbuild or inline test against built TS?
// Even better: use esbuild (already in node_modules via vite!) to bundle services/ooxmlUtils.ts & services/excelUtils.ts & injectSignaturesLossless in-memory!
const esbuild = require('esbuild');

async function loadTsModule(entryPath) {
  const result = await esbuild.build({
    entryPoints: [entryPath],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    write: false,
    external: ['exceljs', 'xlsx', 'jszip'],
  });
  const code = result.outputFiles[0].text;
  const mod = { exports: {} };
  const fn = new Function('require', 'module', 'exports', code);
  fn(require, mod, mod.exports);
  return mod.exports;
}

const PNG_1X1 = new Uint8Array(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  )
);

async function createBaseWorkbook(sheetNames = ['보호구 지급대장']) {
  const wb = new ExcelJS.Workbook();
  for (const name of sheetNames) {
    const ws = wb.addWorksheet(name);
    ws.addRow(['일자', '성명', '주민등록번호', '서명']);
    ws.addRow(['20261119', '서기원', '670305-1234567', '1']);
    ws.addRow(['20261119', '보반탄', 'E01109933', '1']);
  }
  const buf = await wb.xlsx.writeBuffer();
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

function testInRealExcelCom(fileMap) {
  const tmpDir = path.join(process.env.TEMP || 'C:\\Temp', 'smart_sign_ooxml_qa');
  fs.mkdirSync(tmpDir, { recursive: true });

  const entries = [];
  for (const [label, arrBuf] of Object.entries(fileMap)) {
    const filePath = path.join(tmpDir, `${label}.xlsx`);
    fs.writeFileSync(filePath, Buffer.from(arrBuf));
    entries.push({ label, filePath });
  }

  const psScriptPath = path.join(tmpDir, 'verify_com.ps1');
  const lines = [
    '$excel = New-Object -ComObject Excel.Application',
    '$excel.Visible = $false',
    '$excel.DisplayAlerts = $false',
    '$allOk = $true',
  ];
  for (const e of entries) {
    lines.push(`try {`);
    lines.push(`  $wb = $excel.Workbooks.Open('${e.filePath.replace(/'/g, "''")}', 0, $true)`);
    lines.push(`  $shapesSummary = @()`);
    lines.push(`  for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {`);
    lines.push(`    $ws = $wb.Worksheets.Item($i)`);
    lines.push(`    $shapesSummary += "$($ws.Name):$($ws.Shapes.Count)"`);
    lines.push(`  }`);
    lines.push(`  Write-Output "PASS | ${e.label} | Saved=$($wb.Saved) | Sheets=$($wb.Worksheets.Count) | Shapes=$($shapesSummary -join ',')"`);
    lines.push(`  $wb.Close($false)`);
    lines.push(`} catch {`);
    lines.push(`  Write-Output "FAIL | ${e.label} | $($_.Exception.Message)"`);
    lines.push(`  $allOk = $false`);
    lines.push(`}`);
  }
  lines.push('$excel.Quit()');
  lines.push('[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null');
  lines.push('if (-not $allOk) { exit 1 }');

  fs.writeFileSync(psScriptPath, lines.join('\r\n'), 'utf8');
  const out = execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${psScriptPath}"`, {
    encoding: 'utf8',
    timeout: 60000,
  });
  console.log('\n=== REAL MICROSOFT EXCEL COM VERIFICATION ===');
  console.log(out.trim());
  assert.ok(!out.includes('FAIL |'), 'One or more files failed to open cleanly in Microsoft Excel!');
}

async function run() {
  const excelService = await loadTsModule(path.resolve(__dirname, '../services/excelService.ts'));
  const ooxmlUtils = await loadTsModule(path.resolve(__dirname, '../services/ooxmlUtils.ts'));
  const excelUtils = await loadTsModule(path.resolve(__dirname, '../services/excelUtils.ts'));

  const placements = [
    { col: 3, row: 1, colOff: 19050, rowOff: 9525, width: 381000, height: 190500, imageDataU8: PNG_1X1 },
    { col: 3, row: 2, colOff: 19050, rowOff: 9525, width: 381000, height: 190500, imageDataU8: PNG_1X1 },
  ];

  const excelComFiles = {};

  // Scenario A: Sheet with <extLst> & NO prior drawing (exact 보호구 지급대장 bug)
  {
    const base = await createBaseWorkbook(['보호구 지급대장']);
    const z = await JSZip.loadAsync(base);
    let s1 = await z.file('xl/worksheets/sheet1.xml').async('string');
    s1 = s1.replace(
      '</worksheet>',
      '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/></extLst></worksheet>'
    );
    z.file('xl/worksheets/sheet1.xml', s1);
    const inputBuf = await z.generateAsync({ type: 'arraybuffer' });

    const outBuf = await excelService.injectSignaturesLossless(inputBuf, placements, { maskResidentNumber: true });
    const sanitized = await excelService.sanitizeGeneratedWorkbookBuffer(outBuf);
    const check = await ooxmlUtils.validateXlsxIntegrity(sanitized);
    assert.equal(check.valid, true, `Scenario A integrity failed: ${check.issues.join('; ')}`);

    // Verify masking preserved date '20261119' while masking RRN & Foreign ID
    const outZip = await JSZip.loadAsync(sanitized);
    const ss = await outZip.file('xl/sharedStrings.xml').async('string');
    assert.ok(ss.includes('20261119'), 'Date must NOT be masked');
    assert.ok(ss.includes('670305-1******'), 'RRN must be masked');
    assert.ok(ss.includes('E011*****'), 'Foreigner ID must be masked');

    excelComFiles['scenario_A_extLst_no_drawing'] = sanitized;
    console.log('✓ Scenario A (extLst + no prior drawing + masking) integrity PASS');
  }

  // Scenario B: Existing drawing1.xml with logo + <extLst>
  {
    const base = await createBaseWorkbook(['대우건설 실명부']);
    // First inject 1 logo, then inject 2 signatures on top
    const withLogo = await excelService.injectSignaturesLossless(base, [
      { col: 0, row: 0, colOff: 0, rowOff: 0, width: 476250, height: 285750, imageDataU8: PNG_1X1 },
    ]);
    const outBuf = await excelService.injectSignaturesLossless(withLogo, placements);
    const sanitized = await excelService.sanitizeGeneratedWorkbookBuffer(outBuf);
    const check = await ooxmlUtils.validateXlsxIntegrity(sanitized);
    assert.equal(check.valid, true, `Scenario B integrity failed: ${check.issues.join('; ')}`);
    excelComFiles['scenario_B_existing_logo'] = sanitized;
    console.log('✓ Scenario B (existing logo + new signatures) integrity PASS');
  }

  // Scenario C: Default namespace <wsDr xmlns="..."> + uppercase Extension="PNG"
  {
    const base = await createBaseWorkbook(['한셀양식']);
    const z = await JSZip.loadAsync(base);
    let ct = await z.file('[Content_Types].xml').async('string');
    ct = ct.replace(
      '</Types>',
      '<Default Extension="PNG" ContentType="image/png"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
    );
    z.file('[Content_Types].xml', ct);
    z.file('xl/drawings/drawing1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>');
    z.file('xl/drawings/_rels/drawing1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
    z.file('xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
    let s1 = await z.file('xl/worksheets/sheet1.xml').async('string');
    s1 = ooxmlUtils.insertDrawingElementInWorksheetXml(s1, 'rId1');
    z.file('xl/worksheets/sheet1.xml', s1);

    const inputBuf = await z.generateAsync({ type: 'arraybuffer' });
    const outBuf = await excelService.injectSignaturesLossless(inputBuf, placements);
    const sanitized = await excelService.sanitizeGeneratedWorkbookBuffer(outBuf);
    const check = await ooxmlUtils.validateXlsxIntegrity(sanitized);
    assert.equal(check.valid, true, `Scenario C integrity failed: ${check.issues.join('; ')}`);
    excelComFiles['scenario_C_default_ns_uppercase_png'] = sanitized;
    console.log('✓ Scenario C (default namespace wsDr + self-closing rels + uppercase PNG) integrity PASS');
  }

  // Scenario D: Multi-sheet (Sheet 1 has logo in drawing1.xml, inject into Sheet 2 -> drawing2.xml)
  {
    const base = await createBaseWorkbook(['10월', '11월']);
    const withSheet1Logo = await excelService.injectSignaturesLossless(
      base,
      [{ col: 0, row: 0, colOff: 0, rowOff: 0, width: 381000, height: 190500, imageDataU8: PNG_1X1 }],
      { targetSheetIndex: 0 }
    );
    const z = await JSZip.loadAsync(withSheet1Logo);
    let s2 = await z.file('xl/worksheets/sheet2.xml').async('string');
    s2 = s2.replace('</worksheet>', '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"/></extLst></worksheet>');
    z.file('xl/worksheets/sheet2.xml', s2);
    const multiBuf = await z.generateAsync({ type: 'arraybuffer' });

    const outBuf = await excelService.injectSignaturesLossless(multiBuf, placements, { targetSheetIndex: 1 });
    const sanitized = await excelService.sanitizeGeneratedWorkbookBuffer(outBuf);
    const check = await ooxmlUtils.validateXlsxIntegrity(sanitized);
    assert.equal(check.valid, true, `Scenario D integrity failed: ${check.issues.join('; ')}`);

    const outZip = await JSZip.loadAsync(sanitized);
    assert.ok(outZip.file('xl/drawings/drawing1.xml'), 'Sheet 1 drawing1.xml must remain intact');
    assert.ok(outZip.file('xl/drawings/drawing2.xml'), 'Sheet 2 must allocate drawing2.xml');
    excelComFiles['scenario_D_multisheet_drawing2'] = sanitized;
    console.log('✓ Scenario D (multi-sheet drawing1 + drawing2 isolation) integrity PASS');
  }

  // Scenario E: Ghost rows (>1000) with self-closing valid row + calcChain.xml
  {
    const base = await createBaseWorkbook(['유령행양식']);
    const z = await JSZip.loadAsync(base);
    let s1 = await z.file('xl/worksheets/sheet1.xml').async('string');
    s1 = s1.replace(
      '</sheetData>',
      '<row r="10" ht="15" customHeight="1"/><row r="1048576"><c r="A1048576"><v>1</v></c></row></sheetData>'
    );
    z.file('xl/worksheets/sheet1.xml', s1);
    const ghostBuf = await z.generateAsync({ type: 'arraybuffer' });

    const strippedBuf = await excelUtils.stripGhostRowsFromBuffer(ghostBuf, 1000);
    const outBuf = await excelService.injectSignaturesLossless(strippedBuf, placements);
    const sanitized = await excelService.sanitizeGeneratedWorkbookBuffer(outBuf);
    const check = await ooxmlUtils.validateXlsxIntegrity(sanitized);
    assert.equal(check.valid, true, `Scenario E integrity failed: ${check.issues.join('; ')}`);
    const outZip = await JSZip.loadAsync(sanitized);
    const finalS1 = await outZip.file('xl/worksheets/sheet1.xml').async('string');
    assert.ok(finalS1.includes('r="10"'), 'Valid self-closing row r="10" must be preserved');
    assert.ok(!finalS1.includes('1048576'), 'Ghost row 1048576 must be removed');
    excelComFiles['scenario_E_ghost_rows'] = sanitized;
    console.log('✓ Scenario E (self-closing row preservation + ghost row strip) integrity PASS');
  }

  // Open EVERY SINGLE generated file in REAL Microsoft Excel COM!
  testInRealExcelCom(excelComFiles);

  console.log('\n=== ALL 5 REAL-WORLD OOXML SCENARIOS PASSED IN MICROSOFT EXCEL COM! ===');
}

run().catch(err => {
  console.error('VERIFICATION FAILED:', err);
  process.exit(1);
});

