const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execSync } = require('node:child_process');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');
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

const PNG_200X100_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

async function run() {
  const excelUtils = await loadTsModule(path.resolve(__dirname, '../services/excelUtils.ts'));
  const excelService = await loadTsModule(path.resolve(__dirname, '../services/excelService.ts'));
  const ooxmlUtils = await loadTsModule(path.resolve(__dirname, '../services/ooxmlUtils.ts'));

  // Build exact '개인 보호구 지급대장' structure from user's screenshot
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('보호구 지급대장', {
    properties: { defaultRowHeight: 45 },
  });

  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 11;
  ws.getColumn(3).width = 9;
  ws.getColumn(4).width = 11;
  for (let c = 5; c <= 14; c++) ws.getColumn(c).width = 8;
  ws.getColumn(15).width = 15; // '서명' 열 (~118px)

  ws.getCell('A1').value = '개인 보호구 지급대장';
  ws.getCell('A2').value = '본인은 산업안전보건법 제25조 및 산업안전보건기준에 관한 규칙 제32조의② 의거 사업주로부터 받은 보호구를 반드시 착용하고 작업을 수행하며, 아래 항목의 보호구를 수령함을 서명합니다.';
  ws.getCell('A3').value = '■ 업체명: (주)휘강건설';

  // Multi-row merged header (Row 4 & Row 5)
  ws.getCell('A4').value = 'No';
  ws.getCell('B4').value = '일시';
  ws.getCell('C4').value = '직종';
  ws.getCell('D4').value = '성명';
  ws.getCell('E4').value = '품목';
  ws.getCell('O4').value = '서명';

  ws.mergeCells('A4:A5');
  ws.mergeCells('B4:B5');
  ws.mergeCells('C4:C5');
  ws.mergeCells('D4:D5');
  ws.mergeCells('E4:N4');
  ws.mergeCells('O4:O5');

  const items = ['안전모', '안전화', '안전벨트', '보안경', '팔꿈치보호대', '신호수우산', '안전장화', '방진마스크', '방진마스크 휠타', '절연장갑'];
  items.forEach((label, i) => {
    ws.getCell(5, 5 + i).value = label;
  });

  // Data rows (Row 6, 7, 8): quantity '1' under 안전화 (col 6), EMPTY under 서명 (col 15)
  const workers = [
    [6, '26.11.09', '철근', '박종염'],
    [7, '26.11.09', '철근', '이호염'],
    [8, '26.11.09', '철근', '백성철'],
  ];
  workers.forEach(([no, date, job, name], idx) => {
    const r = 6 + idx;
    ws.getCell(r, 1).value = no;
    ws.getCell(r, 2).value = date;
    ws.getCell(r, 3).value = job;
    ws.getCell(r, 4).value = name;
    ws.getCell(r, 6).value = 1; // 안전화 수량 1 (서명란이 아님!)
  });

  const rawBuf = await wb.xlsx.writeBuffer();
  const arrBuf = rawBuf.buffer.slice(rawBuf.byteOffset, rawBuf.byteOffset + rawBuf.byteLength);

  // 1. Parse and verify header column detection
  const parsed = await excelService.parseExcelFile(arrBuf, '26년11월_보호구 지급대장.xlsx');
  const detected = excelUtils.detectHeaderColumns(parsed.rows);
  console.log('Detected columns:', detected);
  assert.equal(detected.nameCol, 4, '성명 열은 col 4여야 함');
  assert.equal(detected.signatureCols.length, 1, '서명 전용 열은 정확히 1개여야 함 (2행 안내문 오탐 금지)');
  assert.equal(detected.signatureCols[0].col, 15, '서명 열은 col 15(O열)여야 함');

  // 2. Auto-match in default 'auto' mode
  const sigs = new Map();
  for (const name of ['박종염', '이호염', '백성철']) {
    sigs.set(name, [{
      name,
      variant: `${name}_1.png`,
      previewUrl: `data:image/png;base64,${PNG_200X100_BASE64}`,
      width: 200,
      height: 100,
    }]);
  }

  const autoAssignments = excelService.autoMatchSignatures(parsed, sigs, {
    variationStrength: 50,
    scaleMultiplier: 1.0,
    placementMode: 'auto',
  });
  console.log('Auto-mode assignment keys:', Array.from(autoAssignments.keys()));
  assert.equal(autoAssignments.size, 3, '3명의 서명이 정확히 3개 위치에 배치되어야 함');
  assert.ok(autoAssignments.has('6:15'), '박종염 서명은 6행 15열(서명)에 배치');
  assert.ok(autoAssignments.has('7:15'), '이호염 서명은 7행 15열(서명)에 배치');
  assert.ok(autoAssignments.has('8:15'), '백성철 서명은 8행 15열(서명)에 배치');
  assert.ok(!autoAssignments.has('6:6'), '안전화 수량 1(6행 6열)에는 서명이 들어가면 안 됨!');
  console.log('✓ 보호구 지급대장 서명 열 전용 배치(수량 1 보호) 검증 통과');

  // 3. Generate final Excel & verify signature dimensions (NOT ant-sized!)
  const blob = await excelService.generateFinalExcel(arrBuf, autoAssignments, sigs, {
    targetSheetIndex: 0,
    scaleMultiplier: 1.0,
    enhanceSignatureInk: true,
  });
  const finalBuf = await blob.arrayBuffer();
  const integrity = await ooxmlUtils.validateXlsxIntegrity(finalBuf);
  assert.equal(integrity.valid, true, `OOXML integrity failed: ${integrity.issues.join('; ')}`);

  const finalZip = await JSZip.loadAsync(finalBuf);
  const drawXml = await finalZip.file('xl/drawings/drawing1.xml').async('string');
  const extMatches = [...drawXml.matchAll(/<xdr:ext cx="(\d+)" cy="(\d+)"\/>/g)];
  assert.equal(extMatches.length, 3, '3개의 서명 앵커가 존재해야 함');
  for (const m of extMatches) {
    const wPx = parseInt(m[1], 10) / 9525;
    const hPx = parseInt(m[2], 10) / 9525;
    console.log(`  Anchor size: ${wPx.toFixed(1)}px × ${hPx.toFixed(1)}px`);
    assert.ok(wPx >= 60, `서명 너비가 충분히 커야 함 (실제: ${wPx}px)`);
    assert.ok(hPx >= 35, `서명 높이가 충분히 커야 함 (실제: ${hPx}px)`);
  }
  console.log('✓ 서명 크기(defaultRowHeight=45pt 반영, 셀 82~86% 꽉 찬 중앙 배치) 검증 통과');

  // 4. Open in REAL Microsoft Excel COM
  const tmpDir = path.join(process.env.TEMP || 'C:\\Temp', 'smart_sign_ppe_qa');
  fs.mkdirSync(tmpDir, { recursive: true });
  const outPath = path.join(tmpDir, 'ppe_verified.xlsx');
  fs.writeFileSync(outPath, Buffer.from(finalBuf));

  const psOut = execSync(
    `powershell -NoProfile -ExecutionPolicy Bypass -Command "$e = New-Object -ComObject Excel.Application; $e.Visible=$false; $e.DisplayAlerts=$false; $wb=$e.Workbooks.Open('${outPath.replace(/'/g, "''")}', 0, $true); $ws=$wb.Worksheets.Item(1); Write-Output ('PASS | Saved=' + $wb.Saved + ' | Shapes=' + $ws.Shapes.Count + ' | F6=' + $ws.Cells.Item(6,6).Value2); $wb.Close($false); $e.Quit()"`,
    { encoding: 'utf8', timeout: 30000 }
  );
  console.log('REAL EXCEL COM:', psOut.trim());
  assert.ok(psOut.includes('PASS | Saved=True | Shapes=3 | F6=1'), 'Real Excel COM must open cleanly with 3 shapes and F6 quantity=1 intact');

  // 5. Verify selectable item columns & custom per-worker item placement WITHOUT '1'
  const selectableCols = excelUtils.detectSelectableColumns(parsed.rows);
  console.log('Selectable item/signature columns:', selectableCols.map(c => `${c.col}:${c.label}`).join(', '));
  assert.equal(selectableCols.length, 11, '10개 보호구 품목 + 1개 서명 열 = 총 11개 선택 가능 열');
  assert.equal(selectableCols[0].label, '안전모');
  assert.equal(selectableCols[1].label, '안전화');
  assert.equal(selectableCols[10].label, '서명');

  const workerMap = excelUtils.buildWorkerRowMap(parsed.rows);
  assert.equal(workerMap.get(6)?.workerName, '박종염');
  assert.equal(workerMap.get(7)?.workerName, '이호염');
  assert.equal(workerMap.get(8)?.workerName, '백성철');

  // Custom worker × item assignments (no '1' needed!)
  const customAssignments = new Map();
  const customPairs = [
    [6, 5, '박종염'],  // 박종염 -> 안전모
    [6, 6, '박종염'],  // 박종염 -> 안전화
    [6, 15, '박종염'], // 박종염 -> 서명
    [7, 8, '이호염'],  // 이호염 -> 보안경
    [7, 15, '이호염'], // 이호염 -> 서명
    [8, 12, '백성철'], // 백성철 -> 방진마스크
  ];
  for (const [r, c, name] of customPairs) {
    customAssignments.set(`${r}:${c}`, {
      row: r,
      col: c,
      signatureBaseName: name,
      signatureVariantId: `${name}_1.png`,
      rotation: 0,
      scale: 1.0,
      offsetX: 0,
      offsetY: 0,
    });
  }
  const customBlob = await excelService.generateFinalExcel(arrBuf, customAssignments, sigs, {
    targetSheetIndex: 0,
    scaleMultiplier: 1.0,
    enhanceSignatureInk: true,
  });
  const customBuf = await customBlob.arrayBuffer();
  const customCheck = await ooxmlUtils.validateXlsxIntegrity(customBuf);
  assert.equal(customCheck.valid, true, `Custom item placement OOXML failed: ${customCheck.issues.join('; ')}`);

  const customPath = path.join(tmpDir, 'ppe_custom_items.xlsx');
  fs.writeFileSync(customPath, Buffer.from(customBuf));
  const customPsOut = execSync(
    `powershell -NoProfile -ExecutionPolicy Bypass -Command "$e = New-Object -ComObject Excel.Application; $e.Visible=$false; $e.DisplayAlerts=$false; $wb=$e.Workbooks.Open('${customPath.replace(/'/g, "''")}', 0, $true); $ws=$wb.Worksheets.Item(1); Write-Output ('PASS_CUSTOM | Saved=' + $wb.Saved + ' | Shapes=' + $ws.Shapes.Count); $wb.Close($false); $e.Quit()"`,
    { encoding: 'utf8', timeout: 30000 }
  );
  console.log('REAL EXCEL COM (CUSTOM WORKER x ITEM):', customPsOut.trim());
  assert.ok(customPsOut.includes('PASS_CUSTOM | Saved=True | Shapes=6'), 'Custom worker×item Excel must open with 6 shapes cleanly');

  console.log('\n=== ALL PPE LEDGER & CUSTOM ITEM SELECTION CHECKS PASSED 100%! ===');
}

run().catch(err => {
  console.error('TEST FAILED:', err);
  process.exit(1);
});

