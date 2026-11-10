const assert = require('node:assert/strict');

const isSignaturePlaceholder = (value) => {
  if (!value) return false;
  const s = String(value).trim().replace(/[\s\u00A0\uFEFF]+/g, '');
  return s === '1' || s === '(1)' || s === '1.' || s === '1)' || s === '서명' || s === '서명란';
};

const extractWorkerRowsFromSheet = (rows, assignments, signatures) => {
  let nameCol = -1;
  let residentCol = -1;
  let jobCol = -1;
  let companyCol = -1;
  let headerRow = -1;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    for (const cell of row.cells) {
      if (!cell.value) continue;
      const cleanVal = String(cell.value).replace(/[\s\u00A0\uFEFF]+/g, '');
      if (nameCol === -1 && /(성명|이름|근로자|작업자)/i.test(cleanVal)) {
        nameCol = cell.col;
        headerRow = r;
      }
      if (residentCol === -1 && /(주민등록번호|주민번호|생년월일|외국인번호)/i.test(cleanVal)) {
        residentCol = cell.col;
      }
      if (jobCol === -1 && /(직위|직종|공종|직책)/i.test(cleanVal)) {
        jobCol = cell.col;
      }
      if (companyCol === -1 && /(소속|업체|협력사|회사)/i.test(cleanVal)) {
        companyCol = cell.col;
      }
    }
    if (nameCol !== -1 && headerRow !== -1) break;
  }

  assert.ok(nameCol !== -1, '성명 열 감지');
  assert.ok(residentCol !== -1, '주민등록번호 열 감지');

  const workerItems = [];

  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    const nameCell = row.cells.find(c => c.col === nameCol);
    if (!nameCell || !nameCell.value) continue;

    const workerName = String(nameCell.value).trim();
    if (!workerName || workerName === '성명' || workerName === '소계' || workerName === '합계') continue;

    const resCell = residentCol !== -1 ? row.cells.find(c => c.col === residentCol) : null;
    const jobCell = jobCol !== -1 ? row.cells.find(c => c.col === jobCol) : null;
    const compCell = companyCol !== -1 ? row.cells.find(c => c.col === companyCol) : null;

    const placeholders = [];
    let assignedVariantId;
    let assignedSignatureName;

    for (const cell of row.cells) {
      if (cell.col === nameCol) continue;
      if (!cell.value) continue;
      const cellStr = String(cell.value).replace(/[\s\u00A0\uFEFF]+/g, '');
      if (isSignaturePlaceholder(cellStr)) {
        placeholders.push({ row: cell.row, col: cell.col, address: cell.address });
        const key = `${cell.row}:${cell.col}`;
        const assignment = assignments.get(key);
        if (assignment) {
          assignedVariantId = assignment.signatureVariantId;
          assignedSignatureName = assignment.signatureBaseName;
        }
      }
    }

    if (placeholders.length > 0) {
      workerItems.push({
        rowNumber: row.index,
        workerName,
        jobTitle: jobCell?.value ? String(jobCell.value).trim() : undefined,
        affiliation: compCell?.value ? String(compCell.value).trim() : undefined,
        residentNumber: resCell?.value ? String(resCell.value).trim() : undefined,
        placeholderCells: placeholders,
        assignedVariantId,
        assignedSignatureName,
        isMatched: !!assignedVariantId,
      });
    }
  }

  return workerItems;
};

// Test with realistic Korean construction sheet structure (from user's screenshots)
const testRows = [
  {
    index: 5,
    cells: [
      { row: 5, col: 1, value: '소 속', address: 'A5' },
      { row: 5, col: 2, value: '직 위', address: 'B5' },
      { row: 5, col: 3, value: '성 명', address: 'C5' },
      { row: 5, col: 4, value: '주민등록번호', address: 'D5' },
      { row: 5, col: 5, value: '공사한내용', address: 'E5' },
      { row: 5, col: 6, value: '서 명', address: 'F5' },
    ]
  },
  {
    index: 6,
    cells: [
      { row: 6, col: 1, value: '휘강건설', address: 'A6' },
      { row: 6, col: 2, value: '형틀', address: 'B6' },
      { row: 6, col: 3, value: '보반탄', address: 'C6' },
      { row: 6, col: 4, value: 'E01109933', address: 'D6' },
      { row: 6, col: 5, value: '형틀', address: 'E6' },
      { row: 6, col: 6, value: '1', address: 'F6' },
    ]
  },
  {
    index: 7,
    cells: [
      { row: 7, col: 1, value: '휘강건설', address: 'A7' },
      { row: 7, col: 2, value: '형틀', address: 'B7' },
      { row: 7, col: 3, value: '서기원', address: 'C7' },
      { row: 7, col: 4, value: '670305-1', address: 'D7' },
      { row: 7, col: 5, value: '형틀', address: 'E7' },
      { row: 7, col: 6, value: '1', address: 'F7' },
    ]
  }
];

const assignments = new Map();
assignments.set('7:6', {
  row: 7,
  col: 6,
  signatureBaseName: '서기원',
  signatureVariantId: '서기원_1.png',
  rotation: 0,
  scale: 1.25,
  offsetX: 0,
  offsetY: 0,
});

const signatures = new Map();
signatures.set('서기원', [{ name: '서기원', variant: '서기원_1.png', previewUrl: 'blob:url1', width: 100, height: 50 }]);
signatures.set('보반탄', [{ name: '보반탄', variant: '보반탄_1.png', previewUrl: 'blob:url2', width: 100, height: 50 }]);

const items = extractWorkerRowsFromSheet(testRows, assignments, signatures);
console.log('추출된 근로자 수:', items.length);
assert.equal(items.length, 2);
assert.equal(items[0].workerName, '보반탄');
assert.equal(items[0].isMatched, false); // 미매칭 상태
assert.equal(items[1].workerName, '서기원');
assert.equal(items[1].isMatched, true); // 매칭 완료 상태
console.log('✓ extractWorkerRowsFromSheet 검증 통과 (보반탄 미매칭, 서기원 매칭 정상 분류)');
