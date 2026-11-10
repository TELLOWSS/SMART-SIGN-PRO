import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

/**
 * 엑셀 열 문자를 숫자로 변환 (A=1, B=2, ..., Z=26, AA=27, etc.)
 */
export const columnLetterToNumber = (letter: string): number => {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
};

/**
 * 엑셀 열 숫자를 문자로 변환 (1=A, 2=B, ..., 26=Z, 27=AA, etc.)
 */
export const columnNumberToLetter = (num: number): string => {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
};

/**
 * 셀 주소를 파싱 (예: "A1" -> {row: 1, col: 1})
 */
export const parseCellAddress = (address: string): { row: number; col: number } | null => {
  const match = address.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  
  const col = columnLetterToNumber(match[1].toUpperCase());
  const row = parseInt(match[2], 10);
  
  return { row, col };
};

/**
 * 서명 placeholder로 사용되는 값들
 */
export const SIGNATURE_PLACEHOLDERS = ['1', '(1)', '1.', '1)', 'o', 'o)', '○'];

/**
 * 값이 서명 placeholder인지 확인
 */
export const isSignaturePlaceholder = (value: string): boolean => {
  return SIGNATURE_PLACEHOLDERS.includes(value.trim());
};

/**
 * 보안 난수용 Uint32 최대값 상수
 * - 나눗셈 기반 정규화 시 일관된 기준값으로 사용
 */
const UINT32_MAX = 0xFFFFFFFF;

/**
 * Web Crypto 기반 보안 난수(0 이상 1 미만) 생성
 * - Math.random() 대신 예측 불가능한 난수를 사용하여
 *   서명 배치 패턴의 재현 가능성을 낮춘다.
 * - 보안 요구사항을 위해 crypto 미지원 환경에서는 명시적으로 오류를 발생시킨다.
 */
const secureRandom = (): number => {
  const cryptoObj = globalThis.crypto;

  if (cryptoObj?.getRandomValues) {
    const array = new Uint32Array(1);
    cryptoObj.getRandomValues(array);
    return array[0] / (UINT32_MAX + 1);
  }

  throw new Error('보안 난수 생성기(crypto.getRandomValues)를 사용할 수 없습니다.');
};

/**
 * 랜덤 정수 생성 헬퍼 함수 (min과 max 포함)
 * @param min 최소값 (포함)
 * @param max 최대값 (포함)
 * @returns min과 max 사이의 랜덤 정수
 */
export const randomInt = (min: number, max: number): number => {
  const normalizedMin = Math.ceil(Math.min(min, max));
  const normalizedMax = Math.floor(Math.max(min, max));
  return Math.floor(secureRandom() * (normalizedMax - normalizedMin + 1)) + normalizedMin;
};

/**
 * 랜덤 실수 생성 헬퍼 함수 (min 이상 max 미만)
 * - 서명 scale처럼 연속값이 필요한 경우 사용
 */
export const randomFloat = (min: number, max: number): number => {
  const normalizedMin = Math.min(min, max);
  const normalizedMax = Math.max(min, max);
  return secureRandom() * (normalizedMax - normalizedMin) + normalizedMin;
};

export interface PrintAreaBounds {
  rows: { start: number; end: number };
  cols: { start: number; end: number };
}

const isValidPrintAreaRange = (tlRow: number, brRow: number, tlCol: number, brCol: number): boolean => {
  return tlRow > 0 && brRow > 0 && tlCol > 0 && brCol > 0 &&
         tlRow <= brRow && tlCol <= brCol;
};

/**
 * Excel printArea 문자열을 안전하게 파싱
 * 지원 예시: A1:C10, Sheet1!A1:C10, $A$1:$C$10, Sheet1!$A$1:$C$10, A1
 */
export const parsePrintAreaBounds = (
  printArea: string | undefined,
  defaultRowEnd: number,
  defaultColEnd: number
): PrintAreaBounds => {
  const fallback: PrintAreaBounds = {
    rows: { start: 1, end: Math.max(1, defaultRowEnd || 1) },
    cols: { start: 1, end: Math.max(1, defaultColEnd || 1) },
  };

  if (!printArea || !printArea.trim()) {
    return fallback;
  }

  try {
    let range = printArea.trim();

    if (range.includes('!')) {
      range = range.split('!').pop() || range;
    }

    range = range.replace(/\$/g, '');
    const parts = range.split(':').map(part => part.trim());

    if (parts.length === 1) {
      const single = parseCellAddress(parts[0]);
      if (!single) return fallback;
      return {
        rows: { start: single.row, end: single.row },
        cols: { start: single.col, end: single.col },
      };
    }

    if (parts.length !== 2) {
      return fallback;
    }

    const topLeft = parseCellAddress(parts[0]);
    const bottomRight = parseCellAddress(parts[1]);

    if (!topLeft || !bottomRight) {
      return fallback;
    }

    if (!isValidPrintAreaRange(topLeft.row, bottomRight.row, topLeft.col, bottomRight.col)) {
      return fallback;
    }

    return {
      rows: { start: topLeft.row, end: bottomRight.row },
      cols: { start: topLeft.col, end: bottomRight.col },
    };
  } catch {
    return fallback;
  }
};

/**
 * 엑셀 워크시트에서 유령 행(Ghost Rows, 예: 104만 행 근처)을 사전 제거
 * - 현장 파일에서 흔히 발생하는 최대 행 부근(1,048,572행)의 더미 셀로 인한
 *   브라우저 메모리 폭증(OOM) 및 ExcelJS 로드 지연/실패를 선제 차단한다.
 * - 실제 서식 및 대우건설 로고 등 미디어는 100% 온전히 보존된다.
 */
export const stripGhostRowsFromBuffer = async (
  buffer: ArrayBuffer,
  maxRow: number = 1000
): Promise<ArrayBuffer> => {
  try {
    const zip = await JSZip.loadAsync(buffer);
    let modified = false;

    for (const filename of Object.keys(zip.files)) {
      if (filename.startsWith('xl/worksheets/sheet') && filename.endsWith('.xml') && !filename.includes('_rels')) {
        let sheetXml = await zip.file(filename)!.async('string');
        let hasGhost = false;

        const cleanedXml = sheetXml.replace(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g, (m, r) => {
          if (parseInt(r, 10) > maxRow) {
            hasGhost = true;
            return '';
          }
          return m;
        });

        if (hasGhost) {
          zip.file(filename, cleanedXml);
          modified = true;
        }
      }
    }

    if (modified) {
      console.log(`[stripGhostRowsFromBuffer] 유령 행(>${maxRow}) 선제 제거 완료`);
      return await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
    }
    return buffer;
  } catch (err) {
    console.warn('[stripGhostRowsFromBuffer] 유령 행 전처리 스킵 (원본 유지):', err);
    return buffer;
  }
};

/**
 * ExcelJS 워크북을 안전하게 로드
 * - 0차: 유령 행(104만 행) 사전 제거로 브라우저 메모리 폭증 방지 및 ExcelJS 로드 보장
 * - 1차: 원본 ArrayBuffer를 ExcelJS로 직접 로드 (모든 서식/스타일/로고 완벽 보존)
 * - 2차: ExcelJS 실패 시(구형 .xls, HTML 테이블, 특수 태그 등) SheetJS로 자동 복구 및 표준 .xlsx 변환 후 로드
 */
export const loadWorkbookSafely = async (
  buffer: ArrayBuffer
): Promise<{ workbook: ExcelJS.Workbook; normalizedBuffer: ArrayBuffer; isTranscoded: boolean }> => {
  if (!buffer || buffer.byteLength === 0) {
    throw new Error("파일 버퍼가 비어있습니다.");
  }

  // 104만 행 등 비정상 유령 행 사전 제거 (스타일 및 로고 100% 보존)
  const cleanBuffer = await stripGhostRowsFromBuffer(buffer, 1000);

  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(cleanBuffer);
    return { workbook: wb, normalizedBuffer: cleanBuffer, isTranscoded: false };
  } catch (excelErr) {
    console.warn("[loadWorkbookSafely] ExcelJS 직접 로드 실패, SheetJS 호환 모드로 복구 시도:", excelErr);

    try {
      const u8 = new Uint8Array(buffer);
      const sheetJsWb = XLSX.read(u8, {
        type: 'array',
        cellDates: true,
        cellStyles: true,
      });

      if (!sheetJsWb.SheetNames || sheetJsWb.SheetNames.length === 0) {
        throw new Error("파일 내에 워크시트가 존재하지 않습니다.");
      }

      const cleanXlsxData = XLSX.write(sheetJsWb, {
        type: 'array',
        bookType: 'xlsx',
      });
      const cleanArrayBuffer = cleanXlsxData instanceof Uint8Array ? cleanXlsxData.buffer : cleanXlsxData;

      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(cleanArrayBuffer);
      return { workbook: wb, normalizedBuffer: cleanArrayBuffer, isTranscoded: true };
    } catch (sheetJsErr) {
      const originalMsg = excelErr instanceof Error ? excelErr.message : String(excelErr);
      const sheetJsMsg = sheetJsErr instanceof Error ? sheetJsErr.message : String(sheetJsErr);
      throw new Error(
        `파일을 읽을 수 없습니다.\n` +
        `- 상세 사유: ${originalMsg || sheetJsMsg}\n` +
        `암호로 보호된 파일이거나 파일 형식이 손상되었을 수 있습니다. Excel 프로그램에서 열어 '다른 이름으로 저장(XLSX)' 후 다시 시도해주세요.`
      );
    }
  }
};

/**
 * DataTransfer 객체에서 파일 및 폴더 내부 파일을 재귀적으로 추출 (Drag and Drop 완벽 지원)
 * - 사용자가 폴더를 통째로 끌어다 놓아도 webkitGetAsEntry를 통해 모든 하위 파일/폴더를 자동 탐색한다.
 * - Windows, Mac, Chrome, Edge, Firefox, Whale, Safari 등 주요 브라우저를 완벽 지원한다.
 */
export const extractFilesFromDataTransfer = async (dataTransfer: DataTransfer): Promise<File[]> => {
  const files: File[] = [];
  const items = dataTransfer.items;

  if (items && items.length > 0 && typeof (items[0] as any).webkitGetAsEntry === 'function') {
    const queue: any[] = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.kind === 'file') {
        const entry = (item as any).webkitGetAsEntry();
        if (entry) queue.push(entry);
      }
    }

    const readAllEntries = async (dirReader: any): Promise<any[]> => {
      const entries: any[] = [];
      let batch: any[];
      do {
        batch = await new Promise<any[]>((resolve) => {
          dirReader.readEntries((results: any[]) => resolve(results || []), () => resolve([]));
        });
        entries.push(...batch);
      } while (batch.length > 0);
      return entries;
    };

    while (queue.length > 0) {
      const entry = queue.shift();
      if (!entry) continue;

      if (entry.isFile) {
        await new Promise<void>((resolve) => {
          entry.file(
            (file: File) => {
              files.push(file);
              resolve();
            },
            () => resolve()
          );
        });
      } else if (entry.isDirectory) {
        try {
          const reader = entry.createReader();
          const children = await readAllEntries(reader);
          queue.push(...children);
        } catch (e) {
          console.warn('[extractFilesFromDataTransfer] 디렉터리 읽기 오류:', e);
        }
      }
    }

    if (files.length > 0) {
      return files;
    }
  }

  // 폴더 탐색 API 미지원 브라우저 또는 일반 파일 드롭 폴백
  if (dataTransfer.files && dataTransfer.files.length > 0) {
    return Array.from(dataTransfer.files);
  }

  return [];
};

/**
 * 주민등록번호 및 외국인등록번호/여권번호 마스킹 (개인정보 보호법 및 현장 제출 감사 규정 준수)
 * - 내국인 주민번호(13자리): 670305-1234567 -> 670305-1******
 * - 내국인 단축형: 670305-1 -> 670305-1******
 * - 외국인등록번호/여권번호: E01109933 -> E011*****, C9859552 -> C985****
 */
export const maskIdentificationNumber = (value: string | number | null | undefined): string => {
  if (value === null || value === undefined) return '';
  const str = String(value).trim();
  if (!str) return '';

  // 1. 표준 주민등록번호 13자리 (하이픈 유무 무관)
  // 예: 670305-1234567 -> 670305-1******
  if (/^\d{6}[-\s]?[1-8]\d{6}$/.test(str)) {
    return str.replace(/^(\d{6})[-\s]?([1-8])\d{6}$/, '$1-$2******');
  }

  // 2. 단축형 주민등록번호 (생년월일 6자리 + 뒷자리 첫 글자)
  // 예: 670305-1 -> 670305-1******
  if (/^\d{6}[-\s][1-8]\*?$/.test(str)) {
    return str.replace(/^(\d{6})[-\s]([1-8])\*?$/, '$1-$2******');
  }

  // 3. 영문+숫자 혼합 외국인등록번호 / 여권번호 형태 (예: E01109933, C9859552, P02535409, K0083508)
  const foreignMatch = str.match(/^([A-Za-z]\d{2,3})(\d{3,8})$/);
  if (foreignMatch) {
    const prefix = foreignMatch[1];
    const suffix = foreignMatch[2];
    return prefix + '*'.repeat(suffix.length);
  }

  // 4. 일반적인 8자리 이상 주민/등록번호 패턴
  if (/^[A-Za-z0-9]{8,14}$/.test(str)) {
    const keepLen = Math.min(4, Math.floor(str.length / 2));
    return str.slice(0, keepLen) + '*'.repeat(str.length - keepLen);
  }

  return str;
};

/**
 * 엑셀 XML 내부의 텍스트 노드(<t>...</t>) 중 주민등록번호/외국인번호를 안전하게 마스킹 치환
 * - 셀 스타일이나 태그 구조는 0.1%도 손대지 않고 값 텍스트만 치환
 */
export const maskXmlContent = (xml: string): string => {
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

/**
 * 시트 데이터에서 근로자 행을 정밀 추출하여 수동 매칭 및 관리 목록 생성
 */
export const extractWorkerRowsFromSheet = (
  rows: { index: number; cells: { row: number; col: number; value: any; address: string }[] }[],
  assignments: Map<string, any>,
  signatures: Map<string, any[]>
) => {
  let nameCol = -1;
  let residentCol = -1;
  let jobCol = -1;
  let companyCol = -1;
  let headerRow = -1;

  const MAX_SEARCH = Math.min(50, rows.length);

  for (let r = 0; r < MAX_SEARCH; r++) {
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

  if (nameCol === -1) return [];

  const workerItems: {
    rowNumber: number;
    workerName: string;
    jobTitle?: string;
    affiliation?: string;
    residentNumber?: string;
    placeholderCells: { row: number; col: number; address: string }[];
    assignedVariantId?: string;
    assignedSignatureName?: string;
    isMatched: boolean;
  }[] = [];

  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r];
    const nameCell = row.cells.find(c => c.col === nameCol);
    if (!nameCell || !nameCell.value) continue;

    const workerName = String(nameCell.value).trim();
    if (!workerName || workerName === '성명' || workerName === '소계' || workerName === '합계') continue;

    const resCell = residentCol !== -1 ? row.cells.find(c => c.col === residentCol) : null;
    const jobCell = jobCol !== -1 ? row.cells.find(c => c.col === jobCol) : null;
    const compCell = companyCol !== -1 ? row.cells.find(c => c.col === companyCol) : null;

    const placeholders: { row: number; col: number; address: string }[] = [];
    let assignedVariantId: string | undefined;
    let assignedSignatureName: string | undefined;

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


