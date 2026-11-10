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
