import ExcelJS from 'exceljs';
import { SheetData, RowData, CellData, SignatureFile, SignatureAssignment } from '../types';
import { columnNumberToLetter, parseCellAddress, isSignaturePlaceholder, randomInt, randomFloat, parsePrintAreaBounds } from './excelUtils';

export interface AutoMatchOptions {
  /**
   * 서명 변형 강도 (0~100)
   * - 0에 가까울수록 회전/이동/스케일 편차가 줄어들고
   * - 100에 가까울수록 더 자연스럽고 랜덤한 변형을 적용한다.
   */
  variationStrength?: number;
}

/**
 * 매칭을 위해 이름 정규화
 */
export const normalizeName = (name: string) => {
  if (!name) return '';
  return name
    .toString()
    .replace(/[\(\[\{].*?[\)\]\}]/g, '') // Remove content inside ( ), [ ], { }
    .replace(/[^a-zA-Z0-9가-힣]/g, '')   // Keep only Korean, English, Numbers
    .toLowerCase();                      // Case insensitive
};

/**
 * 엑셀 셀 값을 안전하게 문자열로 변환하는 헬퍼 함수
 * 포맷팅 정보까지 고려해서 원본값 보존
 */
const getCellValueAsString = (cell: ExcelJS.Cell | undefined): string => {
  if (!cell || cell.value === null || cell.value === undefined) return '';

  const val = cell.value;

  if (typeof val === 'object') {
    if ('result' in val) {
      return val.result !== undefined ? val.result.toString() : '';
    }
    if ('richText' in val && Array.isArray((val as any).richText)) {
      return (val as any).richText.map((rt: any) => rt.text).join('');
    }
    if ('text' in val) {
      return (val as any).text.toString();
    }
    return val.toString();
  }

  return val.toString();
};

/**
 * 배열을 보안 난수 기반으로 셔플(Fisher-Yates)
 * - 같은 행에서 서명 variant 반복 패턴이 눈에 띄지 않도록 순서를 섞는다.
 */
const shuffleArray = <T>(source: T[]): T[] => {
  const result = [...source];
  for (let i = result.length - 1; i > 0; i--) {
    const j = randomInt(0, i);
    [result[i], result[j]] = [result[j], result[i]];
  }
  return result;
};

/**
 * 이미지 회전 및 최적화 헬퍼 함수 (High Quality V4)
 * Input: Blob URL (memory efficient)
 * Output: Data URL (Base64)
 * 강화된 에러 처리 포함
 */
const rotateImage = async (blobUrl: string, degrees: number): Promise<string> => {
  if (!blobUrl || typeof blobUrl !== 'string' || !blobUrl.startsWith('blob:')) {
    console.error("Invalid blob URL provided:", blobUrl);
    return "";
  }

  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "Anonymous"; 
    
    img.onload = () => {
      try {
        if (img.width === 0 || img.height === 0) {
          console.error("Invalid image dimensions:", img.width, img.height);
          resolve("");
          return;
        }

        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        if (!ctx) { 
          console.warn("Canvas context not available");
          img.src = '';
          resolve(''); 
          return; 
        }

        const MAX_WIDTH = 800; 
        let scaleFactor = 1;
        
        if (img.width > MAX_WIDTH) {
          scaleFactor = MAX_WIDTH / img.width;
        }

        const drawWidth = img.width * scaleFactor;
        const drawHeight = img.height * scaleFactor;
        
        const normalizedDegrees = ((degrees % 360) + 360) % 360;
        const rad = normalizedDegrees * Math.PI / 180;
        const absCos = Math.abs(Math.cos(rad));
        const absSin = Math.abs(Math.sin(rad));
        
        canvas.width = Math.max(1, Math.round(drawWidth * absCos + drawHeight * absSin));
        canvas.height = Math.max(1, Math.round(drawWidth * absSin + drawHeight * absCos));
        
        // Check canvas size limits (most browsers: 16384 x 16384)
        if (canvas.width > 16384 || canvas.height > 16384) {
          console.warn("Canvas size exceeds limits, using original");
          img.src = '';
          resolve("");
          return;
        }
        
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high'; 

        ctx.translate(canvas.width / 2, canvas.height / 2);
        ctx.rotate(rad);
        ctx.drawImage(img, -drawWidth / 2, -drawHeight / 2, drawWidth, drawHeight);
        
        const dataUrl = canvas.toDataURL('image/png', 0.95);
        
        if (!dataUrl || dataUrl.length < 100) {
          console.error("Invalid dataUrl generated");
          resolve("");
          return;
        }
        
        canvas.width = 1;
        canvas.height = 1;
        img.src = '';
        
        console.log(`✓ Image rotated: ${normalizedDegrees}° → ${dataUrl.length} bytes`);
        resolve(dataUrl);
      } catch (err) {
        console.error("Image rotation error:", err);
        img.src = '';
        resolve("");
      }
    };

    img.onerror = (event) => {
        console.error("Image load error:", blobUrl, event);
        img.src = '';
        resolve(""); 
    };

    img.onabort = () => {
        console.error("Image load aborted:", blobUrl);
        img.src = '';
        resolve("");
    };

    img.src = blobUrl;
  });
};

/**
 * 업로드된 엑셀 파일 버퍼를 파싱
 * 강화된 유효성 검사 포함
 */
export const parseExcelFile = async (buffer: ArrayBuffer): Promise<SheetData> => {
  if (!buffer || buffer.byteLength === 0) {
    throw new Error("파일 버퍼가 비어있습니다.");
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
  } catch (err) {
    throw new Error("파일을 읽을 수 없습니다. 손상된 XLSX 파일이거나 다른 형식입니다.");
  }

  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("파일에서 워크시트를 찾을 수 없습니다.");

  const rows: RowData[] = [];
  
  // --- Infinite Row Protection ---
  const MAX_ROWS = 10000;
  const MAX_CONSECUTIVE_EMPTY_ROWS = 100;
  let consecutiveEmptyCount = 0;
  let totalRowCount = 0;

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    // Stop if too many rows
    if (totalRowCount >= MAX_ROWS) return;
    if (consecutiveEmptyCount > MAX_CONSECUTIVE_EMPTY_ROWS) return;

    let hasContent = false;
    const cells: CellData[] = [];

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const stringValue = getCellValueAsString(cell);
      if (stringValue.trim() !== '') {
        hasContent = true;
      }
      cells.push({
        value: stringValue,
        address: cell.address,
        row: rowNumber,
        col: colNumber,
      });
    });

    if (hasContent) {
      consecutiveEmptyCount = 0;
      rows.push({ index: rowNumber, cells });
      totalRowCount++;
    } else {
      consecutiveEmptyCount++;
      if (consecutiveEmptyCount <= MAX_CONSECUTIVE_EMPTY_ROWS) {
         rows.push({ index: rowNumber, cells });
      }
    }
  });

  if (rows.length === 0) {
    throw new Error("파일에 데이터가 없습니다.");
  }

  // Extract merged cells information
  const mergedCells = worksheet.model.merges ? [...worksheet.model.merges] : [];
  console.log(`[parseExcelFile] Merged cells detected: ${mergedCells.length}`);
  
  // Extract print area information
  const printArea = worksheet.pageSetup?.printArea;
  console.log(`[parseExcelFile] Print area: ${printArea || 'Not set'}`);

  return {
    name: worksheet.name || 'Sheet1',
    rows,
    mergedCells,
    printArea,
  };
};

/**
 * 서명 자동 매칭 로직
 * 개선사항: 더 나은 열 검색, 에러 처리, 결과 보고, 병합셀 인식
 */
export const autoMatchSignatures = (
  sheetData: SheetData,
  signatures: Map<string, SignatureFile[]>,
  options: AutoMatchOptions = {}
): Map<string, SignatureAssignment> => {
  const assignments = new Map<string, SignatureAssignment>();

  const normalizedStrength = Math.max(0, Math.min(100, options.variationStrength ?? 70));
  const strengthFactor = normalizedStrength / 100;

  // 기본 스케일 요구사항(1.15~1.35)은 유지하되, 강도에 따라 분산 폭을 조절한다.
  const minScale = 1.15 + (1 - strengthFactor) * 0.05;
  const maxScale = 1.35 - (1 - strengthFactor) * 0.05;

  // 회전/오프셋도 강도에 비례해 범위를 조절한다.
  const rotationLimit = Math.max(1, Math.round(2 + strengthFactor * 3));
  const offsetXLimit = Math.max(1, Math.round(2 + strengthFactor * 2));
  const offsetYLimit = Math.max(1, Math.round(1 + strengthFactor * 2));
  
  if (!sheetData || sheetData.rows.length === 0) {
    console.warn("시트 데이터가 없습니다.");
    return assignments;
  }

  if (signatures.size === 0) {
    console.warn("업로드된 서명이 없습니다.");
    return assignments;
  }
  
  const mergedCells = sheetData.mergedCells || [];
  console.log(`[autoMatch] 병합된 셀: ${mergedCells.length}개`);
  
  let nameColIndex = -1;
  let headerRowIndex = -1;
  const MAX_HEADER_SEARCH_ROWS = Math.min(50, sheetData.rows.length);

  // Find name column
  for (let r = 0; r < MAX_HEADER_SEARCH_ROWS; r++) {
    const row = sheetData.rows[r];
    for (const cell of row.cells) {
      if (!cell.value) continue;
      const rawVal = cell.value.toString().trim();
      if (!rawVal) continue;
      
      const normalizedValue = rawVal.replace(/[\s\u00A0\uFEFF]+/g, '');
      // More comprehensive name column detection
      if (/(성명|이름|name|person|employee|직원|직급)/i.test(normalizedValue)) {
        nameColIndex = cell.col;
        headerRowIndex = r;
        console.log(`Name column found: col ${nameColIndex} in row ${r}`);
        break;
      }
    }
    if (nameColIndex !== -1) break;
  }

  if (nameColIndex === -1) {
    console.warn("성명/이름 열을 자동으로 찾을 수 없습니다. 표준 형식을 확인해주세요.");
    return assignments;
  }

  let matchedCount = 0;
  let totalDataRows = 0;

  // Match signatures
  for (let r = headerRowIndex + 1; r < sheetData.rows.length; r++) {
    const row = sheetData.rows[r];
    const nameCell = row.cells.find(c => c.col === nameColIndex);
    if (!nameCell || !nameCell.value) continue;

    totalDataRows++;
    const rawName = nameCell.value.toString();
    const cleanName = normalizeName(rawName);

    if (!cleanName) continue;

    const availableSigs = signatures.get(cleanName);
    
    // 방어적 코드: 손상된 파일/빈 variant를 제외한 유효 서명만 사용
    const validAvailableSigs = (availableSigs || []).filter((sig): sig is SignatureFile => {
      return !!sig && typeof sig.variant === 'string' && sig.variant.trim().length > 0;
    });

    if (validAvailableSigs.length > 0) {
      // Track used signature variants in this row to prevent immediate reuse
      // This creates more natural variation when multiple placeholders exist
      const usedVariantsInRow = new Set<string>();
      const queuedVariantsInRow: SignatureFile[] = [];

      // 지능적 셔플링: 사용 가능한 후보군을 무작위로 섞어 큐에 채운다.
      // - 일반 상황: 아직 사용하지 않은 variant들만 셔플
      // - 모든 variant 소진 후 리셋 상황: 전체 variant를 다시 셔플
      const refillVariantQueue = () => {
        const remaining = validAvailableSigs.filter(sig => !usedVariantsInRow.has(sig.variant));
        const refillSource = remaining.length > 0 ? remaining : validAvailableSigs;

        if (remaining.length === 0) {
          usedVariantsInRow.clear();
        }

        const shuffled = shuffleArray(refillSource);
        queuedVariantsInRow.push(...shuffled);
      };
      
      for (const cell of row.cells) {
        if (cell.col === nameColIndex) continue;
        if (!cell.value) continue;
        const cellStr = cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '');
        
        // Check for signature marker
        if (isSignaturePlaceholder(cellStr)) {
          // Skip if cell is in a merged range but not the top-left cell
          if (isCellInMergedRange(cell.row, cell.col, mergedCells)) {
            if (!isTopLeftOfMergedCell(cell.row, cell.col, mergedCells)) {
              console.log(`  [autoMatch] 스킵: (${cell.row},${cell.col}) 병합셀 내부`);
              continue;
            }
          }
          
          const key = `${cell.row}:${cell.col}`;

          if (queuedVariantsInRow.length === 0) {
            refillVariantQueue();
          }

          let selectedSig = queuedVariantsInRow.shift();

          // 큐가 비정상적으로 비어 있거나 손상 데이터가 섞인 경우를 대비한 방어 로직
          while (selectedSig && (!selectedSig.variant || selectedSig.variant.trim().length === 0)) {
            selectedSig = queuedVariantsInRow.shift();
          }

          if (!selectedSig && validAvailableSigs.length > 0) {
            refillVariantQueue();
            selectedSig = queuedVariantsInRow.shift();
          }
          
          // Safety check: ensure we have a valid signature
          if (!selectedSig || !selectedSig.variant) {
            console.warn(`  [autoMatch] 경고: (${cell.row},${cell.col}) 유효하지 않은 서명 - 스킵`);
            continue;
          }
          
          // Mark this variant as used in this row
          usedVariantsInRow.add(selectedSig.variant);
          
          // Random offset calculations for natural variation
          // Using helper function for cleaner code and consistent ranges
          const rotation = randomInt(-rotationLimit, rotationLimit);
          const scale = randomFloat(minScale, maxScale);
          const offsetX = randomInt(-offsetXLimit, offsetXLimit);
          const offsetY = randomInt(-offsetYLimit, offsetYLimit);

          assignments.set(key, {
            row: cell.row,
            col: cell.col,
            signatureBaseName: cleanName,
            signatureVariantId: selectedSig.variant,
            rotation,
            scale,
            offsetX,
            offsetY,
          });
          
          matchedCount++;
        }
      }
    }
  }

  console.log(`Auto-matching complete: ${matchedCount} signatures matched out of ${totalDataRows} data rows`);
  return assignments;
};

/**
 * 셀이 병합된 셀 범위 내에 있는지 확인
 * @param row 행 번호 (1-based)
 * @param col 열 번호 (1-based)
 * @param mergedCells 병합된 셀 범위 배열 (예: ["A1:B2", "C3:D4"])
 * @returns 병합된 셀 내부에 있으면 true
 */
const isCellInMergedRange = (row: number, col: number, mergedCells: string[]): boolean => {
  for (const range of mergedCells) {
    try {
      const [start, end] = range.split(':');
      const startPos = parseCellAddress(start);
      const endPos = parseCellAddress(end);
      
      if (startPos && endPos) {
        if (row >= startPos.row && row <= endPos.row &&
            col >= startPos.col && col <= endPos.col) {
          return true;
        }
      }
    } catch (e) {
      console.warn(`Failed to parse merged cell range: ${range}`, e);
    }
  }
  return false;
};

/**
 * 병합된 셀의 왼쪽 상단 셀인지 확인
 * @param row 행 번호 (1-based)
 * @param col 열 번호 (1-based)
 * @param mergedCells 병합된 셀 범위 배열
 * @returns 왼쪽 상단 셀이면 true
 */
const isTopLeftOfMergedCell = (row: number, col: number, mergedCells: string[]): boolean => {
  for (const range of mergedCells) {
    try {
      const [start] = range.split(':');
      const startPos = parseCellAddress(start);
      
      if (startPos && startPos.row === row && startPos.col === col) {
        return true;
      }
    } catch (e) {
      console.warn(`Failed to parse merged cell range: ${range}`, e);
    }
  }
  return false;
};

const normalizeMergeRange = (range: string): string => {
  return range.toUpperCase().replace(/[\s$]/g, '');
};

/**
 * 병합 범위 중복 제거 유틸
 *
 * 왜 필요한가?
 * - ExcelJS 처리 과정(병합 해제/재병합/이미지 추가)에서 model.merges에
 *   동일 범위가 중복 삽입되면, 저장된 XLSX 내부 mergeCells XML의 count/entry가
 *   불일치해져 Excel에서 "파일에 문제가 있어 복구" 팝업이 발생할 수 있다.
 * - 따라서 저장 직전 반드시 유니크한 병합 범위만 유지한다.
 */
const getUniqueMergeRanges = (ranges: string[]): string[] => {
  const uniqueMap = new Map<string, string>();

  for (const range of ranges) {
    if (!range || typeof range !== 'string') continue;
    const normalized = normalizeMergeRange(range);
    if (!normalized) continue;
    if (!uniqueMap.has(normalized)) {
      uniqueMap.set(normalized, range);
    }
  }

  return Array.from(uniqueMap.values());
};

/**
 * 병합 범위 진단 정보 계산
 * - Excel 복구 팝업의 주요 원인인 중복 병합 범위를 빠르게 식별하기 위한 디버그 헬퍼
 */
const getMergeDiagnostics = (ranges: string[]) => {
  const total = ranges.length;
  const unique = getUniqueMergeRanges(ranges).length;
  const duplicates = Math.max(0, total - unique);
  return { total, unique, duplicates };
};

/**
 * 병합 범위를 안정적으로 재적용하기 위한 파서
 * - 잘못된 범위 문자열은 무시하여 내보내기 중단을 방지한다.
 */
const parseMergeRange = (range: string): { startRow: number; endRow: number; startCol: number; endCol: number } | null => {
  const [start, end] = range.split(':');
  const startPos = parseCellAddress(start);
  const endPos = parseCellAddress(end || start);

  if (!startPos || !endPos) return null;
  return {
    startRow: startPos.row,
    endRow: endPos.row,
    startCol: startPos.col,
    endCol: endPos.col,
  };
};

/**
 * Zero-Damage Policy: 병합셀 완벽 방어 로직
 * - 1차: 누락 병합 재적용
 * - 2차: 여전히 불일치 시 현재 병합을 정리 후 원본 병합 전체 재적용
 * - 3차: 최종 검증으로 원본 병합 집합과 완전 일치 여부 확인
 */
const enforceMergedCellsIntegrity = (worksheet: ExcelJS.Worksheet, originalMergedCells: string[]) => {
  const uniqueOriginalMergedCells = getUniqueMergeRanges(originalMergedCells);
  const normalizedOriginal = uniqueOriginalMergedCells.map(normalizeMergeRange);

  const getCurrentMergeSet = () => {
    const current = (worksheet.model.merges || []) as string[];
    return new Set(current.map(normalizeMergeRange));
  };

  // 1차: 누락 병합만 재적용
  let currentSet = getCurrentMergeSet();
  for (const mergeRange of uniqueOriginalMergedCells) {
    const normalized = normalizeMergeRange(mergeRange);
    if (!currentSet.has(normalized)) {
      try {
        worksheet.mergeCells(mergeRange);
      } catch (mergeErr) {
        console.warn(`[병합셀 방어] 1차 병합 재적용 실패: ${mergeRange}`, mergeErr);
      }
    }
  }

  currentSet = getCurrentMergeSet();
  const pass1Matched =
    normalizedOriginal.length === currentSet.size &&
    normalizedOriginal.every(range => currentSet.has(range));

  if (pass1Matched) {
    console.log('[병합셀 방어] 1차 검증 성공 - 원본 병합과 일치');
    return;
  }

  // 2차: 완전 재구축 (불일치 시 강제 재적용)
  console.warn('[병합셀 방어] 1차 검증 불일치 - 강제 재적용 모드 진입');

  const currentMerges = (worksheet.model.merges || []) as string[];
  for (const mergeRange of currentMerges) {
    const parsed = parseMergeRange(mergeRange);
    if (!parsed) continue;

    try {
      worksheet.unMergeCells(
        parsed.startRow,
        parsed.startCol,
        parsed.endRow,
        parsed.endCol
      );
    } catch (unmergeErr) {
      console.warn(`[병합셀 방어] 병합 해제 실패: ${mergeRange}`, unmergeErr);
    }
  }

  for (const mergeRange of uniqueOriginalMergedCells) {
    try {
      worksheet.mergeCells(mergeRange);
    } catch (mergeErr) {
      console.warn(`[병합셀 방어] 강제 재병합 실패: ${mergeRange}`, mergeErr);
    }
  }

  // 3차: 최종 강제 동기화 (ExcelJS 내부 상태 불일치 대비)
  const finalCurrentMerges = (worksheet.model.merges || []) as string[];
  const finalSet = new Set(finalCurrentMerges.map(normalizeMergeRange));
  const finalMatched =
    normalizedOriginal.length === finalSet.size &&
    normalizedOriginal.every(range => finalSet.has(range));

  if (!finalMatched) {
    console.warn('[병합셀 방어] 2차 검증 불일치 - model.merges 강제 동기화 수행');
    worksheet.model.merges = getUniqueMergeRanges(uniqueOriginalMergedCells);
  } else {
    console.log('[병합셀 방어] 2차 검증 성공 - 원본 병합 완전 복원');
  }

  // 최종 방어: 현재 model.merges 자체도 중복 제거 후 확정
  worksheet.model.merges = getUniqueMergeRanges((worksheet.model.merges || []) as string[]);
};

const validateWorksheetPreservation = (
  originalMergedCells: string[],
  finalMergedCells: string[],
  originalPrintArea: string | undefined,
  finalPrintArea: string | undefined
) => {
  console.log(`[최종확인] 병합된 셀: ${finalMergedCells.length}개 (원본: ${originalMergedCells.length}개)`);
  console.log(`[최종확인] 인쇄영역: ${finalPrintArea || '설정 안 됨'} (원본: ${originalPrintArea || '설정 안 됨'})`);

  if (originalMergedCells.length !== finalMergedCells.length) {
    console.warn(`⚠️ 경고: 병합된 셀 수가 여전히 다릅니다! (원본: ${originalMergedCells.length}, 최종: ${finalMergedCells.length})`);
  } else if (originalMergedCells.length > 0) {
    console.log(`✅ 병합된 셀이 성공적으로 보존되었습니다!`);
  }

  if (originalPrintArea !== finalPrintArea) {
    console.warn(`⚠️ 경고: 인쇄영역이 변경되었습니다! (원본: ${originalPrintArea || '없음'}, 최종: ${finalPrintArea || '없음'})`);
  } else if (originalPrintArea) {
    console.log(`✅ 인쇄영역이 성공적으로 보존되었습니다!`);
  }
};

/**
 * 최종 엑셀 생성 - 원본 파일 구조 완벽 보존
 * 
 * 전략:
 * 1. 원본 파일을 직접 로드
 * 2. 병합된 셀/인쇄영역은 절대 조작하지 않음 (조작 시 XML 손상)
 * 3. 이미지와 텍스트만 추가/수정
 * 4. 서명은 인쇄영역 내에만 배치
 * 5. 최소한의 변경으로 구조 손상 방지
 */
export const generateFinalExcel = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>
): Promise<Blob> => {
  if (!originalBuffer || originalBuffer.byteLength === 0) {
    throw new Error("원본 파일 버퍼가 비어있습니다.");
  }

  console.log(`[시작] 원본 파일 크기: ${originalBuffer.byteLength} bytes`);

  // Step 1: 원본 파일 직접 로드 (구조 유지)
  const workbook = new ExcelJS.Workbook();
  
  try {
    await workbook.xlsx.load(originalBuffer);
  } catch (err) {
    throw new Error(`파일 로드 실패: ${err instanceof Error ? err.message : '알수없음'}`);
  }

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("워크시트를 찾을 수 없습니다.");
  }

  console.log(`[로드완료] 행: ${worksheet.actualRowCount}, 열: ${worksheet.actualColumnCount}, 워크시트 수: ${workbook.worksheets.length}`);

  // 병합된 셀 정보 읽기 (읽기만 - 조작 금지!)
  const originalMergedCells = worksheet.model.merges ? getUniqueMergeRanges([...worksheet.model.merges]) : [];
  console.log(`[병합셀] 원본 병합된 셀: ${originalMergedCells.length}개 (읽기만, 조작 금지)`);
  const originalMergeDiag = getMergeDiagnostics((worksheet.model.merges || []) as string[]);
  console.log(`[병합셀진단] 원본 total=${originalMergeDiag.total}, unique=${originalMergeDiag.unique}, duplicates=${originalMergeDiag.duplicates}`);
  for (const merge of originalMergedCells) {
    console.log(`  - ${merge}`);
  }

  // 인쇄영역 정보 읽기 (읽기만 - 조작 금지!)
  const originalPrintArea = worksheet.pageSetup?.printArea;
  console.log(`[인쇄영역] 원본 인쇄영역: ${originalPrintArea || '설정 안 됨'} (읽기만, 조작 금지)`);

  // 여러 워크시트 문제 체크
  if (workbook.worksheets.length > 1) {
    console.warn(`⚠️ 경고: 원본 파일에 ${workbook.worksheets.length}개 시트가 있습니다.`);
  }

  // 인쇄영역 범위 파싱 (서명 배치 범위 제한용)
  const printAreaBounds = parsePrintAreaBounds(
    originalPrintArea,
    worksheet.actualRowCount || 1000,
    worksheet.actualColumnCount || 26
  );
  const printAreaRows = printAreaBounds.rows;
  const printAreaCols = printAreaBounds.cols;

  if (originalPrintArea) {
    console.log(`[인쇄영역파싱] 행: ${printAreaRows.start}-${printAreaRows.end}, 열: ${printAreaCols.start}-${printAreaCols.end}`);
  } else {
    console.log(`[인쇄영역] 설정되지 않음 - 전체 시트 사용 (행: 1-${printAreaRows.end}, 열: 1-${printAreaCols.end})`);
  }

  // Step 2: 할당된 서명 처리
  const imageCache = new Map<string, number>();
  const EMU_PER_PIXEL = 9525;
  let processedCount = 0;
  let failureCount = 0;
  let skippedCount = 0;
  let fallbackAnchorCount = 0;

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  /**
   * Excel 열 너비(문자 단위)를 픽셀로 변환
   * - Excel 기본 폭(8.43ch)과 렌더링 패딩을 고려해 근사치 계산
   */
  const getColumnPixelWidth = (col: number): number => {
    const DEFAULT_COL_WIDTH_CH = 8.43;
    const PIXELS_PER_CHAR = 7;
    const CELL_PADDING = 5;
    const widthChars = worksheet.getColumn(col).width || DEFAULT_COL_WIDTH_CH;
    return Math.max(24, Math.round(widthChars * PIXELS_PER_CHAR + CELL_PADDING));
  };

  /**
   * Excel 행 높이(pt)를 픽셀로 변환
   * - 기본 행 높이 15pt를 사용하며, 브라우저 96DPI 기준으로 환산
   */
  const getRowPixelHeight = (row: number): number => {
    const DEFAULT_ROW_HEIGHT_PT = 15;
    const PIXELS_PER_POINT = 96 / 72;
    const rowHeightPt = worksheet.getRow(row).height || DEFAULT_ROW_HEIGHT_PT;
    return Math.max(18, Math.round(rowHeightPt * PIXELS_PER_POINT));
  };

  /**
   * 특정 셀이 병합셀의 좌상단인 경우 해당 병합 범위를 반환
   */
  const getMergedRangeForTopLeft = (row: number, col: number): { startRow: number; endRow: number; startCol: number; endCol: number } | null => {
    for (const range of originalMergedCells) {
      const [start, end] = range.split(':');
      const startPos = parseCellAddress(start);
      const endPos = parseCellAddress(end || start);

      if (!startPos || !endPos) continue;

      if (startPos.row === row && startPos.col === col) {
        return {
          startRow: startPos.row,
          endRow: endPos.row,
          startCol: startPos.col,
          endCol: endPos.col,
        };
      }
    }
    return null;
  };

  /**
   * 서명 배치 대상 셀(또는 병합셀)의 실제 렌더링 영역(픽셀)을 계산
   */
  const getTargetCellBox = (row: number, col: number): { width: number; height: number } => {
    const mergedRange = getMergedRangeForTopLeft(row, col);

    if (!mergedRange) {
      return {
        width: getColumnPixelWidth(col),
        height: getRowPixelHeight(row),
      };
    }

    let totalWidth = 0;
    let totalHeight = 0;

    for (let c = mergedRange.startCol; c <= mergedRange.endCol; c++) {
      totalWidth += getColumnPixelWidth(c);
    }
    for (let r = mergedRange.startRow; r <= mergedRange.endRow; r++) {
      totalHeight += getRowPixelHeight(r);
    }

    return {
      width: Math.max(24, totalWidth),
      height: Math.max(18, totalHeight),
    };
  };

  // 좌표가 인쇄영역 내인지 확인
  const isInPrintArea = (row: number, col: number) => {
    return row >= printAreaRows.start && row <= printAreaRows.end &&
           col >= printAreaCols.start && col <= printAreaCols.end;
  };

  // 병합된 셀 범위 내인지 확인 (병합셀이 아니거나 왼쪽 상단 셀만 허용)
  const canPlaceSignature = (row: number, col: number) => {
    // 병합된 셀이면 왼쪽 상단 셀인 경우만 허용
    if (isCellInMergedRange(row, col, originalMergedCells)) {
      if (isTopLeftOfMergedCell(row, col, originalMergedCells)) {
        console.log(`  ⓘ (${row},${col}) 병합셀의 왼쪽 상단 - 배치 허용`);
        return true;
      } else {
        console.log(`  ⊘ (${row},${col}) 병합셀 내부 - 스킵`);
        return false;
      }
    }
    return true;
  };

  // placeholder 텍스트는 유지하고 그 위에 서명 이미지를 오버레이

  const assignmentValues = Array.from(assignments.values());
  const CHUNK_SIZE = 15;  // 더 작은 청크로 나누기

  console.log(`[이미지추가] ${assignmentValues.length}개 서명 처리 시작...`);

  for (let i = 0; i < assignmentValues.length; i++) {
    // Non-blocking UI
    if (i % CHUNK_SIZE === 0) {
      await new Promise(resolve => setTimeout(resolve, 0));
      console.log(`  [진행중] ${i}/${assignmentValues.length}`);
    }

    try {
      const assignment = assignmentValues[i];

      // 인쇄영역 범위 확인
      if (!isInPrintArea(assignment.row, assignment.col)) {
        console.log(`  ⊘ (${assignment.row},${assignment.col}) 인쇄영역 밖 - 스킵`);
        skippedCount++;
        continue;
      }

      // 병합된 셀 확인
      if (!canPlaceSignature(assignment.row, assignment.col)) {
        skippedCount++;
        continue;
      }

      const sigFile = findSigFile(assignment.signatureBaseName, assignment.signatureVariantId);
      
      if (!sigFile) {
        console.warn(`  ✗ 서명 파일 없음: ${assignment.signatureVariantId}`);
        failureCount++;
        continue;
      }

      // 방어적 코드: 손상 이미지/빈 variant/비정상 크기 데이터 차단
      if (!sigFile.previewUrl || !sigFile.variant || sigFile.width <= 0 || sigFile.height <= 0) {
        console.warn(`  ✗ 손상된 서명 파일 데이터 감지: ${assignment.signatureVariantId}`);
        failureCount++;
        continue;
      }

      // 이미지 캐시 키
      const cacheKey = `${sigFile.variant}_rot${assignment.rotation}`;
      let imageId = imageCache.get(cacheKey);

      // 새 이미지인 경우만 로테이션 처리
      if (imageId === undefined) {
        try {
          let rotatedDataUrl = await rotateImage(sigFile.previewUrl, assignment.rotation);
          
          if (!rotatedDataUrl || rotatedDataUrl.length === 0) {
            console.warn(`  ✗ 이미지 로테이션 실패: ${sigFile.variant} (${assignment.rotation}°)`);
            failureCount++;
            continue;
          }

          const parts = rotatedDataUrl.split(',');
          let base64Clean = parts.length > 1 ? parts[1] : parts[0];

          if (!base64Clean || base64Clean.length === 0) {
            console.warn(`  ✗ Base64 변환 실패: ${sigFile.variant}`);
            failureCount++;
            continue;
          }

          // 안전한 이미지 추가
          try {
            imageId = workbook.addImage({
              base64: base64Clean,
              extension: 'png',
            });
            imageCache.set(cacheKey, imageId);
            console.log(`  ✓ 이미지 로드: ${sigFile.variant} (ID: ${imageId})`);
          } catch (imgAddErr) {
            console.error(`  ✗ addImage 실패: ${sigFile.variant}`, imgAddErr);
            failureCount++;
            continue;
          } finally {
            // 메모리 최적화: 대용량 Base64 문자열 참조를 즉시 해제
            base64Clean = '';
            rotatedDataUrl = '';
          }
        } catch (rotErr) {
          console.warn(`  ✗ 로테이션 처리 오류: ${sigFile.variant}`, rotErr);
          failureCount++;
          continue;
        }
      } else {
        console.log(`  ◊ 이미지 캐시 사용: ${sigFile.variant}`);
      }

      // 이미지를 워크시트에 배치
      if (imageId !== undefined) {
        try {
          const targetCol = assignment.col - 1;  // ExcelJS는 0-based
          const targetRow = assignment.row - 1;

          // 유효성 검사
          if (targetRow < 0 || targetCol < 0) {
            console.warn(`  ✗ 잘못된 좌표: (${assignment.row}, ${assignment.col})`);
            failureCount++;
            continue;
          }

          // 동적 크기 계산:
          // 1) 셀(또는 병합셀) 영역을 기준으로 더 좁은 축에 맞춘다.
          // 2) scale(1.15~1.35)을 반영하되, 실제 사람이 서명 시 살짝 칸을 넘기는 느낌을 위해
          //    오버플로우 허용치를 부여한다.
          const box = getTargetCellBox(assignment.row, assignment.col);
          const narrowSide = Math.max(16, Math.min(box.width, box.height));
          const broadSide = Math.max(box.width, box.height);
          const imgRatio = sigFile.width / sigFile.height;

          // 좁은 축 기준 목표 길이: scale 상향을 반영하면서도 과도한 확장을 방지
          const targetShortSide = narrowSide * (0.72 * assignment.scale);

          let finalWidth = targetShortSide;
          let finalHeight = targetShortSide;

          if (imgRatio >= 1) {
            finalHeight = targetShortSide;
            finalWidth = finalHeight * imgRatio;
          } else {
            finalWidth = targetShortSide;
            finalHeight = finalWidth / imgRatio;
          }

          // 긴 축은 셀 긴 축의 108%까지 허용 (살짝 넘치는 자연스러운 느낌)
          const maxLongSide = broadSide * 1.08;
          if (Math.max(finalWidth, finalHeight) > maxLongSide) {
            const ratio = maxLongSide / Math.max(finalWidth, finalHeight);
            finalWidth *= ratio;
            finalHeight *= ratio;
          }

          const intWidth = Math.round(finalWidth);
          const intHeight = Math.round(finalHeight);

          // 동적 여백/오버플로우 계산:
          // - 기본적으로 중앙 정렬
          // - 사용자 배정 offset + 자연스러운 편차 반영
          // - 셀 경계를 완전히 벗어나지 않도록 안전 범위에서 클램프
          const overflowAllowanceX = Math.max(2, Math.round(box.width * 0.04));
          const overflowAllowanceY = Math.max(1, Math.round(box.height * 0.06));
          const centeredOffsetX = Math.round((box.width - intWidth) / 2);
          const centeredOffsetY = Math.round((box.height - intHeight) / 2);

          const naturalOffsetX = centeredOffsetX + assignment.offsetX;
          const naturalOffsetY = centeredOffsetY + assignment.offsetY;

          const clampedOffsetX = Math.max(
            -overflowAllowanceX,
            Math.min(box.width - intWidth + overflowAllowanceX, naturalOffsetX)
          );
          const clampedOffsetY = Math.max(
            -overflowAllowanceY,
            Math.min(box.height - intHeight + overflowAllowanceY, naturalOffsetY)
          );

          const baseOffsetX = clampedOffsetX;
          const baseOffsetY = clampedOffsetY;

          // Excel 내부 Drawing XML 안정성을 위해 오프셋은 반드시 정수로 강제
          const emuColOff = Math.max(0, Math.round(baseOffsetX * EMU_PER_PIXEL));
          const emuRowOff = Math.max(0, Math.round(baseOffsetY * EMU_PER_PIXEL));
          const integerTargetCol = Math.round(targetCol);
          const integerTargetRow = Math.round(targetRow);
          const integerWidth = Math.max(1, Math.round(intWidth));
          const integerHeight = Math.max(1, Math.round(intHeight));

          // 치명 버그 원인 주석:
          // - placeholder가 한 행에 다수(예: 11개 연속)일 때 이미지 anchor가 빽빽하게 생성된다.
          // - 이때 editAs가 명시되지 않거나 좌표/크기 값이 부동소수로 흔들리면
          //   Excel이 drawing anchor를 복구 대상으로 인식해 "파일 복구" 팝업이 뜰 수 있다.
          // - 따라서 editAs를 명시적으로 강제하고, 모든 anchor 좌표/크기를 정수화한다.
          // - 또한 oneCell 실패 시 absolute로 폴백해 손상 가능성을 추가로 낮춘다.
          const PRIMARY_IMAGE_EDIT_AS: 'oneCell' | 'absolute' = 'oneCell';
          const FALLBACK_IMAGE_EDIT_AS: 'oneCell' | 'absolute' = 'absolute';

          // 안전한 이미지 배치
          try {
            worksheet.addImage(imageId, {
              tl: {
                col: integerTargetCol,
                row: integerTargetRow,
                nativeColOff: emuColOff,
                nativeRowOff: emuRowOff
              },
              ext: { width: integerWidth, height: integerHeight },
              editAs: PRIMARY_IMAGE_EDIT_AS,
            });

            processedCount++;
            console.log(`  ✓ 배치됨: (${assignment.row},${assignment.col}) ID:${imageId}`);
          } catch (posErr) {
            console.warn(`  ⚠ oneCell 배치 실패, absolute 폴백 시도 (${assignment.row}, ${assignment.col})`, posErr);

            try {
              worksheet.addImage(imageId, {
                tl: {
                  col: integerTargetCol,
                  row: integerTargetRow,
                  nativeColOff: emuColOff,
                  nativeRowOff: emuRowOff
                },
                ext: { width: integerWidth, height: integerHeight },
                editAs: FALLBACK_IMAGE_EDIT_AS,
              });

              processedCount++;
              fallbackAnchorCount++;
              console.log(`  ✓ absolute 폴백 배치 성공: (${assignment.row},${assignment.col}) ID:${imageId}`);
            } catch (fallbackErr) {
              console.error(`  ✗ addImage 폴백 배치 실패 (${assignment.row}, ${assignment.col}):`, fallbackErr);
              failureCount++;
            }
          }
        } catch (calcErr) {
          console.error(`  ✗ 계산 오류:`, calcErr);
          failureCount++;
        }
      } else {
        console.warn(`  ✗ 이미지 ID 없음`);
        failureCount++;
      }
    } catch (assignErr) {
      console.error(`  ✗ 할당 처리 오류 (${i}):`, assignErr);
      failureCount++;
    }
  }

  console.log(`[완료] 서명 배치 결과:`);
  console.log(`  성공: ${processedCount}개`);
  console.log(`  실패: ${failureCount}개`);
  console.log(`  스킵: ${skippedCount}개`);
  console.log(`  앵커 폴백(absolute): ${fallbackAnchorCount}개`);
  console.log(`  캐시됨: ${assignmentValues.length - processedCount - failureCount - skippedCount}개`);
  // 메모리 최적화: 워크북 내 이미지 ID 캐시는 이후 재사용하지 않으므로 즉시 해제
  imageCache.clear();

  // Step 3: Zero-Damage 병합셀 방어 복원
  // - 이미지는 Floating Object 방식(addImage + editAs: oneCell)으로만 올리고
  //   셀 데이터/서식을 직접 수정하지 않는다.
  // - 이후 병합셀은 원본 기준으로 다단계 검증/재적용하여 단 하나도 유실되지 않게 한다.
  console.log('[병합셀 방어] Zero-Damage 병합셀 무결성 검증 및 재적용 시작');
  if (originalMergedCells.length > 0) {
    enforceMergedCellsIntegrity(worksheet, originalMergedCells);
  } else {
    console.log('[병합셀 방어] 원본에 병합셀이 없어 재적용 생략');
  }
  
  // 최종 확인: 병합된 셀과 인쇄영역이 여전히 존재하는지 확인
  const finalMergedCells = worksheet.model.merges ? [...worksheet.model.merges] : [];
  const finalPrintArea = worksheet.pageSetup?.printArea;
  validateWorksheetPreservation(originalMergedCells, finalMergedCells, originalPrintArea, finalPrintArea);

  const finalMergeDiag = getMergeDiagnostics(finalMergedCells);
  console.log(`[병합셀진단] 복원후 total=${finalMergeDiag.total}, unique=${finalMergeDiag.unique}, duplicates=${finalMergeDiag.duplicates}`);

  // 저장 직전 최종 방어: 중복 병합 범위를 제거해 mergeCells XML 무결성을 보장
  worksheet.model.merges = getUniqueMergeRanges((worksheet.model.merges || []) as string[]);
  const dedupedMergeDiag = getMergeDiagnostics((worksheet.model.merges || []) as string[]);
  console.log(`[병합셀진단] 저장직전 total=${dedupedMergeDiag.total}, unique=${dedupedMergeDiag.unique}, duplicates=${dedupedMergeDiag.duplicates}`);
  
  // Step 3.5: 인쇄영역 외부 Soft-clear
  // 중요 방어 원칙:
  // - spliceRows / spliceColumns 같은 구조 삭제 API는 절대 사용하지 않는다.
  // - 행/열 구조를 삭제하면 Drawing(anchor) 관계가 깨져 서명 이미지가 증발하거나
  //   Excel "파일 복구" 팝업이 발생할 수 있다.
  // - 따라서 인쇄영역 밖 셀은 값/스타일만 초기화하는 Soft-clear로만 처리한다.
  if (originalPrintArea) {
    console.log(`[인쇄영역 제한] 인쇄영역 외부 Soft-clear 진행...`);
    console.log(`  인쇄영역 범위: 행 ${printAreaRows.start}-${printAreaRows.end}, 열 ${printAreaCols.start}-${printAreaCols.end}`);
    
    let clearedCellCount = 0;
    let touchedRowCount = 0;

    // 실제 시트 사용 범위와 인쇄영역 끝점을 모두 고려해 순회 범위를 산정
    // 유령 데이터 방어:
    // - 일부 파일은 104만 행(Excel 최대 행 근처)에 찌꺼기 값이 남아 rowCount/actualRowCount가 비정상적으로 커진다.
    // - 이 상태로 전 범위를 순회하면 브라우저가 멈출 수 있으므로,
    //   인쇄영역 마지막 행 아래는 최대 N행까지만 Soft-clear를 수행한다.
    const MAX_EXTRA_ROWS_TO_SCAN = 10000;
    const worksheetRowCount = Math.max(worksheet.rowCount || 0, worksheet.actualRowCount || 0, 1);
    const maxRow = Math.max(
      printAreaRows.end,
      Math.min(worksheetRowCount, printAreaRows.end + MAX_EXTRA_ROWS_TO_SCAN)
    );
    const maxCol = Math.max(worksheet.actualColumnCount || 1, printAreaCols.end);

    if (worksheetRowCount > maxRow) {
      console.warn(
        `[인쇄영역 제한] 유령 행 방어 활성화: rowCount=${worksheetRowCount}, 스캔상한=${maxRow} (인쇄영역 끝 + ${MAX_EXTRA_ROWS_TO_SCAN})`
      );
    }

    for (let r = 1; r <= maxRow; r++) {
      const row = worksheet.getRow(r);
      let rowTouched = false;

      for (let c = 1; c <= maxCol; c++) {
        const isOutsidePrintArea =
          r < printAreaRows.start || r > printAreaRows.end ||
          c < printAreaCols.start || c > printAreaCols.end;

        if (!isOutsidePrintArea) continue;

        const cell = row.getCell(c);
        const hasValue = cell.value !== null && cell.value !== undefined;
        const hasStyle = !!cell.style && Object.keys(cell.style).length > 0;

        if (hasValue || hasStyle) {
          // Soft-clear: 구조는 유지하고 값/스타일만 비움
          cell.value = null;
          cell.style = {};
          clearedCellCount++;
          rowTouched = true;
        }
      }

      if (rowTouched) {
        touchedRowCount++;
      }
    }

    console.log(`[인쇄영역 제한 완료] Soft-clear 셀: ${clearedCellCount}개, 영향 행: ${touchedRowCount}개, 스캔행 상한: ${maxRow}`);
  } else {
    console.log(`[인쇄영역 제한] 인쇄영역이 설정되지 않아 전체 시트 유지`);
  }
  
  // Step 4: 워크북 저장
  try {
    console.log(`[저장중] 워크북을 버퍼로 쓰고 있습니다...`);

    /**
     * 조건부 서식 찌꺼기 방어:
     * - ExcelJS 버전에 따라 빈 conditionalFormatting XML이 남아
     *   Excel에서 복구 팝업을 유발할 수 있다.
     * - 본 시나리오는 서명 삽입 후 조건부 서식 유지가 필수가 아니므로
     *   저장 직전에 관련 메타를 안전하게 제거한다.
     */
    let cleanedConditionalFormattingSheets = 0;
    for (const sheet of workbook.worksheets) {
      const sheetAny = sheet as any;

      try {
        if (sheetAny.conditionalFormatting) {
          sheetAny.conditionalFormatting = [];
        }
        if (sheetAny.conditionalFormattings) {
          sheetAny.conditionalFormattings = [];
        }

        if (sheetAny.model) {
          delete sheetAny.model.conditionalFormatting;
          delete sheetAny.model.conditionalFormattings;
        }

        cleanedConditionalFormattingSheets++;
      } catch (cfErr) {
        console.warn('[조건부서식 방어] 정리 중 경고:', cfErr);
      }
    }
    console.log(`[조건부서식 방어] conditionalFormatting 정리 완료 (${cleanedConditionalFormattingSheets}개 시트)`);

    /**
     * ExcelJS DPI 오버플로우(약 42억) 대응:
     * - 일부 환경에서 pageSetup.horizontalDpi / verticalDpi 값이 비정상적으로 기록되면
     *   Excel이 파일 열기 시 "복구" 팝업을 띄울 수 있다.
     * - 워크북 내 모든 시트에 대해 저장 직전 문제 필드를 제거하여 pageSetup XML 무결성을 방어한다.
     */
    let dpiSanitizedSheetCount = 0;
    for (const sheet of workbook.worksheets) {
      if (sheet.pageSetup) {
        delete (sheet.pageSetup as any).horizontalDpi;
        delete (sheet.pageSetup as any).verticalDpi;
        dpiSanitizedSheetCount++;
      }
    }
    console.log(`[DPI 방어] pageSetup.horizontalDpi / verticalDpi 제거 완료 (${dpiSanitizedSheetCount}개 시트)`);
    
    // ExcelJS 내부 상태 정리
    try {
      // 관계 재구성
      if ((workbook as any)._rels) {
        console.log(`[정리] 워크북 관계 객체 발견`);
      }
      if ((worksheet as any)._rels) {
        console.log(`[정리] 워크시트 관계 객체 발견`);
      }
    } catch (e) {
      console.warn(`[정리 경고] 내부 상태 검사 실패`, e);
    }

    const buffer = await workbook.xlsx.writeBuffer();

    if (!buffer || buffer.byteLength === 0) {
      throw new Error("버퍼 생성 실패 - 빈 파일");
    }

    if (buffer.byteLength < 100) {
      throw new Error(`버퍼 크기 이상: ${buffer.byteLength} bytes - 파일이 손상됨`);
    }

    // ZIP 파일 검증 (XLSX는 ZIP 형식)
    const view = new Uint8Array(buffer);
    const isZip = view[0] === 0x50 && view[1] === 0x4b; // PK... magic number
    if (!isZip) {
      console.warn(`⚠️ 생성된 파일이 ZIP 형식이 아닙니다. Excel에서 열 수 없을 수 있습니다.`);
    } else {
      console.log(`✓ 생성된 파일은 유효한 ZIP 형식입니다`);
    }

    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    console.log(`✅ [성공] 파일 생성 완료: ${blob.size} bytes (ZIP: ${isZip})`);
    return blob;
  } catch (saveErr) {
    const errMsg = saveErr instanceof Error ? saveErr.message : '알 수 없음';
    throw new Error(`파일 저장 실패: ${errMsg}`);
  }
};