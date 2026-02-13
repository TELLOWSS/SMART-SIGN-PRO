import ExcelJS from 'exceljs';
import { SheetData, RowData, CellData, SignatureFile, SignatureAssignment } from '../types';
import { columnLetterToNumber, columnNumberToLetter, parseCellAddress, SIGNATURE_PLACEHOLDERS, isSignaturePlaceholder } from './excelUtils';

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
          resolve(blobUrl); 
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
        
        console.log(`✓ Image rotated: ${normalizedDegrees}° → ${dataUrl.length} bytes`);
        resolve(dataUrl);
      } catch (err) {
        console.error("Image rotation error:", err);
        resolve("");
      }
    };

    img.onerror = (event) => {
        console.error("Image load error:", blobUrl, event);
        resolve(""); 
    };

    img.onabort = () => {
        console.error("Image load aborted:", blobUrl);
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
  signatures: Map<string, SignatureFile[]>
): Map<string, SignatureAssignment> => {
  const assignments = new Map<string, SignatureAssignment>();
  
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
    
    if (availableSigs && availableSigs.length > 0) {
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
          
          const randomSigIndex = Math.floor(Math.random() * availableSigs.length);
          const selectedSig = availableSigs[randomSigIndex];
          
          // Random offset calculations for X/Y
          const rotation = Math.floor(Math.random() * 11) - 5;  // -5 to 5 degrees
          const scale = 0.95 + (Math.random() * 0.15); // 0.95 to 1.1
          const offsetX = Math.floor(Math.random() * 9) - 4; // -4 to +4 px
          const offsetY = Math.floor(Math.random() * 5) - 2;  // -2 to +2 px

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
 * 인쇄영역 범위가 유효한지 검증
 */
const isValidPrintAreaRange = (tlRow: number, brRow: number, tlCol: number, brCol: number): boolean => {
  return tlRow > 0 && brRow > 0 && tlCol > 0 && brCol > 0 && 
         tlRow <= brRow && tlCol <= brCol;
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
  const originalMergedCells = worksheet.model.merges ? [...worksheet.model.merges] : [];
  console.log(`[병합셀] 원본 병합된 셀: ${originalMergedCells.length}개 (읽기만, 조작 금지)`);
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
  let printAreaRows = { start: 1, end: worksheet.actualRowCount || 1000 };
  let printAreaCols = { start: 1, end: worksheet.actualColumnCount || 26 };
  
  if (originalPrintArea) {
    try {
      // printArea 형식 처리:
      // 1. "A1:C10" (단순 범위)
      // 2. "Sheet1!A1:C10" (시트명 포함)
      // 3. "$A$1:$C$10" (절대 참조)
      // 4. "Sheet1!$A$1:$C$10" (시트명 + 절대 참조)
      let range = originalPrintArea;
      
      // 시트명 제거
      if (range.includes('!')) {
        range = range.split('!').pop() || range;
      }
      
      // $ 기호 제거 (절대 참조)
      range = range.replace(/\$/g, '');
      
      const parts = range.split(':');
      if (parts.length === 2) {
        const [topLeft, bottomRight] = parts;
        const tlMatch = topLeft.trim().match(/^([A-Z]+)(\d+)$/i);
        const brMatch = bottomRight.trim().match(/^([A-Z]+)(\d+)$/i);
        
        if (tlMatch && brMatch) {
          const tlCol = columnLetterToNumber(tlMatch[1].toUpperCase());
          const brCol = columnLetterToNumber(brMatch[1].toUpperCase());
          const tlRow = parseInt(tlMatch[2], 10);
          const brRow = parseInt(brMatch[2], 10);
          
          // 유효성 검사
          if (isValidPrintAreaRange(tlRow, brRow, tlCol, brCol)) {
            printAreaRows = { start: tlRow, end: brRow };
            printAreaCols = { start: tlCol, end: brCol };
            console.log(`[인쇄영역파싱성공] 행: ${printAreaRows.start}-${printAreaRows.end}, 열: ${printAreaCols.start}-${printAreaCols.end}`);
          } else {
            console.warn(`[인쇄영역파싱실패] 잘못된 범위 값: ${originalPrintArea}`);
          }
        } else {
          console.warn(`[인쇄영역파싱실패] 형식 오류: ${originalPrintArea}`);
        }
      } else if (parts.length === 1 && parts[0].trim()) {
        // 단일 셀인 경우 (예: "A1")
        const cellMatch = parts[0].trim().match(/^([A-Z]+)(\d+)$/i);
        if (cellMatch) {
          const col = columnLetterToNumber(cellMatch[1].toUpperCase());
          const row = parseInt(cellMatch[2], 10);
          printAreaRows = { start: row, end: row };
          printAreaCols = { start: col, end: col };
          console.log(`[인쇄영역파싱성공] 단일 셀: (${row}, ${col})`);
        }
      } else {
        console.warn(`[인쇄영역파싱실패] 예상치 못한 형식: ${originalPrintArea}`);
      }
    } catch (parseErr) {
      console.error(`[인쇄영역파싱실패] 예외 발생:`, parseErr);
      // 인쇄영역을 파싱할 수 없을 때는 전체 시트를 사용 (기본값 유지)
      console.log(`[인쇄영역] 전체 시트 사용 (행: 1-${printAreaRows.end}, 열: 1-${printAreaCols.end})`);
    }
  } else {
    console.log(`[인쇄영역] 설정되지 않음 - 전체 시트 사용 (행: 1-${printAreaRows.end}, 열: 1-${printAreaCols.end})`);
  }

  // Step 2: 할당된 서명 처리
  const imageCache = new Map<string, number>();
  const EMU_PER_PIXEL = 9525;
  let processedCount = 0;
  let failureCount = 0;
  let skippedCount = 0;

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
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

  // 이미지 추가 전에 선택된 셀만 먼저 텍스트 정리
  console.log(`[사전처리] placeholder 텍스트 제거 중...`);
  for (const [key, assignment] of assignments) {
    try {
      const [rowStr, colStr] = key.split(':');
      const row = parseInt(rowStr, 10);
      const col = parseInt(colStr, 10);

      // 인쇄영역 범위 확인
      if (!isInPrintArea(row, col)) {
        console.log(`  ⊘ (${row},${col}) 인쇄영역 밖 - 스킵`);
        skippedCount++;
        continue;
      }

      // 병합된 셀 확인
      if (!canPlaceSignature(row, col)) {
        skippedCount++;
        continue;
      }

      const cell = worksheet.getCell(row, col);
      if (cell) {
        const cellVal = cell.value ? cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '') : '';
        
        if (isSignaturePlaceholder(cellVal)) {
          cell.value = null;
          console.log(`  ✓ (${row},${col}) 텍스트 제거`);
        }
      }
    } catch (e) {
      console.warn(`  ✗ (${key}) 텍스트 제거 오류:`, e);
    }
  }

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

      // 이미지 캐시 키
      const cacheKey = `${sigFile.variant}_rot${assignment.rotation}`;
      let imageId = imageCache.get(cacheKey);

      // 새 이미지인 경우만 로테이션 처리
      if (imageId === undefined) {
        try {
          const rotatedDataUrl = await rotateImage(sigFile.previewUrl, assignment.rotation);
          
          if (!rotatedDataUrl || rotatedDataUrl.length === 0) {
            console.warn(`  ✗ 이미지 로테이션 실패: ${sigFile.variant} (${assignment.rotation}°)`);
            failureCount++;
            continue;
          }

          const parts = rotatedDataUrl.split(',');
          const base64Clean = parts.length > 1 ? parts[1] : parts[0];

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

          const MAX_BOX_WIDTH = 140 * assignment.scale;
          const MAX_BOX_HEIGHT = 65 * assignment.scale;

          const imgRatio = sigFile.width / sigFile.height;
          let finalWidth = MAX_BOX_WIDTH;
          let finalHeight = MAX_BOX_WIDTH / imgRatio;

          if (finalHeight > MAX_BOX_HEIGHT) {
            finalHeight = MAX_BOX_HEIGHT;
            finalWidth = MAX_BOX_HEIGHT * imgRatio;
          }

          const intWidth = Math.round(finalWidth);
          const intHeight = Math.round(finalHeight);

          const baseOffsetX = 5 + assignment.offsetX;
          const baseOffsetY = 2 + assignment.offsetY;

          const emuColOff = Math.max(0, Math.round(baseOffsetX * EMU_PER_PIXEL));
          const emuRowOff = Math.max(0, Math.round(baseOffsetY * EMU_PER_PIXEL));

          // 안전한 이미지 배치
          try {
            worksheet.addImage(imageId, {
              tl: {
                col: targetCol,
                row: targetRow,
                nativeColOff: emuColOff,
                nativeRowOff: emuRowOff
              },
              ext: { width: intWidth, height: intHeight },
              editAs: 'oneCell',
            });

            processedCount++;
            console.log(`  ✓ 배치됨: (${assignment.row},${assignment.col}) ID:${imageId}`);
          } catch (posErr) {
            console.error(`  ✗ addImage 배치 실패 (${assignment.row}, ${assignment.col}):`, posErr);
            failureCount++;
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
  console.log(`  캐시됨: ${assignmentValues.length - processedCount - failureCount - skippedCount}개`);

  // Step 3: 병합된 셀 복원
  // ExcelJS 버그 대응: 이미지를 추가한 후 병합된 셀 정보가 손실될 수 있음
  // 해결책: 원본에서 읽은 병합된 셀을 명시적으로 다시 적용
  console.log(`[병합셀 복원] 원본 병합된 셀을 다시 적용합니다...`);
  
  // 먼저 현재 병합 상태 확인
  const currentMergedCells = worksheet.model.merges ? [...worksheet.model.merges] : [];
  console.log(`[병합셀 복원] 현재 병합된 셀: ${currentMergedCells.length}개 (원본: ${originalMergedCells.length}개)`);
  
  // 병합 범위 정규화 함수 (대소문자 무시, 공백 제거, $ 기호 제거)
  const normalizeMergeRange = (range: string): string => {
    return range.toUpperCase().replace(/[\s$]/g, '');
  };
  
  // 병합된 셀이 손실되었다면 다시 적용
  if (originalMergedCells.length > 0) {
    let reappliedCount = 0;
    let errorCount = 0;
    let alreadyMergedCount = 0;
    
    // 현재 병합된 셀을 정규화하여 Set으로 저장 (빠른 조회)
    const normalizedCurrentMerges = new Set(currentMergedCells.map(normalizeMergeRange));
    
    // 실패한 병합 범위를 추적 (오류 보고용)
    const failedMerges: string[] = [];
    
    for (const mergeRange of originalMergedCells) {
      try {
        const normalizedRange = normalizeMergeRange(mergeRange);
        
        // 이미 병합되어 있는지 확인
        if (!normalizedCurrentMerges.has(normalizedRange)) {
          // 병합되지 않았으면 다시 병합
          worksheet.mergeCells(mergeRange);
          reappliedCount++;
          // 처음 몇 개만 상세 로그 출력 (성능 최적화)
          if (reappliedCount <= 5) {
            console.log(`  ✓ 병합 복원: ${mergeRange}`);
          }
        } else {
          alreadyMergedCount++;
        }
      } catch (mergeErr) {
        errorCount++;
        failedMerges.push(mergeRange);
        // 에러는 항상 출력 (중요)
        console.warn(`  ✗ 병합 실패: ${mergeRange}`, mergeErr);
      }
    }
    
    // 복원 요약 출력
    if (reappliedCount > 5) {
      console.log(`  ... 그리고 ${reappliedCount - 5}개 더 복원됨`);
    }
    
    console.log(`[병합셀 복원 완료] 복원: ${reappliedCount}개, 유지: ${alreadyMergedCount}개, 실패: ${errorCount}개`);
    
    if (failedMerges.length > 0) {
      console.warn(`[병합셀 복원 실패 목록]`, failedMerges);
    }
  } else {
    console.log(`[병합셀 복원] 원본에 병합된 셀이 없음 - 복원 불필요`);
  }
  
  // 최종 확인: 병합된 셀과 인쇄영역이 여전히 존재하는지 확인
  const finalMergedCells = worksheet.model.merges ? [...worksheet.model.merges] : [];
  const finalPrintArea = worksheet.pageSetup?.printArea;
  console.log(`[최종확인] 병합된 셀: ${finalMergedCells.length}개 (원본: ${originalMergedCells.length}개)`);
  console.log(`[최종확인] 인쇄영역: ${finalPrintArea || '설정 안 됨'} (원본: ${originalPrintArea || '설정 안 됨'})`);
  
  // 병합된 셀 수나 인쇄영역이 변경된 경우 경고
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
  
  // Step 3.5: 인쇄영역 외부의 행/열 제거 (엑셀 파일 크기 및 구조 최적화)
  if (originalPrintArea) {
    console.log(`[인쇄영역 제한] 인쇄영역 외부 데이터 정리 중...`);
    console.log(`  인쇄영역 범위: 행 ${printAreaRows.start}-${printAreaRows.end}, 열 ${printAreaCols.start}-${printAreaCols.end}`);
    
    let clearedRows = 0;
    let clearedCols = 0;
    
    // 인쇄영역 외부의 행 제거 (아래쪽)
    const currentRowCount = worksheet.actualRowCount;
    if (currentRowCount > printAreaRows.end) {
      console.log(`  현재 행 수: ${currentRowCount}, 인쇄영역 끝: ${printAreaRows.end}`);
      
      // 인쇄영역 이후의 행들을 순회하며 내용 제거
      for (let r = printAreaRows.end + 1; r <= currentRowCount; r++) {
        const row = worksheet.getRow(r);
        if (row && row.values && (row.values as any[]).some(v => v !== undefined && v !== null)) {
          // 행의 모든 셀 값 제거
          row.eachCell({ includeEmpty: true }, (cell) => {
            cell.value = null;
            cell.style = {};
          });
          clearedRows++;
        }
      }
      
      console.log(`  ✓ ${clearedRows}개 행 정리됨 (${printAreaRows.end + 1}행 이후)`);
    }
    
    // 인쇄영역 외부의 열 제거 (오른쪽)
    const currentColCount = worksheet.actualColumnCount;
    if (currentColCount > printAreaCols.end) {
      console.log(`  현재 열 수: ${currentColCount}, 인쇄영역 끝: ${printAreaCols.end}`);
      
      // 각 행에서 인쇄영역 이후의 열들을 순회하며 내용 제거
      for (let r = 1; r <= printAreaRows.end; r++) {
        const row = worksheet.getRow(r);
        for (let c = printAreaCols.end + 1; c <= currentColCount; c++) {
          const cell = row.getCell(c);
          if (cell && cell.value !== null && cell.value !== undefined) {
            cell.value = null;
            cell.style = {};
            clearedCols++;
          }
        }
      }
      
      console.log(`  ✓ ${clearedCols}개 셀 정리됨 (열 ${printAreaCols.end + 1} 이후)`);
    }
    
    console.log(`[인쇄영역 제한 완료] 행: ${clearedRows}개, 셀: ${clearedCols}개 정리됨`);
  } else {
    console.log(`[인쇄영역 제한] 인쇄영역이 설정되지 않아 전체 시트 유지`);
  }
  
  // Step 4: 워크북 저장
  try {
    console.log(`[저장중] 워크북을 버퍼로 쓰고 있습니다...`);
    
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