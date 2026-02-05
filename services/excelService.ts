import ExcelJS from 'exceljs';
import { SheetData, RowData, CellData, SignatureFile, SignatureAssignment } from '../types';

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
 * Resizing: Increased to 800px for high-quality printing.
 */
const rotateImage = async (blobUrl: string, degrees: number): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "Anonymous"; 
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(blobUrl); return; }

      // --- High Quality Logic V4 ---
      // User confirmed the root cause was infinite rows, so we can use high quality images.
      // 800px is excellent for printing (approx 6-7cm width at 300DPI).
      const MAX_WIDTH = 800; 
      let scaleFactor = 1;
      
      if (img.width > MAX_WIDTH) {
        scaleFactor = MAX_WIDTH / img.width;
      }

      const drawWidth = img.width * scaleFactor;
      const drawHeight = img.height * scaleFactor;
      
      const rad = degrees * Math.PI / 180;
      const absCos = Math.abs(Math.cos(rad));
      const absSin = Math.abs(Math.sin(rad));
      
      canvas.width = drawWidth * absCos + drawHeight * absSin;
      canvas.height = drawWidth * absSin + drawHeight * absCos;
      
      // Use High quality for crisp printing
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high'; 

      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate(rad);
      ctx.drawImage(img, -drawWidth / 2, -drawHeight / 2, drawWidth, drawHeight);
      
      const dataUrl = canvas.toDataURL('image/png');
      
      canvas.width = 1;
      canvas.height = 1;
      
      resolve(dataUrl);
    };
    img.onerror = () => {
        console.error("Image load error for rotation");
        resolve(""); 
    };
    img.src = blobUrl;
  });
};

/**
 * 업로드된 엑셀 파일 버퍼를 파싱
 * 중요: 빈 행이 연속으로 발견되면 파싱을 중단하여 메모리 누수를 방지함 (Smart Row Parsing)
 */
export const parseExcelFile = async (buffer: ArrayBuffer): Promise<SheetData> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("파일에서 워크시트를 찾을 수 없습니다.");

  const rows: RowData[] = [];
  
  // --- Infinite Row Protection ---
  // 건설/공무 양식은 10만 줄까지 서식이 적용된 경우가 많음.
  // 연속으로 50줄 이상 데이터가 없으면 문서 끝으로 간주하고 중단.
  const MAX_CONSECUTIVE_EMPTY_ROWS = 50;
  let consecutiveEmptyCount = 0;

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    // 이미 종료 조건을 만났으면 스킵
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
      consecutiveEmptyCount = 0; // 내용이 있으면 카운터 초기화
      rows.push({ index: rowNumber, cells });
    } else {
      consecutiveEmptyCount++;
      if (consecutiveEmptyCount <= MAX_CONSECUTIVE_EMPTY_ROWS) {
         rows.push({ index: rowNumber, cells });
      }
    }
  });

  return {
    name: worksheet.name,
    rows,
  };
};

/**
 * 서명 자동 매칭 로직
 */
export const autoMatchSignatures = (
  sheetData: SheetData,
  signatures: Map<string, SignatureFile[]>
): Map<string, SignatureAssignment> => {
  const assignments = new Map<string, SignatureAssignment>();
  
  let nameColIndex = -1;
  let headerRowIndex = -1;
  const MAX_HEADER_SEARCH_ROWS = 50; // 헤더 찾는 범위 약간 확장

  for (let r = 0; r < Math.min(sheetData.rows.length, MAX_HEADER_SEARCH_ROWS); r++) {
    const row = sheetData.rows[r];
    for (const cell of row.cells) {
      if (!cell.value) continue;
      const rawVal = cell.value.toString();
      // "성 명", "성명", "이 름" 등 공백 제거 후 비교
      const normalizedValue = rawVal.replace(/[\s\u00A0\uFEFF]+/g, '');
      if (/(성명|이름|Name)/i.test(normalizedValue)) {
        nameColIndex = cell.col;
        headerRowIndex = r;
        break;
      }
    }
    if (nameColIndex !== -1) break;
  }

  if (nameColIndex === -1) {
    console.warn("성명/이름 열을 자동으로 찾을 수 없습니다.");
    return assignments;
  }

  for (let r = headerRowIndex + 1; r < sheetData.rows.length; r++) {
    const row = sheetData.rows[r];
    const nameCell = row.cells.find(c => c.col === nameColIndex);
    if (!nameCell || !nameCell.value) continue;

    const rawName = nameCell.value.toString();
    const cleanName = normalizeName(rawName);

    if (!cleanName) continue;

    const availableSigs = signatures.get(cleanName);
    
    if (availableSigs && availableSigs.length > 0) {
      for (const cell of row.cells) {
        if (cell.col === nameColIndex) continue;
        if (!cell.value) continue;
        const cellStr = cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '');
        
        // 1, (1), 1. 등 서명 마킹 확인
        if (['1', '(1)', '1.', '1)'].includes(cellStr)) {
          const key = `${cell.row}:${cell.col}`;
          
          const randomSigIndex = Math.floor(Math.random() * availableSigs.length);
          const selectedSig = availableSigs[randomSigIndex];
          
          // Integer rotation (-8 to +8) to maximize cache hits
          const rotation = Math.floor(Math.random() * 17) - 8; 
          
          const scale = 1.0 + (Math.random() * 0.3); // 1.0 ~ 1.3
          const offsetX = Math.floor(Math.random() * 5) - 2; // -2 ~ +2 integer
          const offsetY = 0; 

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
        }
      }
    }
  }

  return assignments;
};

/**
 * 최종 엑셀 생성 (Batch Processing applied)
 */
export const generateFinalExcel = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>
): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalBuffer);
  const worksheet = workbook.worksheets[0];

  const imageIdMap = new Map<string, number>();

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  const assignmentValues = Array.from(assignments.values());
  
  // --- BATCH PROCESSING LOOP ---
  // Process items in small chunks to prevent UI freeze and allow potential GC
  const CHUNK_SIZE = 20; 

  for (let i = 0; i < assignmentValues.length; i++) {
    // Every CHUNK_SIZE items, pause for a moment
    if (i % CHUNK_SIZE === 0) {
      await new Promise(resolve => setTimeout(resolve, 0));
    }

    const assignment = assignmentValues[i];
    const sigFile = findSigFile(assignment.signatureBaseName, assignment.signatureVariantId);
    if (!sigFile) continue;

    const cacheKey = `${sigFile.variant}_${assignment.rotation}`;
    
    let imageId = imageIdMap.get(cacheKey);

    if (imageId === undefined) {
      const rotatedDataUrl = await rotateImage(sigFile.previewUrl, assignment.rotation);
      
      if (rotatedDataUrl) {
          imageId = workbook.addImage({
            base64: rotatedDataUrl,
            extension: 'png',
          });
          imageIdMap.set(cacheKey, imageId);
      }
    }

    if (imageId !== undefined) {
        const targetCol = assignment.col - 1;
        const targetRow = assignment.row - 1;

        const baseHeight = 20; 
        const baseWidth = 50; 
        
        const finalWidth = baseWidth * assignment.scale; 
        const finalHeight = baseHeight * assignment.scale;

        let colOffset = 0.1 + (assignment.offsetX / 100);
        let rowOffset = 0.1; 

        colOffset = Math.max(0.05, Math.min(0.95, colOffset));
        rowOffset = Math.max(0.05, Math.min(0.5, rowOffset)); 

        worksheet.addImage(imageId, {
          tl: { 
            col: targetCol + colOffset, 
            row: targetRow + rowOffset 
          },
          ext: { width: finalWidth, height: finalHeight },
          editAs: 'oneCell',
        });
        
        // Remove text placeholder
        try {
          const cell = worksheet.getCell(assignment.row, assignment.col);
          const cellVal = cell.value ? cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '') : '';
          if (['1', '(1)', '1.', '1)'].includes(cellVal)) {
             cell.value = '';
          }
        } catch (e) {
          // Ignore
        }
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};