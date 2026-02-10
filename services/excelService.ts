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
 * Output: Data URL (Base64)
 */
const rotateImage = async (blobUrl: string, degrees: number): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "Anonymous"; 
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(blobUrl); return; }

      // --- High Quality Logic ---
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
 */
export const parseExcelFile = async (buffer: ArrayBuffer): Promise<SheetData> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("파일에서 워크시트를 찾을 수 없습니다.");

  const rows: RowData[] = [];
  
  // --- Infinite Row Protection ---
  const MAX_CONSECUTIVE_EMPTY_ROWS = 50;
  let consecutiveEmptyCount = 0;

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
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
  const MAX_HEADER_SEARCH_ROWS = 50;

  for (let r = 0; r < Math.min(sheetData.rows.length, MAX_HEADER_SEARCH_ROWS); r++) {
    const row = sheetData.rows[r];
    for (const cell of row.cells) {
      if (!cell.value) continue;
      const rawVal = cell.value.toString();
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
        
        if (['1', '(1)', '1.', '1)'].includes(cellStr)) {
          const key = `${cell.row}:${cell.col}`;
          
          const randomSigIndex = Math.floor(Math.random() * availableSigs.length);
          const selectedSig = availableSigs[randomSigIndex];
          
          // Random offset calculations for X/Y
          // Limit to small values to prevent layout breakage
          const rotation = Math.floor(Math.random() * 17) - 8; 
          const scale = 1.0 + (Math.random() * 0.3); // 1.0 ~ 1.3
          const offsetX = Math.floor(Math.random() * 10) - 5; // -5px ~ +5px
          const offsetY = Math.floor(Math.random() * 6) - 3;  // -3px ~ +3px

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
 * 최종 엑셀 생성
 * 
 * Major Fixes for "File Corrupted" Error:
 * 1. Coordinates: Uses standard Integer Column/Row + EMU Offsets (English Metric Units).
 *    - 1 px is approx 9525 EMUs.
 *    - This avoids floating point errors in XML generation which causes Excel to repair the file.
 * 2. Dimensions: Rounds image dimensions to Integers.
 * 3. Cell Safety: Checks if a cell is merged before clearing content.
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
  const EMU_PER_PIXEL = 9525;

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  const assignmentValues = Array.from(assignments.values());
  const CHUNK_SIZE = 20; 

  for (let i = 0; i < assignmentValues.length; i++) {
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
          const parts = rotatedDataUrl.split(',');
          const base64Clean = parts.length > 1 ? parts[1] : parts[0];

          if (base64Clean) {
            imageId = workbook.addImage({
              base64: base64Clean,
              extension: 'png',
            });
            imageIdMap.set(cacheKey, imageId);
          }
      }
    }

    if (imageId !== undefined) {
        const targetCol = assignment.col - 1; // 0-based index
        const targetRow = assignment.row - 1; // 0-based index

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

        // Center the image in the cell roughly + random offset
        // We assume a standard cell padding.
        // X Offset: 5px padding + random
        // Y Offset: 2px padding + random
        const baseOffsetX = 5 + assignment.offsetX;
        const baseOffsetY = 2 + assignment.offsetY;

        // Convert to EMUs for strict OpenXML compliance
        // Ensure non-negative to avoid XML validation errors
        const emuColOff = Math.max(0, Math.round(baseOffsetX * EMU_PER_PIXEL));
        const emuRowOff = Math.max(0, Math.round(baseOffsetY * EMU_PER_PIXEL));

        // Use nativeColOff/nativeRowOff with integer col/row
        // This is the most stable method for ExcelJS images
        worksheet.addImage(imageId, {
          tl: { 
            col: targetCol, 
            row: targetRow,
            nativeColOff: emuColOff, 
            nativeRowOff: emuRowOff 
          },
          ext: { width: intWidth, height: intHeight },
          editAs: 'oneCell', // Moves with cells
        });
        
        // Remove text placeholder safely
        try {
          const cell = worksheet.getCell(assignment.row, assignment.col);
          
          // Only clear if it's the specific placeholder text
          // And check if it's not part of a weird merge that shouldn't be touched (though value null is usually safe)
          // We can check cell.isMerged, but sometimes '1' is in a merged cell. 
          // If it is merged, ExcelJS shares the value. Setting master cell value is correct.
          // If this is a slave cell in a merge, getCell returns the master usually? No, it returns the specific cell.
          // We should find the master if merged.
          
          const master = cell.master; // If merged, this is the top-left cell
          const cellToEdit = master || cell;

          const cellVal = cellToEdit.value ? cellToEdit.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '') : '';
          
          if (['1', '(1)', '1.', '1)'].includes(cellVal)) {
             cellToEdit.value = null; 
          }
        } catch (e) {
          console.warn("Error clearing cell value", e);
        }
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};