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

  return {
    name: worksheet.name || 'Sheet1',
    rows,
  };
};

/**
 * 서명 자동 매칭 로직
 * 개선사항: 더 나은 열 검색, 에러 처리, 결과 보고
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
        if (['1', '(1)', '1.', '1)', 'o', 'o)', '○'].includes(cellStr)) {
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
 * 최종 엑셀 생성
 * 
 * 주요 개선:
 * 1. 새 워크북 생성 → 안전한 구조
 * 2. 원본 데이터만 읽음 (수정 안함)
 * 3. 새 시트에 데이터 복사
 * 4. 이미지 추가
 * 5. 셀 텍스트 처리
 */
export const generateFinalExcel = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>
): Promise<Blob> => {
  if (!originalBuffer || originalBuffer.byteLength === 0) {
    throw new Error("원본 파일 버퍼가 비어있습니다.");
  }

  // Step 1: 원본 파일에서 데이터 읽기 (읽기 전용)
  const sourceWorkbook = new ExcelJS.Workbook();
  sourceWorkbook.xlsx.date1904 = false;
  
  try {
    await sourceWorkbook.xlsx.load(originalBuffer);
  } catch (err) {
    throw new Error(`원본 파일 읽음 불가: ${err instanceof Error ? err.message : '알 수 없음'}`);
  }

  const sourceWorksheet = sourceWorkbook.worksheets[0];
  if (!sourceWorksheet) {
    throw new Error("워크시트를 찾을 수 없습니다.");
  }

  console.log(`[원본 로드] 행: ${sourceWorksheet.actualRowCount}, 열: ${sourceWorksheet.actualColumnCount}`);

  // Step 2: 새로운 클린 워크북 생성
  const newWorkbook = new ExcelJS.Workbook();
  newWorkbook.xlsx.date1904 = false;
  
  if (newWorkbook.worksheets.length > 0) {
    newWorkbook.removeWorksheet(newWorkbook.worksheets[0]);
  }
  
  const newWorksheet = newWorkbook.addWorksheet(sourceWorksheet.name || 'Sheet1');

  // Step 3: 원본 셀 데이터 안전하게 복사
  const cellsToModify = new Set<string>();
  
  sourceWorksheet.eachRow({ includeEmpty: true }, (sourceRow, rowNum) => {
    const newRow = newWorksheet.getRow(rowNum);
    
    // 행 높이 복사
    if (sourceRow.height) {
      newRow.height = sourceRow.height;
    }

    sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNum) => {
      const newCell = newRow.getCell(colNum);

      // 셀 값 복사 (타입 유지)
      if (sourceCell.value !== null && sourceCell.value !== undefined) {
        try {
          // 값의 타입 확인 후 복사
          if (typeof sourceCell.value === 'object' && 'text' in sourceCell.value) {
            // RichText 타입
            newCell.value = sourceCell.value;
          } else if (typeof sourceCell.value === 'number' || typeof sourceCell.value === 'string' || typeof sourceCell.value === 'boolean') {
            newCell.value = sourceCell.value;
          } else if (sourceCell.value instanceof Date) {
            newCell.value = sourceCell.value;
          } else {
            // 객체 타입은 문자열로 변환
            newCell.value = sourceCell.value.toString();
          }
          cellsToModify.add(`${rowNum}:${colNum}`);
        } catch (e) {
          console.warn(`Failed to copy cell value at ${sourceCell.address}`, e);
        }
      }

      // 셀 스타일 깊은 복사 (null 체크 강화)
      if (sourceCell.style && typeof sourceCell.style === 'object') {
        try {
          const styleCopy = JSON.parse(JSON.stringify(sourceCell.style));
          newCell.style = styleCopy;
        } catch (e) {
          console.warn(`Failed to copy cell style at ${sourceCell.address}`, e);
          // 스타일이 없어도 계속 진행
        }
      }

      // 셀 서식 복사 (number format 등)
      if (sourceCell.numFmt) {
        try {
          newCell.numFmt = sourceCell.numFmt;
        } catch (e) {
          console.warn(`Failed to copy cell format at ${sourceCell.address}`, e);
        }
      }
    });
  });

  // Step 4: 열 너비 복사
  if (sourceWorksheet.columns && Array.isArray(sourceWorksheet.columns)) {
    sourceWorksheet.columns.forEach((col, idx) => {
      if (col && col.width && idx < newWorksheet.columns.length) {
        try {
          const newCol = newWorksheet.columns[idx];
          if (newCol) {
            newCol.width = col.width;
          }
        } catch (e) {
          console.warn(`Failed to set column width at index ${idx}`, e);
        }
      }
    });
  }

  // Step 5: 병합된 셀 정보 복사 (안전하게)
  try {
    if (sourceWorksheet._mergedCells && sourceWorksheet._mergedCells.length > 0) {
      sourceWorksheet._mergedCells.forEach((merged: string) => {
        try {
          newWorksheet.mergeCells(merged);
        } catch (e) {
          console.warn(`Failed to merge cells: ${merged}`, e);
        }
      });
    }
  } catch (e) {
    console.warn("Failed to copy merged cells information", e);
  }

  // Step 6: 할당된 서명 정보 처리 (이미지 추가 전에 셀 텍스트 먼저 수정)
  for (const [key, assignment] of assignments) {
    try {
      const [rowStr, colStr] = key.split(':');
      const row = parseInt(rowStr, 10);
      const col = parseInt(colStr, 10);

      const cell = newWorksheet.getCell(row, col);
      if (cell) {
        const cellVal = cell.value ? cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '') : '';
        
        if (['1', '(1)', '1.', '1)', 'o', 'o)', '○'].includes(cellVal)) {
          cell.value = null;
        }
      }
    } catch (e) {
      console.warn(`Failed to clear cell at ${key}`, e);
    }
  }

  // Step 7: 이미지 추가 (셀 수정 완료 후)
  const imageCache = new Map<string, number>();
  const EMU_PER_PIXEL = 9525;
  let processedCount = 0;
  let failureCount = 0;

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  const assignmentValues = Array.from(assignments.values());
  const CHUNK_SIZE = 25;

  for (let i = 0; i < assignmentValues.length; i++) {
    if (i % CHUNK_SIZE === 0) {
      await new Promise(resolve => setTimeout(resolve, 0));
    }

    try {
      const assignment = assignmentValues[i];
      const sigFile = findSigFile(assignment.signatureBaseName, assignment.signatureVariantId);
      
      if (!sigFile) {
        console.warn(`Signature file not found: ${assignment.signatureVariantId}`);
        failureCount++;
        continue;
      }

      const cacheKey = `${sigFile.variant}_rot${assignment.rotation}`;
      let imageId = imageCache.get(cacheKey);

      if (imageId === undefined) {
        try {
          const rotatedDataUrl = await rotateImage(sigFile.previewUrl, assignment.rotation);
          
          if (rotatedDataUrl && rotatedDataUrl.length > 0) {
            const parts = rotatedDataUrl.split(',');
            const base64Clean = parts.length > 1 ? parts[1] : parts[0];

            if (base64Clean && base64Clean.length > 0) {
              imageId = newWorkbook.addImage({
                base64: base64Clean,
                extension: 'png',
              });
              imageCache.set(cacheKey, imageId);
            }
          }
        } catch (imgErr) {
          console.warn(`Failed to process image: ${sigFile.variant}`, imgErr);
          failureCount++;
          continue;
        }
      }

      if (imageId !== undefined) {
        try {
          const targetCol = assignment.col - 1;  // 0-based for ExcelJS
          const targetRow = assignment.row - 1;

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

          newWorksheet.addImage(imageId, {
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
        } catch (posErr) {
          console.warn(`Failed to position image at ${assignment.row}:${assignment.col}`, posErr);
          failureCount++;
        }
      }
    } catch (assignErr) {
      console.error("Error processing assignment", assignErr);
      failureCount++;
    }
  }

  console.log(`Excel generation complete: ${processedCount} signatures added, ${failureCount} failed`);

  // Step 8-1: 중간 저장 (데이터 구조 안정화)
  let intermediateBuffer: Buffer;
  try {
    intermediateBuffer = await newWorkbook.xlsx.writeBuffer() as any as Buffer;
    
    if (!intermediateBuffer || intermediateBuffer.byteLength < 50) {
      throw new Error("중간 파일 생성 실패");
    }
    
    console.log(`Intermediate file created: ${intermediateBuffer.byteLength} bytes`);
  } catch (intermediateErr) {
    throw new Error(`중간 파일 작성 중 오류: ${intermediateErr instanceof Error ? intermediateErr.message : '알 수 없는 오류'}`);
  }

  // Step 8-2: 최종 검증 및 반환
  try {
    const finalBlob = new Blob([intermediateBuffer], { 
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    if (!finalBlob || finalBlob.size === 0) {
      throw new Error("최종 Blob 생성 실패");
    }

    if (finalBlob.size < 50) {
      throw new Error(`파일 크기 이상: ${finalBlob.size} bytes (손상된 파일일 수 있음)`);
    }

    console.log(`✅ Successfully generated Excel file: ${finalBlob.size} bytes`);
    return finalBlob;
  } catch (finalErr) {
    throw new Error(`최종 파일 생성 실패: ${finalErr instanceof Error ? finalErr.message : '알 수 없는 오류'}`);
  } finally {
    // 메모리 정리
    try {
      sourceWorkbook.worksheets.forEach(ws => {
        try { sourceWorkbook.removeWorksheet(ws); } catch (e) { /* ignore */ }
      });
    } catch (e) {
      console.warn("Memory cleanup warning:", e);
    }
  }
};