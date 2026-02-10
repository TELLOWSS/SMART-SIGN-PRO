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
 * 최종 엑셀 생성 - 원본 파일 구조 완벽 보존
 * 
 * 전략변경:
 * 1. 원본 파일을 직접 로드하고 수정
 * 2. 병합된 셀 완벽 유지
 * 3. 인쇄영역 설정 유지
 * 4. 이미지와 텍스트만 추가/수정
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

  // 병합된 셀 정보 저장 (보존용)
  const originalMergedCells = worksheet.merged ? [...worksheet.merged] : [];
  console.log(`[병합셀] 원본 병합된 셀: ${originalMergedCells.length}개`);
  for (const merge of originalMergedCells) {
    console.log(`  - ${merge}`);
  }

  // 인쇄영역 정보 저장
  const originalPrintArea = worksheet.pageSetup?.printArea;
  console.log(`[인쇄영역] 원본 인쇄영역: ${originalPrintArea || '설정 안 됨'}`);

  // 여러 워크시트 문제 체크
  if (workbook.worksheets.length > 1) {
    console.warn(`⚠️ 경고: 원본 파일에 ${workbook.worksheets.length}개 시트가 있습니다.`);
  }

  // 인쇄영역 범위 파싱
  let printAreaRows = { start: 1, end: worksheet.rowCount };
  let printAreaCols = { start: 1, end: worksheet.columnCount };
  
  if (originalPrintArea) {
    try {
      // printArea 형식: "A1:C10" 또는 "Sheet1!A1:C10"
      const range = originalPrintArea.split('!').pop() || originalPrintArea;
      const [topLeft, bottomRight] = range.split(':');
      if (topLeft && bottomRight) {
        const tlMatch = topLeft.match(/([A-Z]+)(\d+)/);
        const brMatch = bottomRight.match(/([A-Z]+)(\d+)/);
        if (tlMatch && brMatch) {
          const colToNum = (col: string) => col.charCodeAt(0) - 64; // 'A'=1, 'B'=2 등
          printAreaRows = { start: parseInt(tlMatch[2]), end: parseInt(brMatch[2]) };
          printAreaCols = { start: colToNum(tlMatch[1]), end: colToNum(brMatch[1]) };
          console.log(`[인쇄영역파싱] 행: ${printAreaRows.start}-${printAreaRows.end}, 열: ${printAreaCols.start}-${printAreaCols.end}`);
        }
      }
    } catch (parseErr) {
      console.warn(`[인쇄영역파싱실패]`, parseErr);
    }
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

      const cell = worksheet.getCell(row, col);
      if (cell) {
        const cellVal = cell.value ? cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '') : '';
        
        if (['1', '(1)', '1.', '1)', 'o', 'o)', '○'].includes(cellVal)) {
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

  // Step 3: 병합된 셀 및 인쇄영역 복원
  console.log(`[복원] 병합된 셀 및 인쇄영역 복원 중...`);
  
  // 기존 병합 셀 제거 (ExcelJS 재저장 중 손상 방지)
  if (worksheet.merged && worksheet.merged.length > 0) {
    console.log(`  [제거] 현재 병합 셀 ${worksheet.merged.length}개 제거`);
    const mergedCopy = [...worksheet.merged];
    for (const merge of mergedCopy) {
      try {
        worksheet.unmerge(merge);
      } catch (e) {
        console.warn(`  ⚠️ Unmerge 실패: ${merge}`, e);
      }
    }
  }

  // 원본 병합 셀 복원
  for (const merge of originalMergedCells) {
    try {
      worksheet.merge(merge);
      console.log(`  ✓ 병합 복원: ${merge}`);
    } catch (e) {
      console.warn(`  ✗ 병합 복원 실패: ${merge}`, e);
    }
  }

  // 인쇄영역 복원
  if (originalPrintArea) {
    try {
      if (!worksheet.pageSetup) {
        worksheet.pageSetup = {};
      }
      worksheet.pageSetup.printArea = originalPrintArea;
      console.log(`  ✓ 인쇄영역 복원: ${originalPrintArea}`);
    } catch (e) {
      console.warn(`  ✗ 인쇄영역 복원 실패:`, e);
    }
  }

  // Step 4: 워크북 저장
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