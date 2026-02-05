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
 * 이미지 회전 및 최적화 헬퍼 함수 (Smart Resizing & Rotation)
 * 1. 이미지를 회전시킵니다.
 * 2. 인쇄 품질(300DPI)을 유지할 수 있는 최적의 크기(Max Width 600px)로 리사이징하여 메모리를 절약합니다.
 */
const rotateImage = async (dataUrl: string, degrees: number): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(dataUrl); return; }

      // --- Smart Resizing Logic ---
      // 서명란(보통 3~5cm) 기준 300DPI 인쇄 시 약 400~600px이면 충분히 선명함.
      // 원본이 4000px인 사진을 그대로 쓰면 메모리 폭발함. 이를 방지.
      const MAX_WIDTH = 600; 
      let scaleFactor = 1;
      
      // 원본이 너무 크면 축소, 작으면 유지 (확대하지 않음)
      if (img.width > MAX_WIDTH) {
        scaleFactor = MAX_WIDTH / img.width;
      }

      const drawWidth = img.width * scaleFactor;
      const drawHeight = img.height * scaleFactor;
      
      const rad = degrees * Math.PI / 180;
      const absCos = Math.abs(Math.cos(rad));
      const absSin = Math.abs(Math.sin(rad));
      
      // 회전 후 캔버스 크기 계산 (리사이징된 크기 기준)
      canvas.width = drawWidth * absCos + drawHeight * absSin;
      canvas.height = drawWidth * absSin + drawHeight * absCos;
      
      // 렌더링 품질 설정 (High Quality)
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';

      // 중심점 이동 후 회전 및 그리기
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate(rad);
      ctx.drawImage(img, -drawWidth / 2, -drawHeight / 2, drawWidth, drawHeight);
      
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
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

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const cells: CellData[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const stringValue = getCellValueAsString(cell);
      cells.push({
        value: stringValue,
        address: cell.address,
        row: rowNumber,
        col: colNumber,
      });
    });
    rows.push({ index: rowNumber, cells });
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
  const MAX_HEADER_SEARCH_ROWS = 30;

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
          
          // --- Memory Optimization: Integer Rotation ---
          // 회전 각도를 정수(-8, -7 ... 7, 8)로 제한합니다.
          // 이는 generateFinalExcel 단계에서 이미지 캐싱 적중률(Cache Hit Rate)을 높여
          // 파일 용량을 획기적으로 줄이는 핵심 기술입니다.
          const rotation = Math.floor(Math.random() * 17) - 8; // -8 ~ +8 integer
          
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
 * 최종 엑셀 생성 (고속 캐싱 적용)
 */
export const generateFinalExcel = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>
): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalBuffer);
  const worksheet = workbook.worksheets[0];

  // --- Image Cache Optimization ---
  // Key: "SignatureID_RotationAngle" (예: "Hong_1_-5")
  // Value: ExcelJS Image ID
  // 회전 각도가 정수로 제한되어 있으므로, 동일한 서명이 여러 번 쓰일 때 
  // 이미지를 새로 만들지 않고 기존 ID를 재사용합니다. (파일 용량 대폭 감소)
  const imageIdMap = new Map<string, number>();

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  for (const assignment of assignments.values()) {
    const sigFile = findSigFile(assignment.signatureBaseName, assignment.signatureVariantId);
    if (!sigFile) continue;

    // 캐시 키 생성 (정수 회전값 사용으로 캐시 효율 극대화)
    const cacheKey = `${sigFile.variant}_${assignment.rotation}`;
    
    let imageId = imageIdMap.get(cacheKey);

    if (imageId === undefined) {
      // 캐시 미스: 이미지 처리 수행 (비용이 큼)
      // 여기서 rotateImage는 Smart Resizing을 수행하여 600px 이하로 최적화된 이미지를 반환함
      const rotatedDataUrl = await rotateImage(sigFile.dataUrl, assignment.rotation);
      
      imageId = workbook.addImage({
        base64: rotatedDataUrl,
        extension: 'png',
      });
      
      // 캐시에 저장
      imageIdMap.set(cacheKey, imageId);
    }

    // 좌표계 조정
    const targetCol = assignment.col - 1;
    const targetRow = assignment.row - 1;

    // 배치 크기 계산 (Base 20px -> Scale 적용)
    const baseHeight = 20; 
    const baseWidth = 50; 
    
    const finalWidth = baseWidth * assignment.scale; 
    const finalHeight = baseHeight * assignment.scale;

    // 배치 오프셋
    let colOffset = 0.1 + (assignment.offsetX / 100);
    let rowOffset = 0.1; 

    // Safe clamping
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
    
    // 원본 텍스트 제거
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

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};