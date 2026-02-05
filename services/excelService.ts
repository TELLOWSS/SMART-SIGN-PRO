import ExcelJS from 'exceljs';
import { SheetData, RowData, CellData, SignatureFile, SignatureAssignment } from '../types';

/**
 * 매칭을 위해 이름 정규화
 * 1. 괄호 및 괄호 안의 내용 제거 (예: "홍길동 (주)", "John (Manager)" -> "홍길동", "John")
 * 2. 한글, 영문, 숫자 이외의 특수문자 및 공백 제거 (예: "Hong-Gil-Dong" -> "HongGilDong")
 * 3. 영문 대소문자 통일 (소문자)
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
 * RichText, 수식(Formula), 숫자, 문자열 등 다양한 타입을 처리합니다.
 */
const getCellValueAsString = (cell: ExcelJS.Cell | undefined): string => {
  if (!cell || cell.value === null || cell.value === undefined) return '';

  const val = cell.value;

  // 1. 객체 타입 (RichText, Hyperlink, Formula 등)
  if (typeof val === 'object') {
    // 수식(Formula)의 경우 계산된 결과값(result)을 사용
    if ('result' in val) {
      return val.result !== undefined ? val.result.toString() : '';
    }
    // RichText (스타일이 적용된 텍스트)
    if ('richText' in val && Array.isArray((val as any).richText)) {
      return (val as any).richText.map((rt: any) => rt.text).join('');
    }
    // Hyperlink
    if ('text' in val) {
      return (val as any).text.toString();
    }
    // 기타 객체는 JSON 문자열 또는 toString 시도
    return val.toString();
  }

  // 2. 기본 타입 (string, number)
  return val.toString();
};

/**
 * 이미지 회전 헬퍼 함수
 * 캔버스를 사용하여 이미지를 회전시키고 새로운 Base64 문자열을 반환합니다.
 */
const rotateImage = async (dataUrl: string, degrees: number): Promise<string> => {
  if (Math.abs(degrees) < 0.1) return dataUrl; // 회전이 거의 없으면 원본 반환

  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(dataUrl); return; }
      
      const rad = degrees * Math.PI / 180;
      // 회전 시 잘림 방지를 위해 캔버스 크기 재계산
      const absCos = Math.abs(Math.cos(rad));
      const absSin = Math.abs(Math.sin(rad));
      canvas.width = img.width * absCos + img.height * absSin;
      canvas.height = img.width * absSin + img.height * absCos;
      
      // 중심점 이동 후 회전
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate(rad);
      ctx.drawImage(img, -img.width / 2, -img.height / 2);
      
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => resolve(dataUrl); // 에러 시 원본 반환
    img.src = dataUrl;
  });
};

/**
 * 업로드된 엑셀 파일 버퍼를 파싱하여 미리보기 UI용 데이터를 추출합니다.
 */
export const parseExcelFile = async (buffer: ArrayBuffer): Promise<SheetData> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  // 첫 번째 시트를 대상으로 가정
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("파일에서 워크시트를 찾을 수 없습니다.");

  const rows: RowData[] = [];

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const cells: CellData[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      // 미리보기용 데이터 추출 시에도 강력한 변환기 사용
      const stringValue = getCellValueAsString(cell);
      
      cells.push({
        value: stringValue, // UI 표시용 문자열
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
 * '1' 표시가 된 셀을 찾아 서명과 자동으로 매칭하는 휴리스틱 로직입니다.
 */
export const autoMatchSignatures = (
  sheetData: SheetData,
  signatures: Map<string, SignatureFile[]>
): Map<string, SignatureAssignment> => {
  const assignments = new Map<string, SignatureAssignment>();
  
  // 1. "이름/성명" 열 인덱스 찾기
  let nameColIndex = -1;
  let headerRowIndex = -1;

  // 헤더 탐색 범위 확장 (상단 30행) - 헤더가 복잡할 경우를 대비
  const MAX_HEADER_SEARCH_ROWS = 30;

  for (let r = 0; r < Math.min(sheetData.rows.length, MAX_HEADER_SEARCH_ROWS); r++) {
    const row = sheetData.rows[r];
    for (const cell of row.cells) {
      if (!cell.value) continue;
      
      // 공백 및 특수문자 제거 후 비교
      const rawVal = cell.value.toString();
      const normalizedValue = rawVal.replace(/[\s\u00A0\uFEFF]+/g, '');
      
      // 정규식으로 '성명', '이름', 'Name' 패턴 확인
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

  // 2. 헤더 이후의 행 순회
  for (let r = headerRowIndex + 1; r < sheetData.rows.length; r++) {
    const row = sheetData.rows[r];
    
    // 이 행의 이름 찾기
    const nameCell = row.cells.find(c => c.col === nameColIndex);
    if (!nameCell || !nameCell.value) continue;

    const rawName = nameCell.value.toString();
    const cleanName = normalizeName(rawName);

    // 이름이 비어있으면 스킵
    if (!cleanName) continue;

    // 이 이름에 해당하는 서명 파일이 있는지 확인
    const availableSigs = signatures.get(cleanName);
    
    // 서명이 있고, 행 내에 '1' 표시가 있다면 매칭
    if (availableSigs && availableSigs.length > 0) {
      for (const cell of row.cells) {
        // 이름 열 자체는 건너뜀
        if (cell.col === nameColIndex) continue;
        if (!cell.value) continue;

        // "1" 마커 확인
        const cellStr = cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '');
        
        // 유연한 매칭 (괄호, 점 등 허용)
        if (['1', '(1)', '1.', '1)'].includes(cellStr)) {
          const key = `${cell.row}:${cell.col}`;
          
          // --- 서명 랜덤 선택 로직 (Random Selection Logic) ---
          const randomSigIndex = Math.floor(Math.random() * availableSigs.length);
          const selectedSig = availableSigs[randomSigIndex];
          
          // --- 분석 및 개선: Safe Print Mode V2 ---
          // 분석 결과: 이전의 1.1~1.4 배율과 24px 높이는 일반적인 엑셀 행 높이(약 22px)를 초과하여 하단이 잘릴 위험이 큼.
          // 개선: 
          // 1. 회전 각도를 -8도~+8도로 더욱 제한하여 모서리 튀어나옴 방지
          const rotation = (Math.random() * 16) - 8;
          
          // 2. 크기 배율을 1.0~1.3으로 축소 (베이스 높이도 20px로 줄일 예정)
          const scale = 1.0 + (Math.random() * 0.3);

          // 3. 위치 오프셋: 수평은 허용하되, 수직은 0으로 완전 고정
          const offsetX = (Math.random() * 4) - 2;
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
 * 이미지가 포함된 최종 엑셀 파일을 생성합니다.
 */
export const generateFinalExcel = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>
): Promise<Blob> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalBuffer);
  const worksheet = workbook.worksheets[0];

  // 1. 이미지 ID 캐시
  const imageIdMap = new Map<string, number>();

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  // 2. 할당된 정보(assignments)를 순회하며 이미지 배치
  for (const assignment of assignments.values()) {
    const sigFile = findSigFile(assignment.signatureBaseName, assignment.signatureVariantId);
    if (!sigFile) continue;

    // 이미지 회전 처리
    const rotatedDataUrl = await rotateImage(sigFile.dataUrl, assignment.rotation);
    const cacheKey = `${sigFile.variant}_${assignment.rotation.toFixed(2)}`;

    let imageId = imageIdMap.get(cacheKey);
    if (imageId === undefined) {
      imageId = workbook.addImage({
        base64: rotatedDataUrl,
        extension: 'png',
      });
      imageIdMap.set(cacheKey, imageId);
    }

    // 좌표계 조정 (1-based -> 0-based)
    const targetCol = assignment.col - 1;
    const targetRow = assignment.row - 1;

    // --- 분석 기반 개선: 크기 및 배치 최적화 ---
    // 기존 24px -> 20px로 축소 (표준 엑셀 행 높이 16.5pt ~= 22px 내에 안전하게 안착 유도)
    // Max Height at 1.3 scale = 26px. 
    // 여전히 표준 행보다 약간 클 수 있으나, 보통 서명란은 행 높이를 키우는 경우가 많으므로 적절한 타협점.
    const baseHeight = 20; 
    const baseWidth = 50; // 비율 유지하며 너비도 소폭 축소
    
    const finalWidth = baseWidth * assignment.scale; 
    const finalHeight = baseHeight * assignment.scale;

    // 배치 로직: 상단 고정 오프셋
    // 0.15 (15%) 오프셋은 22px 행 기준 약 3.3px 띄움. 
    // 3.3px(top) + 26px(height) = 29.3px. 
    // 표준 행(22px)이라면 여전히 7px 넘침. 
    // 하지만 완전히 작게 만들면 가시성이 떨어짐.
    // --> rowOffset을 0.1(10%)로 줄여 최대한 위로 붙임.
    let colOffset = 0.1 + (assignment.offsetX / 100);
    let rowOffset = 0.1; 

    // Safe clamping
    colOffset = Math.max(0.05, Math.min(0.95, colOffset));
    // rowOffset은 고정이지만 안전장치 유지
    rowOffset = Math.max(0.05, Math.min(0.5, rowOffset)); 

    worksheet.addImage(imageId, {
      tl: { 
        col: targetCol + colOffset, 
        row: targetRow + rowOffset 
      },
      ext: { width: finalWidth, height: finalHeight },
      editAs: 'oneCell',
    });
    
    // 원본 셀의 '1' 텍스트 제거
    try {
      const cell = worksheet.getCell(assignment.row, assignment.col);
      const cellVal = cell.value ? cell.value.toString().replace(/[\s\u00A0\uFEFF]+/g, '') : '';
      if (['1', '(1)', '1.', '1)'].includes(cellVal)) {
         cell.value = '';
      }
    } catch (e) {
      // Ignore cell access error
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
};