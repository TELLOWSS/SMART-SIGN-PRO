import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import ExcelJS from 'exceljs';
import { SignatureAssignment, SignatureFile } from '../types';
import { columnLetterToNumber, isSignaturePlaceholder } from './excelUtils';

/**
 * 엑셀 시트를 HTML 테이블로 렌더링하여 이미지로 변환하는 헬퍼 함수
 */
const renderSheetToCanvas = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>,
  printAreaOnly: boolean = true
): Promise<HTMLCanvasElement> => {
  // 워크북 로드
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalBuffer);
  
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("워크시트를 찾을 수 없습니다.");
  }

  // 인쇄영역 파싱
  const originalPrintArea = worksheet.pageSetup?.printArea;
  let printAreaRows = { start: 1, end: worksheet.actualRowCount || 100 };
  let printAreaCols = { start: 1, end: worksheet.actualColumnCount || 26 };

  if (printAreaOnly && originalPrintArea) {
    // 인쇄영역 파싱 로직 (excelService.ts와 동일)
    try {
      let range = originalPrintArea;
      if (range.includes('!')) {
        range = range.split('!').pop() || range;
      }
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
          
          printAreaRows = { start: tlRow, end: brRow };
          printAreaCols = { start: tlCol, end: brCol };
        }
      }
    } catch (err) {
      console.warn('인쇄영역 파싱 실패, 전체 시트 사용', err);
    }
  }

  console.log(`[렌더링] 행: ${printAreaRows.start}-${printAreaRows.end}, 열: ${printAreaCols.start}-${printAreaCols.end}`);

  // HTML 테이블 생성
  const container = document.createElement('div');
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.top = '0';
  container.style.background = 'white';
  container.style.padding = '20px';
  document.body.appendChild(container);

  const table = document.createElement('table');
  table.style.borderCollapse = 'collapse';
  table.style.fontFamily = 'Arial, sans-serif';
  table.style.fontSize = '12px';
  table.style.background = 'white';

  // 서명 파일 찾기 헬퍼
  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  // 행 렌더링
  for (let r = printAreaRows.start; r <= printAreaRows.end; r++) {
    const row = worksheet.getRow(r);
    const tr = document.createElement('tr');

    for (let c = printAreaCols.start; c <= printAreaCols.end; c++) {
      const cell = row.getCell(c);
      const td = document.createElement('td');
      
      // 기본 스타일
      td.style.border = '1px solid #ddd';
      td.style.padding = '8px';
      td.style.minWidth = '80px';
      td.style.minHeight = '40px';
      td.style.position = 'relative';
      td.style.verticalAlign = 'middle';
      td.style.textAlign = 'center';

      // 셀 병합 처리
      if (cell.isMerged) {
        const master = cell.master;
        if (master && master.row !== r) {
          // 병합된 셀의 하위 셀은 숨김
          td.style.display = 'none';
        } else if (master) {
          // 병합된 셀의 마스터 셀
          const mergeRange = worksheet.model.merges?.find((merge: string) => {
            const match = merge.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
            if (match) {
              const startRow = parseInt(match[2]);
              const startCol = match[1];
              return startRow === r;
            }
            return false;
          });
          
          if (mergeRange) {
            // rowspan/colspan 계산
            const match = mergeRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
            if (match) {
              const startRow = parseInt(match[2]);
              const endRow = parseInt(match[4]);
              td.rowSpan = endRow - startRow + 1;
            }
          }
        }
      }

      // 셀 값 표시
      const cellValue = cell.value;
      let displayValue = '';
      
      if (cellValue !== null && cellValue !== undefined) {
        if (typeof cellValue === 'object') {
          if ('result' in cellValue) {
            displayValue = cellValue.result?.toString() || '';
          } else if ('text' in cellValue) {
            displayValue = (cellValue as any).text.toString();
          } else {
            displayValue = cellValue.toString();
          }
        } else {
          displayValue = cellValue.toString();
        }
      }

      // 서명 배치 확인
      const assignKey = `${r}:${c}`;
      const assignment = assignments.get(assignKey);
      
      if (assignment) {
        const sigFile = findSigFile(assignment.signatureBaseName, assignment.signatureVariantId);
        if (sigFile) {
          // 서명 이미지 추가
          const img = document.createElement('img');
          img.src = sigFile.previewUrl;
          img.style.maxWidth = '120px';
          img.style.maxHeight = '60px';
          img.style.display = 'block';
          img.style.margin = '0 auto';
          img.style.transform = `rotate(${assignment.rotation}deg) scale(${assignment.scale})`;
          td.appendChild(img);
        }
      } else if (displayValue && !isSignaturePlaceholder(displayValue)) {
        // 일반 텍스트 표시 (placeholder가 아닌 경우)
        td.textContent = displayValue;
      }

      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  container.appendChild(table);

  try {
    // html2canvas로 렌더링
    const canvas = await html2canvas(container, {
      backgroundColor: '#ffffff',
      scale: 2, // 고해상도
      logging: false,
      useCORS: true,
    });

    return canvas;
  } finally {
    // 정리
    document.body.removeChild(container);
  }
};

/**
 * PNG 이미지로 내보내기
 */
export const exportToPNG = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>,
  filename: string = 'export.png'
): Promise<void> => {
  console.log('[PNG 내보내기] 시작...');
  
  const canvas = await renderSheetToCanvas(originalBuffer, assignments, signaturesMap, true);
  
  // Canvas를 Blob으로 변환 (Promise로 감싸서 에러 처리)
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error('이미지 생성 실패'));
        return;
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      console.log(`[PNG 내보내기] 완료: ${filename}`);
      resolve();
    }, 'image/png');
  });
};

/**
 * PDF로 내보내기
 */
export const exportToPDF = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>,
  filename: string = 'export.pdf'
): Promise<void> => {
  console.log('[PDF 내보내기] 시작...');
  
  const canvas = await renderSheetToCanvas(originalBuffer, assignments, signaturesMap, true);
  
  // Canvas 크기에 맞는 PDF 생성
  const imgWidth = 210; // A4 width in mm
  const imgHeight = (canvas.height * imgWidth) / canvas.width;
  
  const pdf = new jsPDF({
    orientation: imgHeight > imgWidth ? 'portrait' : 'landscape',
    unit: 'mm',
    format: 'a4'
  });

  const imgData = canvas.toDataURL('image/png');
  
  // 이미지를 PDF에 추가 (여러 페이지가 필요한 경우 처리)
  if (imgHeight > 297) { // A4 height
    // 큰 이미지는 여러 페이지로 나누기
    let heightLeft = imgHeight;
    let position = 0;
    
    pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
    heightLeft -= 297;
    
    while (heightLeft > 0) {
      position = heightLeft - imgHeight;
      pdf.addPage();
      pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
      heightLeft -= 297;
    }
  } else {
    pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
  }

  pdf.save(filename);
  console.log(`[PDF 내보내기] 완료: ${filename}`);
};
