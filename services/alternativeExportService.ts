import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import ExcelJS from 'exceljs';
import { SignatureAssignment, SignatureFile } from '../types';
import { columnLetterToNumber, isSignaturePlaceholder, parsePrintAreaBounds } from './excelUtils';

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
    const bounds = parsePrintAreaBounds(
      originalPrintArea,
      worksheet.actualRowCount || 100,
      worksheet.actualColumnCount || 26
    );
    printAreaRows = bounds.rows;
    printAreaCols = bounds.cols;
  }

  console.log(`[렌더링] 행: ${printAreaRows.start}-${printAreaRows.end}, 열: ${printAreaCols.start}-${printAreaCols.end}`);

  // HTML 테이블 생성
  const container = document.createElement('div');
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.top = '0';
  container.style.background = 'white';
  container.style.padding = '20px';
  // 렌더링 품질 보강: 폰트/테두리 깨짐을 줄이기 위한 브라우저 렌더 힌트
  container.style.textRendering = 'geometricPrecision';
  container.style.webkitFontSmoothing = 'antialiased';
  container.style.fontKerning = 'normal';
  document.body.appendChild(container);

  const table = document.createElement('table');
  table.style.borderCollapse = 'collapse';
  table.style.fontFamily = 'Arial, sans-serif';
  table.style.fontSize = '12px';
  table.style.background = 'white';
  table.style.tableLayout = 'fixed';
  table.style.borderSpacing = '0';

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
      td.style.boxSizing = 'border-box';
      td.style.whiteSpace = 'pre-wrap';
      td.style.overflow = 'hidden';

      // 엑셀 셀 스타일을 최대한 HTML에 반영하여 PDF/PNG에서도 시각 일관성 유지
      if (cell.style?.font) {
        if (cell.style.font.name) td.style.fontFamily = cell.style.font.name;
        if (cell.style.font.size) td.style.fontSize = `${cell.style.font.size}px`;
        if (cell.style.font.bold) td.style.fontWeight = '700';
        if (cell.style.font.italic) td.style.fontStyle = 'italic';
      }
      if (cell.style?.alignment) {
        const horizontal = cell.style.alignment.horizontal;
        const vertical = cell.style.alignment.vertical;
        if (horizontal === 'left' || horizontal === 'center' || horizontal === 'right' || horizontal === 'justify') {
          td.style.textAlign = horizontal;
        }
        if (vertical === 'top' || vertical === 'middle' || vertical === 'bottom') {
          td.style.verticalAlign = vertical;
        }
      }

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
              const startCol = columnLetterToNumber(match[1]);
              return startRow === r && startCol === c;
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
            displayValue = cellValue.result?.toString() ?? '';
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
        if (sigFile && sigFile.previewUrl && sigFile.width > 0 && sigFile.height > 0) {
          // placeholder 텍스트 유지 후 서명 이미지 오버레이
          td.style.position = 'relative';
          if (displayValue) {
            const span = document.createElement('span');
            span.textContent = displayValue;
            span.style.position = 'relative';
            span.style.zIndex = '0';
            td.appendChild(span);
          }
          const img = document.createElement('img');
          img.src = sigFile.previewUrl;
          img.style.maxWidth = '120px';
          img.style.maxHeight = '60px';
          img.style.display = 'block';
          img.style.position = 'absolute';
          img.style.top = '0';
          img.style.left = '50%';
          img.style.zIndex = '1';
          img.style.transform = `translateX(-50%) rotate(${assignment.rotation}deg) scale(${assignment.scale})`;
          td.appendChild(img);
        } else if (displayValue) {
          // 방어적 코드: 손상된 서명 파일이면 프로세스를 중단하지 않고 텍스트만 유지
          td.textContent = displayValue;
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
      scale: 2, // 요청사항 반영: PDF/PNG 기본 고해상도 렌더링
      logging: false,
      useCORS: true,
      allowTaint: false,
      removeContainer: true,
      foreignObjectRendering: false,
      imageTimeout: 15000,
      onclone: (clonedDocument) => {
        // 복제 DOM에서도 폰트/테두리 렌더 안정성을 유지하기 위한 스타일 보강
        const clonedBody = clonedDocument.body;
        clonedBody.style.textRendering = 'geometricPrecision';
        clonedBody.style.webkitFontSmoothing = 'antialiased';
      },
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
        canvas.width = 1;
        canvas.height = 1;
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
      // 메모리 최적화: 다운로드 직후 canvas 버퍼 참조 해제
      canvas.width = 1;
      canvas.height = 1;

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
  
  // PDF 페이지 크기 상수
  const A4_WIDTH_MM = 210;
  const A4_HEIGHT_MM = 297;
  
  // Canvas 크기에 맞는 PDF 생성
  const imgWidth = A4_WIDTH_MM;
  const imgHeight = (canvas.height * imgWidth) / canvas.width;
  
  const pdf = new jsPDF({
    orientation: imgHeight > imgWidth ? 'portrait' : 'landscape',
    unit: 'mm',
    format: 'a4'
  });

  let imgData = canvas.toDataURL('image/png');
  
  // 이미지를 PDF에 추가 (여러 페이지가 필요한 경우 처리)
  if (imgHeight > A4_HEIGHT_MM) {
    // 큰 이미지는 여러 페이지로 나누기
    let heightLeft = imgHeight;
    let position = 0;
    
    pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
    heightLeft -= A4_HEIGHT_MM;
    
    while (heightLeft > 0) {
      position = heightLeft - imgHeight;
      pdf.addPage();
      pdf.addImage(imgData, 'PNG', 0, position, imgWidth, imgHeight);
      heightLeft -= A4_HEIGHT_MM;
    }
  } else {
    pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
  }

  pdf.save(filename);
  // 메모리 최적화: 대용량 문자열/캔버스 버퍼 즉시 해제
  imgData = '';
  canvas.width = 1;
  canvas.height = 1;
  console.log(`[PDF 내보내기] 완료: ${filename}`);
};
