import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import ExcelJS from 'exceljs';
import { SignatureAssignment, SignatureFile } from '../types';
import { columnLetterToNumber, isSignaturePlaceholder, parsePrintAreaBounds } from './excelUtils';

export interface PreviewCellModel {
  key: string;
  row: number;
  col: number;
  text: string;
  hidden: boolean;
  rowSpan?: number;
  colSpan?: number;
  style: {
    fontFamily: string;
    fontSize: string;
    fontWeight: string;
    fontStyle: string;
    textAlign: 'left' | 'center' | 'right' | 'justify';
    verticalAlign: 'top' | 'middle' | 'bottom';
  };
  signature?: {
    src: string;
    transform: string;
    opacity: number;
  };
}

export interface PreviewRowModel {
  row: number;
  cells: PreviewCellModel[];
}

export interface SheetPreviewModel {
  rows: PreviewRowModel[];
  printAreaRows: { start: number; end: number };
  printAreaCols: { start: number; end: number };
}

/**
 * 병합 범위 문자열(예: A1:C3)을 파싱한다.
 * - 잘못된 범위 문자열은 null로 반환하여 호출부에서 안전하게 무시한다.
 */
const parseMergeRange = (range: string): { startRow: number; endRow: number; startCol: number; endCol: number } | null => {
  const match = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
  if (!match) return null;

  const startCol = columnLetterToNumber(match[1]);
  const startRow = parseInt(match[2], 10);
  const endCol = columnLetterToNumber(match[3]);
  const endRow = parseInt(match[4], 10);

  if (!startCol || !startRow || !endCol || !endRow) return null;
  return { startRow, endRow, startCol, endCol };
};

/**
 * 엑셀/미리보기 공용 모델 생성
 * - alternativeExportService의 HTML 렌더링 로직과 동일한 데이터 기반으로
 *   React 실시간 미리보기를 구성하기 위해 사용한다.
 */
export const buildSheetPreviewModel = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>,
  printAreaOnly: boolean = true
): Promise<SheetPreviewModel> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalBuffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('워크시트를 찾을 수 없습니다.');
  }

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

  // 병합셀 인덱스 구성
  const mergeRanges = (worksheet.model.merges || []) as string[];
  const hiddenCellSet = new Set<string>();
  const topLeftMergeMap = new Map<string, { rowSpan: number; colSpan: number }>();

  for (const mergeText of mergeRanges) {
    const parsed = parseMergeRange(mergeText);
    if (!parsed) continue;

    const { startRow, endRow, startCol, endCol } = parsed;
    topLeftMergeMap.set(`${startRow}:${startCol}`, {
      rowSpan: endRow - startRow + 1,
      colSpan: endCol - startCol + 1,
    });

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (!(r === startRow && c === startCol)) {
          hiddenCellSet.add(`${r}:${c}`);
        }
      }
    }
  }

  const findSigFile = (name: string, variant: string) => {
    const list = signaturesMap.get(name);
    return list?.find(s => s.variant === variant);
  };

  const rows: PreviewRowModel[] = [];

  for (let r = printAreaRows.start; r <= printAreaRows.end; r++) {
    const row = worksheet.getRow(r);
    const cells: PreviewCellModel[] = [];

    for (let c = printAreaCols.start; c <= printAreaCols.end; c++) {
      const cell = row.getCell(c);
      const cellKey = `${r}:${c}`;
      const isHidden = hiddenCellSet.has(cellKey);

      let displayValue = '';
      const cellValue = cell.value;
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

      const assignment = assignments.get(cellKey);
      let signature: PreviewCellModel['signature'];

      if (assignment) {
        const sigFile = findSigFile(assignment.signatureBaseName, assignment.signatureVariantId);
        if (sigFile && sigFile.previewUrl && sigFile.width > 0 && sigFile.height > 0) {
          // 미리보기에서 실제 배치와 동일한 변형값을 보여주되,
          // 글자/테두리 가림을 줄이기 위해 투명도와 blend를 함께 사용한다.
          signature = {
            src: sigFile.previewUrl,
            transform: `translate(${assignment.offsetX}px, ${assignment.offsetY}px) rotate(${assignment.rotation}deg) scale(${assignment.scale})`,
            opacity: 0.78,
          };
        }
      }

      const horizontal = cell.style?.alignment?.horizontal;
      const vertical = cell.style?.alignment?.vertical;

      const previewCell: PreviewCellModel = {
        key: `${r}-${c}`,
        row: r,
        col: c,
        text: displayValue,
        hidden: isHidden,
        style: {
          fontFamily: cell.style?.font?.name || 'Arial, sans-serif',
          fontSize: `${cell.style?.font?.size || 12}px`,
          fontWeight: cell.style?.font?.bold ? '700' : '400',
          fontStyle: cell.style?.font?.italic ? 'italic' : 'normal',
          textAlign: (horizontal === 'left' || horizontal === 'center' || horizontal === 'right' || horizontal === 'justify') ? horizontal : 'center',
          verticalAlign: (vertical === 'top' || vertical === 'middle' || vertical === 'bottom') ? vertical : 'middle',
        },
        signature,
      };

      const mergeInfo = topLeftMergeMap.get(cellKey);
      if (mergeInfo) {
        previewCell.rowSpan = mergeInfo.rowSpan;
        previewCell.colSpan = mergeInfo.colSpan;
      }

      cells.push(previewCell);
    }

    rows.push({ row: r, cells });
  }

  return {
    rows,
    printAreaRows,
    printAreaCols,
  };
};

/**
 * 프리뷰 모델을 실제 HTML 테이블로 변환
 * - PDF/PNG 렌더링과 React 미리보기의 결과를 최대한 일치시키기 위해
 *   공용 모델을 동일하게 사용한다.
 */
const createTableFromPreviewModel = (model: SheetPreviewModel): HTMLTableElement => {
  const table = document.createElement('table');
  table.style.borderCollapse = 'collapse';
  table.style.fontFamily = 'Arial, sans-serif';
  table.style.fontSize = '12px';
  table.style.background = 'white';
  table.style.tableLayout = 'fixed';
  table.style.borderSpacing = '0';

  for (const rowModel of model.rows) {
    const tr = document.createElement('tr');

    for (const cellModel of rowModel.cells) {
      if (cellModel.hidden) {
        continue;
      }

      const td = document.createElement('td');
      td.style.border = '1px solid #ddd';
      td.style.padding = '8px';
      td.style.minWidth = '80px';
      td.style.minHeight = '40px';
      td.style.position = 'relative';
      td.style.boxSizing = 'border-box';
      td.style.whiteSpace = 'pre-wrap';
      td.style.overflow = 'hidden';
      td.style.textAlign = cellModel.style.textAlign;
      td.style.verticalAlign = cellModel.style.verticalAlign;
      td.style.fontFamily = cellModel.style.fontFamily;
      td.style.fontSize = cellModel.style.fontSize;
      td.style.fontWeight = cellModel.style.fontWeight;
      td.style.fontStyle = cellModel.style.fontStyle;

      if (cellModel.rowSpan && cellModel.rowSpan > 1) {
        td.rowSpan = cellModel.rowSpan;
      }
      if (cellModel.colSpan && cellModel.colSpan > 1) {
        td.colSpan = cellModel.colSpan;
      }

      if (cellModel.signature) {
        if (cellModel.text) {
          const span = document.createElement('span');
          span.textContent = cellModel.text;
          span.style.position = 'relative';
          span.style.zIndex = '0';
          td.appendChild(span);
        }

        const img = document.createElement('img');
        img.src = cellModel.signature.src;
        img.style.maxWidth = '120px';
        img.style.maxHeight = '60px';
        img.style.display = 'block';
        img.style.position = 'absolute';
        img.style.top = '50%';
        img.style.left = '50%';
        img.style.zIndex = '1';
        img.style.opacity = `${cellModel.signature.opacity}`;
        img.style.mixBlendMode = 'multiply';
        img.style.transform = `translate(-50%, -50%) ${cellModel.signature.transform}`;
        td.appendChild(img);
      } else if (cellModel.text && !isSignaturePlaceholder(cellModel.text)) {
        td.textContent = cellModel.text;
      } else {
        td.textContent = cellModel.text;
      }

      tr.appendChild(td);
    }

    table.appendChild(tr);
  }

  return table;
};

/**
 * 엑셀 시트를 HTML 테이블로 렌더링하여 이미지로 변환하는 헬퍼 함수
 */
const renderSheetToCanvas = async (
  originalBuffer: ArrayBuffer,
  assignments: Map<string, SignatureAssignment>,
  signaturesMap: Map<string, SignatureFile[]>,
  printAreaOnly: boolean = true
): Promise<HTMLCanvasElement> => {
  const previewModel = await buildSheetPreviewModel(originalBuffer, assignments, signaturesMap, printAreaOnly);
  console.log(`[렌더링] 행: ${previewModel.printAreaRows.start}-${previewModel.printAreaRows.end}, 열: ${previewModel.printAreaCols.start}-${previewModel.printAreaCols.end}`);

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

  const table = createTableFromPreviewModel(previewModel);

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
