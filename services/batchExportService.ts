import JSZip from 'jszip';
import { SignatureFile, SheetData } from '../types';
import { autoMatchSignatures, generateFinalExcel } from './excelService';

export interface BatchExportProgress {
  current: number;
  total: number;
  phase: 'generate' | 'zip' | 'done';
  percent: number;
  mode?: 'standard' | 'high-volume';
  message?: string;
}

export interface BatchExcelZipOptions {
  originalBuffer: ArrayBuffer;
  sheetData: SheetData;
  signatures: Map<string, SignatureFile[]>;
  sourceFileName: string;
  count: number;
  variationStrength: number;
  onProgress?: (progress: BatchExportProgress) => void;
  signal?: AbortSignal;
}

/**
 * AbortSignal 상태를 점검하여 취소 요청 시 즉시 AbortError를 발생시킨다.
 */
const throwIfAborted = (signal?: AbortSignal) => {
  if (signal?.aborted) {
    throw new DOMException('사용자에 의해 작업이 취소되었습니다.', 'AbortError');
  }
};

/**
 * N개의 무작위 서명 버전을 생성해 ZIP Blob으로 반환
 *
 * 설계 포인트:
 * 1) 매 회차 autoMatchSignatures를 다시 호출해 완전히 독립된 랜덤 배치를 생성한다.
 * 2) 생성 단계와 ZIP 압축 단계를 분리해 진행률을 명확히 전달한다.
 * 3) 루프 중 setTimeout(0)으로 이벤트 루프를 양보해 브라우저 멈춤 현상을 완화한다.
 */
export const buildBatchExcelZip = async (options: BatchExcelZipOptions): Promise<Blob> => {
  const {
    originalBuffer,
    sheetData,
    signatures,
    sourceFileName,
    count,
    variationStrength,
    onProgress,
    signal,
  } = options;

  const total = Math.max(1, Math.min(50, count));
  const zip = new JSZip();
  const HIGH_VOLUME_THRESHOLD = 20;
  const mode: 'standard' | 'high-volume' = total >= HIGH_VOLUME_THRESHOLD ? 'high-volume' : 'standard';

  const timestamp = new Date().toISOString().slice(11, 19).replace(/:/g, '');
  const baseFileName = sourceFileName.replace(/\.xlsx$/i, '');

  onProgress?.({
    current: 0,
    total,
    phase: 'generate',
    percent: 0,
    mode,
    message: mode === 'high-volume' ? '고용량 모드로 파일 생성을 시작합니다.' : '표준 모드로 파일 생성을 시작합니다.',
  });

  for (let index = 0; index < total; index++) {
    throwIfAborted(signal);

    // 브라우저 렌더링/입력 이벤트 처리 시간을 확보
    await new Promise(resolve => setTimeout(resolve, mode === 'high-volume' ? 16 : 0));
    throwIfAborted(signal);

    const assignments = autoMatchSignatures(sheetData, signatures, { variationStrength });
    const excelBlob = await generateFinalExcel(originalBuffer, assignments, signatures);
    throwIfAborted(signal);

    // JSZip에는 ArrayBuffer로 넣어 메모리 복사 오버헤드를 줄인다.
    const fileBuffer = await excelBlob.arrayBuffer();
    const sequence = String(index + 1).padStart(2, '0');

    zip.file(`서명완료_${timestamp}_${baseFileName}_${sequence}.xlsx`, fileBuffer);

    onProgress?.({
      current: index + 1,
      total,
      phase: 'generate',
      percent: ((index + 1) / total) * 75,
      mode,
      message: mode === 'high-volume'
        ? '고용량 모드: UI 응답성을 유지하며 생성 중입니다.'
        : '파일을 생성 중입니다.',
    });
  }

  const compressionLevel = mode === 'high-volume' ? 3 : 6;

  throwIfAborted(signal);

  const zipPromise = zip.generateAsync(
    {
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    },
    (metadata) => {
      onProgress?.({
        current: total,
        total,
        phase: 'zip',
        percent: 75 + (metadata.percent / 100) * 25,
        mode,
        message: `ZIP 압축 중 (${compressionLevel} 레벨)`,
      });
    }
  );

  const zipBlob = signal
    ? await Promise.race<Blob>([
        zipPromise,
        new Promise<Blob>((_, reject) => {
          signal.addEventListener(
            'abort',
            () => reject(new DOMException('사용자에 의해 작업이 취소되었습니다.', 'AbortError')),
            { once: true }
          );
        }),
      ])
    : await zipPromise;

  throwIfAborted(signal);

  onProgress?.({
    current: total,
    total,
    phase: 'done',
    percent: 100,
    mode,
    message: '일괄 생성 및 압축이 완료되었습니다.',
  });
  return zipBlob;
};
