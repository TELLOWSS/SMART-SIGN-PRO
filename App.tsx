import React, { useState, useEffect, useRef, useMemo } from 'react';
import { Upload, FileSpreadsheet, Image as ImageIcon, CheckCircle, RotateCcw, Download, Settings, RefreshCw, AlertCircle, HelpCircle, X, ArrowRight, FileText, MousePointer2, Copy, FileDown, FolderUp, PenTool, Users, ShieldCheck, HardDrive, Trash2, Sparkles, Lock } from 'lucide-react';
import { parseExcelFile, autoMatchSignatures, generateFinalExcel, normalizeName } from './services/excelService';
import { buildSheetPreviewModel, exportToPDF, exportToPNG, SheetPreviewModel } from './services/alternativeExportService';
import { extractFilesFromDataTransfer, extractWorkerRowsFromSheet, detectHeaderColumns, detectSelectableColumns, buildWorkerRowMap, randomInt, randomFloat } from './services/excelUtils';
import { saveSignaturesToVault, loadSignaturesFromVault, clearSignatureVault } from './services/signatureStorage';
import { AppState, SignatureFile, SheetData, SignatureAssignment, SignaturePlacementMode } from './types';
import SignatureWorkspace from './components/SignatureWorkspace';
import SignaturePadModal from './components/SignaturePadModal';
import ManualMatchModal from './components/ManualMatchModal';
import { buildBatchExcelZip, BatchExportProgress } from './services/batchExportService';

// Factory function to ensure fresh state on reset
const getInitialState = (): AppState => ({
  step: 'upload',
  excelFile: null,
  excelBuffer: null,
  sheetData: null,
  signatures: new Map(),
  assignments: new Map(),
  availableSheets: [],
  activeSheetIndex: 0,
  maskResidentNumber: false,
  signatureScalePreset: 1.0,
  placementMode: 'auto',
  enhanceSignatureInk: true,
});

export default function App() {
  const [state, setState] = useState<AppState>(getInitialState());
  const [processing, setProcessing] = useState(false);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewModel, setPreviewModel] = useState<SheetPreviewModel | null>(null);
  const [variationStrength, setVariationStrength] = useState(70);
  const [batchCount, setBatchCount] = useState(5);
  const [batchProgress, setBatchProgress] = useState<BatchExportProgress | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [showGuide, setShowGuide] = useState(false);
  const [toast, setToast] = useState<{msg: string, type: 'success' | 'info'} | null>(null);
  const [exportFormat, setExportFormat] = useState<'excel' | 'pdf' | 'png'>('excel');
  const signaturesRef = useRef<Map<string, SignatureFile[]>>(new Map());
  const batchAbortRef = useRef<AbortController | null>(null);
  const objectUrlRegistryRef = useRef<Set<string>>(new Set());

  // 5대 실무 기능 모달 및 보관함/히스토리 상태
  const [isSignaturePadOpen, setIsSignaturePadOpen] = useState(false);
  const [isManualMatchOpen, setIsManualMatchOpen] = useState(false);
  const [signaturePadTargetWorker, setSignaturePadTargetWorker] = useState<string>('');
  const [vaultRestoredCount, setVaultRestoredCount] = useState<number>(0);
  const [historyPast, setHistoryPast] = useState<Map<string, SignatureAssignment>[]>([]);
  const [historyFuture, setHistoryFuture] = useState<Map<string, SignatureAssignment>[]>([]);
  
  // Refs to clear file inputs
  const excelInputRef = useRef<HTMLInputElement>(null);
  const sigInputRef = useRef<HTMLInputElement>(null);
  const sigFolderInputRef = useRef<HTMLInputElement>(null);

  // Drag and Drop 상태
  const [isDraggingExcel, setIsDraggingExcel] = useState(false);
  const [isDraggingSig, setIsDraggingSig] = useState(false);

  // 앱 시작 시 브라우저 로컬 서명 보관함(IndexedDB)에서 자동 복원
  useEffect(() => {
    let mounted = true;
    loadSignaturesFromVault((url) => objectUrlRegistryRef.current.add(url)).then((restored) => {
      if (!mounted || restored.size === 0) return;
      setState(prev => {
        if (prev.signatures.size > 0) return prev;
        return { ...prev, signatures: restored };
      });
      setVaultRestoredCount(restored.size);
    });
    return () => {
      mounted = false;
    };
  }, []);

  // 서명 배치 히스토리 기록 (Undo / Redo 지원)
  const commitAssignmentsWithHistory = (
    nextAssignments: Map<string, SignatureAssignment>,
    extraState?: Partial<AppState>
  ) => {
    setHistoryPast(prev => [...prev.slice(-29), new Map(state.assignments)]);
    setHistoryFuture([]);
    setState(prev => ({ ...prev, ...extraState, assignments: nextAssignments }));
  };

  const handleUndo = () => {
    if (historyPast.length === 0) return;
    const previous = historyPast[historyPast.length - 1];
    setHistoryPast(prev => prev.slice(0, -1));
    setHistoryFuture(prev => [new Map(state.assignments), ...prev.slice(0, 29)]);
    setState(prev => ({ ...prev, assignments: previous }));
  };

  const handleRedo = () => {
    if (historyFuture.length === 0) return;
    const next = historyFuture[0];
    setHistoryFuture(prev => prev.slice(1));
    setHistoryPast(prev => [...prev.slice(-29), new Map(state.assignments)]);
    setState(prev => ({ ...prev, assignments: next }));
  };

  // 키보드 단축키: Ctrl+Z (실행 취소) / Ctrl+Y 또는 Ctrl+Shift+Z (다시 실행)
  useEffect(() => {
    if (state.step !== 'preview') return;
    const onKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT')) {
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) handleRedo();
        else handleUndo();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        handleRedo();
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [state.step, historyPast, historyFuture, state.assignments]);

  // 브라우저 서명 보관함 비우기
  const handleClearVault = async () => {
    await clearSignatureVault();
    cleanupBlobUrls(state.signatures);
    setState(prev => ({ ...prev, signatures: new Map(), assignments: new Map() }));
    setVaultRestoredCount(0);
    setToast({ msg: '🗑️ 브라우저 서명 보관함을 비웠습니다.', type: 'info' });
  };

  // 브라우저 기본 드래그 방지 (바탕화면/탐색기에서 파일/폴더 드롭 시 새 창 열림/이탈 방지)
  useEffect(() => {
    const handleDragOver = (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
    };
    const handleDrop = (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
    };
    window.addEventListener('dragover', handleDragOver);
    window.addEventListener('drop', handleDrop);
    return () => {
      window.removeEventListener('dragover', handleDragOver);
      window.removeEventListener('drop', handleDrop);
    };
  }, []);

  // Auto-hide toast
  useEffect(() => {
    if (toast) {
      const timer = setTimeout(() => setToast(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [toast]);

  useEffect(() => {
    signaturesRef.current = state.signatures;
  }, [state.signatures]);

  /**
   * 생성된 Object URL을 추적 등록
   * - Start Over 시점에 일괄 해제하여 메모리 누수를 방지한다.
   */
  const trackObjectUrl = (url: string) => {
    if (!url) return;
    objectUrlRegistryRef.current.add(url);
  };

  /**
   * 특정 Object URL 해제 + 추적 목록 제거
   */
  const revokeTrackedObjectUrl = (url: string) => {
    if (!url) return;
    try {
      URL.revokeObjectURL(url);
    } catch (e) {
      console.warn('Failed to revoke object URL:', e);
    } finally {
      objectUrlRegistryRef.current.delete(url);
    }
  };

  /**
   * 추적된 Object URL 전체 해제
   * - 다운로드 Blob URL, 미리보기용 시그니처 URL 등 이전 작업 잔존 리소스를 모두 정리
   */
  const revokeAllTrackedObjectUrls = () => {
    for (const url of objectUrlRegistryRef.current) {
      try {
        URL.revokeObjectURL(url);
      } catch (e) {
        console.warn('Failed to revoke object URL:', e);
      }
    }
    objectUrlRegistryRef.current.clear();
  };

  useEffect(() => {
    return () => {
      cleanupBlobUrls(signaturesRef.current);
      revokeAllTrackedObjectUrls();
    };
  }, []);

  /**
   * 실시간 미리보기 모델 생성
   * - '내보내기' 전에 실제 출력과 유사한 결과를 즉시 확인할 수 있도록
   *   대체 내보내기 서비스의 HTML 변환 로직 기반 모델을 재사용한다.
   * - 주민번호 마스킹 및 선택된 활성 시트 반영
   */
  useEffect(() => {
    if (state.step !== 'preview' || !state.excelBuffer) {
      return;
    }

    let canceled = false;

    const loadPreview = async () => {
      setPreviewLoading(true);

      try {
        const model = await buildSheetPreviewModel(
          state.excelBuffer as ArrayBuffer,
          state.assignments,
          state.signatures,
          true,
          {
            maskResidentNumber: state.maskResidentNumber,
            targetSheetIndex: state.activeSheetIndex,
          }
        );

        if (!canceled) {
          setPreviewModel(model);
        }
      } catch (previewErr) {
        console.error('Preview build error:', previewErr);
        if (!canceled) {
          setError('미리보기 생성 중 오류가 발생했습니다. 파일을 다시 확인해주세요.');
          setPreviewModel(null);
        }
      } finally {
        if (!canceled) {
          setPreviewLoading(false);
        }
      }
    };

    loadPreview();

    return () => {
      canceled = true;
    };
  }, [state.step, state.excelBuffer, state.assignments, state.signatures, state.maskResidentNumber, state.activeSheetIndex]);

  // 감지된 전용 '서명' 열 라벨 계산
  const detectedSignatureColLabel = useMemo(() => {
    if (!state.sheetData) return null;
    const detected = detectHeaderColumns(state.sheetData.rows);
    if (detected.signatureCols.length === 0) return null;
    return detected.signatureCols.map(s => s.label).join(', ');
  }, [state.sheetData]);

  // '1' 없이도 클릭으로 자유 지정 가능한 모든 품목/서명 열 + 현재 배치 건수
  const selectableColumns = useMemo(() => {
    if (!state.sheetData) return [];
    const cols = detectSelectableColumns(state.sheetData.rows);
    const colCounts = new Map<number, number>();
    for (const a of state.assignments.values()) {
      colCounts.set(a.col, (colCounts.get(a.col) || 0) + 1);
    }
    return cols.map(c => ({
      ...c,
      assignedCount: colCounts.get(c.col) || 0,
    }));
  }, [state.sheetData, state.assignments]);

  // 1-based 행 번호 -> 근로자 성명 매핑 (라이브 프리뷰 셀 직접 클릭용)
  const workerRowMap = useMemo(() => {
    if (!state.sheetData) return new Map<number, { workerName: string; nameCol: number }>();
    return buildWorkerRowMap(state.sheetData.rows);
  }, [state.sheetData]);

  // 근로자 매칭 상태 목록 및 미매칭 카운트 계산
  const workerItems = useMemo(() => {
    if (!state.sheetData) return [];
    return extractWorkerRowsFromSheet(state.sheetData.rows, state.assignments, state.signatures, state.placementMode);
  }, [state.sheetData, state.assignments, state.signatures, state.placementMode]);

  const unmatchedWorkerCount = useMemo(() => {
    return workerItems.filter(w => !w.isMatched).length;
  }, [workerItems]);

  const getUserFacingExportError = (errorMessage: string, format: 'excel' | 'pdf' | 'png') => {
    if (/out of memory|allocation failed|array buffer allocation failed/i.test(errorMessage)) {
      return `메모리 부족으로 ${format.toUpperCase()} 내보내기에 실패했습니다.\n파일 크기나 이미지 해상도를 낮춘 뒤 다시 시도해주세요.`;
    }

    if (/파일이 비어|empty|0 bytes|too small|손상|zip 형식/i.test(errorMessage)) {
      return `생성된 파일이 유효하지 않습니다.\n원본 파일 형식(XLSX)과 서명 이미지를 확인한 뒤 다시 시도해주세요.`;
    }

    if (/worksheet|워크시트|print area|인쇄영역/i.test(errorMessage)) {
      return `워크시트 구조를 처리하는 중 오류가 발생했습니다.\n병합셀/인쇄영역이 복잡한 경우 단순화한 파일로 먼저 테스트해주세요.`;
    }

    return `${format.toUpperCase()} 내보내기 중 오류가 발생했습니다.\n잠시 후 다시 시도하거나 브라우저를 새로고침해주세요.`;
  };

  // --- Handlers ---

  const processExcelFile = async (file: File) => {
    if (!file) return;

    // Validate file type (지원 형식: .xlsx, .xls, .csv, .xlsm, .xlsb)
    const isSupportedExtension = /\.(xlsx|xls|csv|xlsm|xlsb)$/i.test(file.name);
    const isSupportedMime = file.type.includes('spreadsheet') || file.type.includes('excel') || file.type.includes('csv') || file.type === '';

    if (!isSupportedExtension && !isSupportedMime) {
      setError("엑셀 파일(.xlsx, .xls, .csv)만 지원합니다. 파일 확장명을 확인해주세요.");
      if (excelInputRef.current) excelInputRef.current.value = '';
      return;
    }

    // Check file size (Warning if > 15MB)
    if (file.size > 15 * 1024 * 1024) {
      if (!window.confirm(`선택하신 엑셀 파일의 용량이 큽니다 (${(file.size / 1024 / 1024).toFixed(1)}MB).\n파일 내에 이미지가 많거나 행이 매우 많으면 브라우저 메모리 부하가 발생할 수 있습니다.\n\n계속 진행하시겠습니까?`)) {
        if (excelInputRef.current) excelInputRef.current.value = '';
        return;
      }
    }

    try {
      setProcessing(true);
      setError(null);
      const buffer = await file.arrayBuffer();
      
      if (buffer.byteLength === 0) {
        throw new Error("파일이 비어있습니다.");
      }

      const parseResult = await parseExcelFile(buffer, file.name);
      const sheetData = parseResult;
      const normalizedBuffer = parseResult.normalizedBuffer || buffer;
      const availableSheets = parseResult.sheetNames || [sheetData.name];
      const activeSheetIndex = parseResult.activeSheetIndex ?? 0;
      
      if (sheetData.rows.length === 0) {
        throw new Error("데이터가 없는 파일입니다. 성명 열과 데이터가 포함된 파일을 확인해주세요.");
      }

      setState(prev => ({
        ...prev,
        excelFile: file,
        excelBuffer: normalizedBuffer,
        sheetData,
        availableSheets,
        activeSheetIndex,
        step: 'upload'
      }));
      if (parseResult.isTranscoded) {
        setToast({ msg: `${file.name} 호환 모드로 자동 변환 로드됨 (${sheetData.rows.length}개 행)`, type: 'success' });
      } else {
        setToast({ msg: `${file.name} 로드됨 (${sheetData.rows.length}개 행)`, type: 'success' });
      }
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "알 수 없는 오류";
      setError(`엑셀 파일 읽기 실패: ${errorMsg}`);
      console.error('Excel upload error:', err);
    } finally {
      setProcessing(false);
    }
  };

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await processExcelFile(file);
  };

  const processSignatureFiles = async (files: File[] | FileList) => {
    const fileArray = Array.from(files);
    if (!fileArray || fileArray.length === 0) return;

    setProcessing(true);
    const newSignatures = new Map<string, SignatureFile[]>(state.signatures);
    let count = 0;
    const failedFiles: string[] = [];

    for (let i = 0; i < fileArray.length; i++) {
      const file = fileArray[i];

      // 숨김 파일 또는 시스템 파일 건너뛰기
      if (file.name.startsWith('.') || file.name === 'Thumbs.db' || file.name === 'desktop.ini') {
        continue;
      }

      let objectUrl: string | null = null;
      // Only images (MIME check or extension check)
      const isImg = file.type.startsWith('image/') || /\.(png|jpe?g|webp|bmp|gif|tif|tiff)$/i.test(file.name);
      if (!isImg) {
        failedFiles.push(`${file.name} (이미지 파일이 아님)`);
        continue;
      }

      // File size check (warning if > 5MB per image)
      if (file.size > 5 * 1024 * 1024) {
        failedFiles.push(`${file.name} (5MB 초과 - 용량 과다)`);
        continue;
      }

      try {
        const getImageDims = async (): Promise<{ w: number; h: number }> => {
          if ('createImageBitmap' in window) {
            try {
              const bitmap = await createImageBitmap(file);
              const dims = { w: bitmap.width, h: bitmap.height };
              bitmap.close();
              return dims;
            } catch {
              // fallback으로 진행
            }
          }

          return await new Promise<{ w: number; h: number }>((resolve) => {
            const tempUrl = URL.createObjectURL(file);
            const img = new Image();
            img.onload = () => {
              const dims = { w: img.width, h: img.height };
              img.src = '';
              URL.revokeObjectURL(tempUrl);
              resolve(dims);
            };
            img.onerror = () => {
              img.src = '';
              URL.revokeObjectURL(tempUrl);
              resolve({ w: 100, h: 50 });
            };
            img.src = tempUrl;
          });
        };

        const { w, h } = await getImageDims();

        if (w === 100 && h === 50) {
          failedFiles.push(`${file.name} (이미지 로드 실패)`);
          continue;
        }

        // Parse name logic:
        const fileNameNoExt = file.name.substring(0, file.name.lastIndexOf('.'));
        const lastUnderscoreIdx = fileNameNoExt.lastIndexOf('_');
        
        let baseNameString = fileNameNoExt;
        if (lastUnderscoreIdx > 0) {
          baseNameString = fileNameNoExt.substring(0, lastUnderscoreIdx);
        }

        const baseName = normalizeName(baseNameString);
        
        if (!baseName) {
          failedFiles.push(`${file.name} (이름 파싱 불가)`);
          continue;
        }

        // 실제 앱에서 표시/내보내기에 사용하는 URL만 유지
        objectUrl = URL.createObjectURL(file);
        trackObjectUrl(objectUrl);
        
        const sigFile: SignatureFile = {
          name: baseName,
          variant: file.name,
          previewUrl: objectUrl,
          width: w,
          height: h
        };

        const list: SignatureFile[] = newSignatures.get(sigFile.name) || [];
        if (!list.find(s => s.variant === sigFile.variant)) {
          list.push(sigFile);
          newSignatures.set(sigFile.name, list);
          count++;
          // 브라우저 로컬 서명 보관함(IndexedDB)에 자동 저장
          saveSignaturesToVault([{
            name: sigFile.name,
            variant: sigFile.variant,
            blob: file,
            width: w,
            height: h,
          }]).catch(() => {});
        } else {
          if (objectUrl) {
            revokeTrackedObjectUrl(objectUrl);
            objectUrl = null;
          }
          failedFiles.push(`${file.name} (중복)`);
        }
      } catch (err) {
        console.error('Image upload error:', err);
        if (objectUrl) {
          revokeTrackedObjectUrl(objectUrl);
          objectUrl = null;
        }
        failedFiles.push(`${file.name} (처리 실패)`);
      }
    }

    setState(prev => ({ ...prev, signatures: newSignatures }));
    setProcessing(false);
    if (sigInputRef.current) sigInputRef.current.value = '';
    if (sigFolderInputRef.current) sigFolderInputRef.current.value = '';
    
    const toastMsg = failedFiles.length > 0 
      ? `${count}개 추가됨${failedFiles.length > 0 ? ` (${failedFiles.length}개 제외됨: ${failedFiles.slice(0, 2).join(', ')}${failedFiles.length > 2 ? '...' : ''})` : ''}`
      : `${count}개의 서명이 추가되었습니다.`;
    
    setToast({ msg: toastMsg, type: count > 0 ? 'success' : 'info' });
  };

  const handleSignatureUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    await processSignatureFiles(files);
  };

  const handleFolderUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    await processSignatureFiles(files);
  };

  const runAutoMatch = (scaleOverride?: number) => {
    if (!state.sheetData) {
      setError("엑셀 파일이 없습니다.");
      return;
    }

    if (state.signatures.size === 0) {
      setError("업로드된 서명이 없습니다. 서명 이미지를 먼저 추가해주세요.");
      return;
    }

    setProcessing(true);
    setTimeout(() => {
        if (!state.sheetData) {
          setError("시트 데이터가 없습니다.");
          setProcessing(false);
          return;
        }
        const effectiveScale = scaleOverride ?? state.signatureScalePreset;
        const assignments = autoMatchSignatures(state.sheetData, state.signatures, {
          variationStrength,
          scaleMultiplier: effectiveScale,
          placementMode: state.placementMode,
        });
        setState(prev => ({ ...prev, assignments, step: 'preview' }));
        setProcessing(false);
        
        if (assignments.size === 0) {
          setError(null);
          setToast({ msg: '💡 상단 품목 버튼이나 표의 원하는 칸을 클릭해 서명을 자유롭게 지정하세요!', type: 'info' });
        } else {
          const signatureCount = new Set(Array.from(assignments.values()).map(a => a.signatureBaseName)).size;
          setToast({ msg: `✅ ${assignments.size}개 위치에 ${signatureCount}명의 서명이 배치되었습니다`, type: 'success' });
        }
    }, 100);
  };

  /**
   * 특정 근로자(행)에 매칭되는 서명 키와 후보 목록 탐색
   * - 같은 행에 이미 수동 매칭된 서명이 있으면 우선 재사용
   * - 없으면 정확 일치 -> 공백/특수문자 무시 부분 일치 순으로 탐색
   */
  const findSignatureForWorker = (
    workerName: string,
    rowNumber: number,
    currentAssignments: Map<string, SignatureAssignment>,
    signaturesMap: Map<string, SignatureFile[]>
  ): { baseName: string; files: SignatureFile[] } | null => {
    // 1) 같은 행에 이미 수동/자동 지정된 서명이 있으면 재사용
    for (const a of currentAssignments.values()) {
      if (a.row === rowNumber) {
        const files = signaturesMap.get(a.signatureBaseName);
        if (files && files.length > 0) {
          return { baseName: a.signatureBaseName, files };
        }
      }
    }

    // 2) 이름 정규화 매칭 + 스마트 부분 일치
    const cleanName = normalizeName(workerName);
    if (!cleanName) return null;

    const exact = signaturesMap.get(cleanName);
    if (exact && exact.length > 0) {
      return { baseName: cleanName, files: exact };
    }

    for (const [sigKey, sigFiles] of signaturesMap.entries()) {
      const normSigKey = normalizeName(sigKey);
      if (normSigKey && sigFiles.length > 0 && (normSigKey === cleanName || cleanName.includes(normSigKey) || normSigKey.includes(cleanName))) {
        return { baseName: sigKey, files: sigFiles };
      }
    }

    return null;
  };

  /**
   * 라이브 프리뷰 개별 셀 직접 클릭 -> 어느 이름에 어떤 품목도 자유롭게 서명 추가/해제
   */
  const handleToggleCellSignature = (row: number, col: number) => {
    const key = `${row}:${col}`;
    const workerInfo = workerRowMap.get(row);
    if (!workerInfo) return;

    // 이미 서명이 있으면 즉시 해제
    if (state.assignments.has(key)) {
      const next = new Map(state.assignments);
      next.delete(key);
      commitAssignmentsWithHistory(next);
      return;
    }

    // 서명이 없으면 해당 근로자의 서명을 찾아 즉시 추가
    const matched = findSignatureForWorker(workerInfo.workerName, row, state.assignments, state.signatures);
    if (!matched) {
      setToast({
        msg: `⚠️ '${workerInfo.workerName}'님의 서명 파일이 없습니다. [수동 서명 매칭] 또는 [현장 즉석 서명]에서 등록해주세요.`,
        type: 'info',
      });
      setSignaturePadTargetWorker(workerInfo.workerName);
      setIsSignaturePadOpen(true);
      return;
    }

    const strengthFactor = Math.max(0, Math.min(100, variationStrength)) / 100;
    const minScale = 1.0 - 0.06 * strengthFactor;
    const maxScale = 1.0 + 0.06 * strengthFactor;
    const rotLimit = Math.max(1, Math.round(1 + strengthFactor * 3));
    const oxLimit = Math.max(1, Math.round(1 + strengthFactor * 3));
    const oyLimit = Math.max(1, Math.round(1 + strengthFactor * 2));
    const pickDistinctVariant = (files: SignatureFile[], avoidVariant?: string) => {
      if (files.length <= 1) return files[0];
      const others = avoidVariant ? files.filter(f => f.variant !== avoidVariant) : files;
      const pool = others.length > 0 ? others : files;
      return pool[Math.floor(Math.random() * pool.length)];
    };

    const picked = pickDistinctVariant(matched.files);

    const next = new Map(state.assignments);
    next.set(key, {
      row,
      col,
      signatureBaseName: matched.baseName,
      signatureVariantId: picked.variant,
      rotation: randomInt(-rotLimit, rotLimit),
      scale: randomFloat(minScale, maxScale) * state.signatureScalePreset,
      offsetX: randomInt(-oxLimit, oxLimit),
      offsetY: randomInt(-oyLimit, oyLimit),
    });

    // '품목 + 서명 동시' 모드에서는 품목 칸을 클릭하면 해당 행의 전용 '서명' 열에도 받았다는 서명을 자동 추가
    if (state.placementMode === 'auto') {
      const sigCols = selectableColumns.filter(c => c.isSignatureCol);
      for (const sc of sigCols) {
        if (sc.col === col) continue;
        const sigKey = `${row}:${sc.col}`;
        if (!next.has(sigKey)) {
          const picked2 = pickDistinctVariant(matched.files, picked.variant);
          next.set(sigKey, {
            row,
            col: sc.col,
            signatureBaseName: matched.baseName,
            signatureVariantId: picked2.variant,
            rotation: randomInt(-rotLimit, rotLimit),
            scale: randomFloat(minScale, maxScale) * state.signatureScalePreset,
            offsetX: randomInt(-oxLimit, oxLimit),
            offsetY: randomInt(-oyLimit, oyLimit),
          });
        }
      }
    }

    commitAssignmentsWithHistory(next);
  };

  /**
   * 상단 품목/서명 열 칩 클릭 -> '1' 표시 없이도 해당 품목 열 전체 근로자 서명 일괄 지정/해제
   * - '품목 + 서명 동시' 모드에서 품목 열 지정 시 전용 '서명' 열에도 받았다는 서명 자동 연동
   */
  const handleToggleColumnSignatures = (col: number, label: string) => {
    const next = new Map(state.assignments);
    const existingKeysInCol: string[] = [];
    for (const [key, a] of next.entries()) {
      if (a.col === col) existingKeysInCol.push(key);
    }

    if (existingKeysInCol.length > 0) {
      for (const k of existingKeysInCol) next.delete(k);
      commitAssignmentsWithHistory(next);
      setToast({ msg: `🗑️ [${label}] 열 서명 ${existingKeysInCol.length}건 일괄 해제`, type: 'info' });
      return;
    }

    const strengthFactor = Math.max(0, Math.min(100, variationStrength)) / 100;
    const minScale = 1.0 - 0.06 * strengthFactor;
    const maxScale = 1.0 + 0.06 * strengthFactor;
    const rotLimit = Math.max(1, Math.round(1 + strengthFactor * 3));
    const oxLimit = Math.max(1, Math.round(1 + strengthFactor * 3));
    const oyLimit = Math.max(1, Math.round(1 + strengthFactor * 2));
    const sigCols = state.placementMode === 'auto'
      ? selectableColumns.filter(c => c.isSignatureCol && c.col !== col)
      : [];

    let added = 0;
    let linkedSigCount = 0;
    for (const [rowNumber, info] of workerRowMap.entries()) {
      const matched = findSignatureForWorker(info.workerName, rowNumber, next, state.signatures);
      if (!matched) continue;
      const picked = matched.files[Math.floor(Math.random() * matched.files.length)];
      next.set(`${rowNumber}:${col}`, {
        row: rowNumber,
        col,
        signatureBaseName: matched.baseName,
        signatureVariantId: picked.variant,
        rotation: randomInt(-rotLimit, rotLimit),
        scale: randomFloat(minScale, maxScale) * state.signatureScalePreset,
        offsetX: randomInt(-oxLimit, oxLimit),
        offsetY: randomInt(-oyLimit, oyLimit),
      });
      added++;

      for (const sc of sigCols) {
        const sigKey = `${rowNumber}:${sc.col}`;
        if (!next.has(sigKey)) {
          const others = matched.files.length > 1
            ? matched.files.filter(f => f.variant !== picked.variant)
            : matched.files;
          const picked2 = others[Math.floor(Math.random() * others.length)] || picked;
          next.set(sigKey, {
            row: rowNumber,
            col: sc.col,
            signatureBaseName: matched.baseName,
            signatureVariantId: picked2.variant,
            rotation: randomInt(-rotLimit, rotLimit),
            scale: randomFloat(minScale, maxScale) * state.signatureScalePreset,
            offsetX: randomInt(-oxLimit, oxLimit),
            offsetY: randomInt(-oyLimit, oyLimit),
          });
          linkedSigCount++;
        }
      }
    }

    commitAssignmentsWithHistory(next);
    if (added > 0) {
      const extraMsg = linkedSigCount > 0 ? ` + 서명 열 ${linkedSigCount}건 자동 연동` : '';
      setToast({ msg: `✅ [${label}] 열에 ${added}명 서명 지정${extraMsg}`, type: 'success' });
    } else {
      setToast({ msg: `⚠️ 매칭되는 서명 파일이 없습니다. 서명을 먼저 등록해주세요.`, type: 'info' });
    }
  };

  /**
   * 현재 배치된 서명 전체 비우기 (원하는 품목/칸만 직접 선택 시작)
   */
  const handleClearAllAssignments = () => {
    commitAssignmentsWithHistory(new Map());
    setToast({ msg: '모든 서명 배치를 초기화했습니다. (Ctrl+Z로 되돌리기 가능)', type: 'info' });
  };

  /**
   * 0. 서명 배치 위치 모드 전환 핸들러 (서명 열 전용 vs 1 표시 셀 전체)
   */
  const handlePlacementModeChange = (mode: SignaturePlacementMode) => {
    setState(prev => {
      if (!prev.sheetData) return { ...prev, placementMode: mode };
      const updatedAssignments = prev.signatures.size > 0
        ? autoMatchSignatures(prev.sheetData, prev.signatures, {
            variationStrength,
            scaleMultiplier: prev.signatureScalePreset,
            placementMode: mode,
          })
        : prev.assignments;
      return {
        ...prev,
        placementMode: mode,
        assignments: updatedAssignments,
      };
    });
    setToast({
      msg: mode === 'auto'
        ? "배치 모드 변경: 품목에 1번 + 서명 열에 1번 (2번 동시 서명)"
        : mode === 'signatureColumnOnly'
          ? "배치 모드 변경: 전용 '서명' 열에만 1번 배치"
          : "배치 모드 변경: '1' 표시 셀에만 배치",
      type: 'info',
    });
  };

  /**
   * 0-2. 서명 여백 제거 및 잉크 선명도 보정 토글
   */
  const handleEnhanceSignatureInkChange = (value: boolean) => {
    setState(prev => ({ ...prev, enhanceSignatureInk: value }));
    setToast({
      msg: value ? '✨ 서명 여백 자동 제거 · 진하게 보정 활성화' : '서명 원본 명암 유지',
      type: 'info',
    });
  };

  /**
   * 1. 주민등록번호 뒷자리 마스킹 토글 핸들러 (키퍼 - 보안)
   */
  const handleMaskResidentNumberChange = (value: boolean) => {
    setState(prev => ({ ...prev, maskResidentNumber: value }));
    setToast({
      msg: value ? '🛡️ 주민번호 뒷자리 마스킹 활성화' : '주민번호 마스킹 해제',
      type: 'info',
    });
  };

  /**
   * 2. 서명 크기 프리셋 변경 핸들러 (루미 - 디자인)
   */
  const handleSignatureScalePresetChange = (newScale: number) => {
    setState(prev => {
      const updatedAssignments = new Map<string, SignatureAssignment>();
      for (const [key, assignment] of prev.assignments.entries()) {
        updatedAssignments.set(key, {
          ...assignment,
          scale: newScale,
        });
      }
      return {
        ...prev,
        signatureScalePreset: newScale,
        assignments: updatedAssignments,
      };
    });
  };

  /**
   * 3. 다중 워크시트 선택 전환 핸들러 (코어 - 기술개발)
   */
  const handleSelectSheet = async (sheetIdx: number) => {
    if (!state.excelBuffer || sheetIdx === state.activeSheetIndex) return;

    try {
      setProcessing(true);
      const parseResult = await parseExcelFile(state.excelBuffer, state.excelFile?.name, sheetIdx);
      const newSheetData = parseResult;
      
      const newAssignments = state.signatures.size > 0
        ? autoMatchSignatures(newSheetData, state.signatures, {
            variationStrength,
            scaleMultiplier: state.signatureScalePreset,
            placementMode: state.placementMode,
          })
        : new Map();

      setState(prev => ({
        ...prev,
        sheetData: newSheetData,
        activeSheetIndex: sheetIdx,
        assignments: newAssignments,
      }));

      setToast({ msg: `시트 전환: ${newSheetData.name} (${newSheetData.rows.length}개 행)`, type: 'info' });
    } catch (sheetErr) {
      console.error('Sheet switch error:', sheetErr);
      setError('시트 전환 중 오류가 발생했습니다.');
    } finally {
      setProcessing(false);
    }
  };

  /**
   * 4. 현장 즉석 수기 서명 패드 저장 핸들러 (아틀라스 - 확장성)
   */
  const handleSaveInstantSignature = async (file: File, workerName: string) => {
    await processSignatureFiles([file]);

    // 해당 근로자가 현재 시트에 있으면 자동으로 즉시 서명 배정
    if (state.sheetData) {
      setTimeout(() => {
        if (!state.sheetData) return;
        const updatedAssignments = autoMatchSignatures(state.sheetData, signaturesRef.current, {
          variationStrength,
          scaleMultiplier: state.signatureScalePreset,
          placementMode: state.placementMode,
        });
        setState(prev => ({
          ...prev,
          assignments: updatedAssignments,
        }));
        setToast({ msg: `✅ ${workerName}님의 즉석 서명이 등록 및 반영되었습니다`, type: 'success' });
      }, 150);
    }
  };

  /**
   * 5. 외국인 근로자 및 미매칭 서명 수동 연결 적용 (나비 - 사용성)
   */
  const handleApplyManualMatches = (manualMatches: { rowNumber: number; signatureBaseName: string; signatureVariantId: string }[]) => {
    if (!state.sheetData) return;

    const nextAssignments = new Map(state.assignments);
    const workerItemsMap = new Map(workerItems.map(w => [w.rowNumber, w]));

    for (const match of manualMatches) {
      const worker = workerItemsMap.get(match.rowNumber);
      if (!worker) continue;

      for (const cell of worker.placeholderCells) {
        const key = `${cell.row}:${cell.col}`;
        nextAssignments.set(key, {
          row: cell.row,
          col: cell.col,
          signatureBaseName: match.signatureBaseName,
          signatureVariantId: match.signatureVariantId,
          rotation: 0,
          scale: state.signatureScalePreset,
          offsetX: 0,
          offsetY: 0,
        });
      }
    }

    setState(prev => ({
      ...prev,
      assignments: nextAssignments,
    }));

    setToast({ msg: `✅ ${manualMatches.length}명의 서명이 수동 연결되었습니다`, type: 'success' });
  };

  const handleOpenSignaturePadForWorker = (workerName: string) => {
    setSignaturePadTargetWorker(workerName);
    setIsSignaturePadOpen(true);
  };

  const handleExport = async (isRetry: boolean = false) => {
    if (!state.excelBuffer || !state.sheetData) return;
    
    if (state.assignments.size === 0) {
      setError("배치된 서명이 없습니다. 자동 매칭을 수행해주세요.");
      return;
    }

    setProcessing(true);
    const startTime = performance.now();
    const errorId = `EXP-${Date.now().toString(36).toUpperCase()}`;
    
    // 콘솔 로그 활성화
    const originalLog = console.log;
    const logBuffer: string[] = [];
    console.log = (...args) => {
      originalLog(...args);
      logBuffer.push(args.map(a => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(' '));
    };

    try {
      let assignmentsToUse = state.assignments;
      if (isRetry && state.sheetData) {
        assignmentsToUse = autoMatchSignatures(state.sheetData, state.signatures, {
          variationStrength,
          scaleMultiplier: state.signatureScalePreset,
          placementMode: state.placementMode,
        });
        setState(prev => ({ ...prev, assignments: assignmentsToUse }));
      }

      console.log(`========== [내보내기 시작] ==========`);
      console.log(`형식: ${exportFormat.toUpperCase()}`);
      console.log(`원본 버퍼 크기: ${state.excelBuffer.byteLength} bytes`);
      console.log(`서명 배치 수: ${assignmentsToUse.size}`);
      console.log(`업로드된 서명: ${state.signatures.size}명`);
      
      // Extract HH:MM:SS from ISO timestamp (format: "2024-01-01T14:30:25.123Z")
      // slice(11, 19) extracts the time portion, then replace colons for filename safety
      const ISO_TIME_START = 11; // Position of hours in ISO string
      const ISO_TIME_END = 19;   // Position after seconds in ISO string
      const timestamp = new Date().toISOString().slice(ISO_TIME_START, ISO_TIME_END).replace(/:/g,'');
      const baseFilename = state.excelFile?.name.replace(/\.xlsx$/i, '') || 'output';
      
      if (exportFormat === 'pdf') {
        // PDF 내보내기
        const filename = `서명완료_${timestamp}_${baseFilename}.pdf`;
        await exportToPDF(state.excelBuffer, assignmentsToUse, state.signatures, filename, {
          maskResidentNumber: state.maskResidentNumber,
          targetSheetIndex: state.activeSheetIndex,
        });
        
        const elapsed = performance.now() - startTime;
        console.log(`========== [내보내기 결과] ==========`);
        console.log(`PDF 생성 완료`);
        console.log(`소요 시간: ${elapsed.toFixed(1)}ms`);
        
        setState(prev => ({ ...prev, step: 'export' }));
        setToast({ msg: `✅ PDF 파일이 생성되었습니다: ${filename}`, type: 'success' });
        setError(null);
      } else if (exportFormat === 'png') {
        // PNG 내보내기
        const filename = `서명완료_${timestamp}_${baseFilename}.png`;
        await exportToPNG(state.excelBuffer, assignmentsToUse, state.signatures, filename, {
          maskResidentNumber: state.maskResidentNumber,
          targetSheetIndex: state.activeSheetIndex,
        });
        
        const elapsed = performance.now() - startTime;
        console.log(`========== [내보내기 결과] ==========`);
        console.log(`PNG 생성 완료`);
        console.log(`소요 시간: ${elapsed.toFixed(1)}ms`);
        
        setState(prev => ({ ...prev, step: 'export' }));
        setToast({ msg: `✅ PNG 이미지가 생성되었습니다: ${filename}`, type: 'success' });
        setError(null);
      } else {
        // Excel 내보내기 (기본)
        const blob = await generateFinalExcel(state.excelBuffer, assignmentsToUse, state.signatures, {
          maskResidentNumber: state.maskResidentNumber,
          targetSheetIndex: state.activeSheetIndex,
          scaleMultiplier: state.signatureScalePreset,
          enhanceSignatureInk: state.enhanceSignatureInk,
        });
        
        const elapsed = performance.now() - startTime;
        console.log(`========== [내보내기 결과] ==========`);
        console.log(`생성 파일 크기: ${blob.size} bytes`);
        console.log(`소요 시간: ${elapsed.toFixed(1)}ms`);
        
        if (!blob || blob.size === 0) {
          throw new Error("생성된 파일이 비어있습니다 (0 bytes)");
        }

        if (blob.size < 100) {
          console.error(`❌ [실패] 파일 크기 이상: ${blob.size} bytes - 파일이 손상됨`);
          console.error(`디버그 로그:\n${logBuffer.join('\n')}`);
          throw new Error(`생성된 파일이 너무 작습니다 (${blob.size} bytes). 아래 디버그 정보를 확인해주세요.\n\n${logBuffer.slice(-5).join('\n')}`);
        }

        // ZIP 파일 검증
        const arrayBuffer = await blob.arrayBuffer();
        const view = new Uint8Array(arrayBuffer);
        const isZip = view.length > 1 && view[0] === 0x50 && view[1] === 0x4b;
        console.log(`ZIP 형식 검증: ${isZip ? '✓ 정상' : '✗ 비정상'}`);

        const url = URL.createObjectURL(blob);
        trackObjectUrl(url);
        console.log(`Object URL 생성: ${url.substring(0, 50)}...`);
        
        const a = document.createElement('a');
        a.href = url;
        
        const filename = `서명완료_${timestamp}_${state.excelFile?.name || 'output.xlsx'}`;
        
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        
        console.log(`다운로드 시작: ${filename}`);
        console.log(`========== [완료] ==========\n`);
        
        setTimeout(() => {
          revokeTrackedObjectUrl(url);
          console.log(`메모리 정리: Object URL 해제`);
        }, 100);
        
        setState(prev => ({ ...prev, step: 'export' }));
        setToast({ msg: `✅ 파일이 생성되었습니다: ${filename}\n(${(blob.size / 1024).toFixed(1)}KB)`, type: 'success' });
        setError(null);
      }
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "알 수 없는 오류";
      console.error(`========== [오류 발생] ==========`);
      console.error(`에러 ID: ${errorId}`);
      console.error(`에러 메시지: ${errorMsg}`);
      console.error(`스택:\n${err instanceof Error ? err.stack : '없음'}`);
      console.error(`========== [디버그 로그] ==========`);
      console.error(logBuffer.join('\n'));
      console.error(`========================================\n`);

      const userMessage = getUserFacingExportError(errorMsg, exportFormat);
      setError(`❌ 파일 생성 실패 (${exportFormat.toUpperCase()})\n\n${userMessage}\n\n문의/재현 확인용 오류 ID: ${errorId}`);
    } finally {
      console.log = originalLog;
      setProcessing(false);
    }
  };

  /**
   * 다중 무작위 버전 일괄 생성 + ZIP 1회 다운로드
   * - 브라우저 다중 다운로드 차단을 피하기 위해 JSZip으로 단일 파일화
   * - 각 회차마다 autoMatchSignatures를 다시 호출해 완전 독립 랜덤 결과를 생성
   */
  const handleBatchZipExport = async () => {
    if (!state.excelBuffer || !state.sheetData || !state.excelFile) {
      setError('엑셀 파일이 준비되지 않았습니다.');
      return;
    }

    if (state.signatures.size === 0) {
      setError('서명 이미지가 없습니다.');
      return;
    }

    const total = Math.max(1, Math.min(50, batchCount));
    const abortController = new AbortController();
    batchAbortRef.current = abortController;

    setProcessing(true);
    setBatchProgress({ current: 0, total, phase: 'generate', percent: 0 });
    setError(null);

    try {
      const zipBlob = await buildBatchExcelZip({
        originalBuffer: state.excelBuffer,
        sheetData: state.sheetData,
        signatures: state.signatures,
        sourceFileName: state.excelFile.name,
        count: total,
        variationStrength,
        maskResidentNumber: state.maskResidentNumber,
        scaleMultiplier: state.signatureScalePreset,
        targetSheetIndex: state.activeSheetIndex,
        placementMode: state.placementMode,
        enhanceSignatureInk: state.enhanceSignatureInk,
        baseAssignments: state.assignments,
        onProgress: setBatchProgress,
        signal: abortController.signal,
      });

      const zipUrl = URL.createObjectURL(zipBlob);
      trackObjectUrl(zipUrl);
      const link = document.createElement('a');
      link.href = zipUrl;
      link.download = `서명완료_${state.excelFile.name.replace(/\.xlsx$/i, '')}_${total}부.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      revokeTrackedObjectUrl(zipUrl);

      setBatchProgress({ current: total, total, phase: 'done', percent: 100 });
      setToast({ msg: `✅ ${total}개 파일을 ZIP으로 생성했습니다.`, type: 'success' });
    } catch (err) {
      if (err instanceof DOMException && err.name === 'AbortError') {
        setToast({ msg: '⏹️ 일괄 생성 작업이 취소되었습니다.', type: 'info' });
        setBatchProgress(null);
        return;
      }

      const msg = err instanceof Error ? err.message : '알 수 없는 오류';
      console.error('Batch ZIP export error:', err);
      setError(`일괄 생성 중 오류가 발생했습니다: ${msg}`);
    } finally {
      batchAbortRef.current = null;
      setTimeout(() => setBatchProgress(null), 700);
      setProcessing(false);
    }
  };

  /**
   * 현재 진행 중인 배치 작업을 즉시 중단
   */
  const handleCancelBatchExport = () => {
    if (batchAbortRef.current) {
      batchAbortRef.current.abort();
    }
  };

  const cleanupBlobUrls = (signatures: Map<string, SignatureFile[]>) => {
    signatures.forEach(list => {
      list.forEach(s => {
        revokeTrackedObjectUrl(s.previewUrl);
      });
    });
  };

  /**
   * Start Over: 완전 초기화 함수
   *
   * 동작 순서:
   * 1) 진행 중인 배치/비동기 작업 중단
   * 2) 시그니처 및 다운로드에 사용된 Object URL 전부 해제
   * 3) React 상태를 초기값으로 되돌림(파일/서명/프리뷰/진행률/옵션)
   * 4) Step 상태를 업로드 화면(첫 단계)으로 복귀
   */
  const startOver = () => {
    if (batchAbortRef.current) {
      batchAbortRef.current.abort();
      batchAbortRef.current = null;
    }

    setState(prev => ({
      ...getInitialState(),
      // 서명 보관함은 유지하여 다음 엑셀 양식을 3초 만에 연속 작업할 수 있도록 함
      signatures: prev.signatures,
    }));
    setHistoryPast([]);
    setHistoryFuture([]);
    setPreviewModel(null);
    setPreviewLoading(false);
    setBatchProgress(null);
    setProcessing(false);
    setVariationStrength(70);
    setBatchCount(5);
    setExportFormat('excel');
    setError(null);
    setToast({ msg: '새 엑셀 양식을 올리면 보관된 서명으로 바로 작업할 수 있습니다.', type: 'info' });

    if (excelInputRef.current) excelInputRef.current.value = '';
    if (sigInputRef.current) sigInputRef.current.value = '';
    if (sigFolderInputRef.current) sigFolderInputRef.current.value = '';
  };

  const handleReset = () => {
    startOver();
  };

  const handleBackToPreview = () => {
     setState(prev => ({ ...prev, step: 'preview' }));
  };

  // --- Render Steps ---

  const renderUploadStep = () => (
    <div className="w-full max-w-5xl mx-auto mt-8 px-4 pb-12 space-y-8">
      {/* Hero Banner */}
      <div className="text-center space-y-3">
        <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-indigo-50 border border-indigo-200/80 text-indigo-700 text-xs font-semibold">
          <Sparkles size={13} /> 건설·안전 현장 서명 자동화 스튜디오 v2.0
        </div>
        <h2 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
          엑셀 양식 그대로, 수기 서명 1초 자동화
        </h2>
        <p className="text-sm text-slate-500 max-w-2xl mx-auto">
          원본 로고·테두리·서식을 100% 무손실 보존하며, 엑셀에 '1'을 적지 않아도 화면에서 원하는 품목과 근로자를 클릭해 서명할 수 있습니다.
        </p>
        <div className="flex flex-wrap items-center justify-center gap-2 pt-1 text-[11px] font-medium text-slate-600">
          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white border border-slate-200 shadow-sm">
            <ShieldCheck size={13} className="text-emerald-600" /> OOXML 무손실 원본 보존
          </span>
          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white border border-slate-200 shadow-sm">
            <HardDrive size={13} className="text-indigo-600" /> 브라우저 서명 보관함 자동 저장
          </span>
          <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white border border-slate-200 shadow-sm">
            <Lock size={13} className="text-slate-600" /> 서버 전송 없음 (100% 로컬 처리)
          </span>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
      {/* Excel Upload Card (Drag & Drop 지원) */}
      <div 
        onDragOver={(e) => {
          e.preventDefault();
          e.stopPropagation();
          if (!isDraggingExcel) setIsDraggingExcel(true);
        }}
        onDragEnter={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setIsDraggingExcel(true);
        }}
        onDragLeave={(e) => {
          e.preventDefault();
          e.stopPropagation();
          if (!e.currentTarget.contains(e.relatedTarget as Node)) {
            setIsDraggingExcel(false);
          }
        }}
        onDrop={async (e) => {
          e.preventDefault();
          e.stopPropagation();
          setIsDraggingExcel(false);
          const files = await extractFilesFromDataTransfer(e.dataTransfer);
          const excelFile = files.find(f => /\.(xlsx|xls|csv|xlsm|xlsb)$/i.test(f.name));
          if (excelFile) {
            await processExcelFile(excelFile);
          } else if (files.length > 0 && files.some(f => f.type.startsWith('image/') || /\.(png|jpe?g|webp|bmp)$/i.test(f.name))) {
            setToast({ msg: '서명 이미지를 감지했습니다. 오른쪽 2번 서명 이미지 카드에 넣어주세요.', type: 'info' });
          } else {
            setError('엑셀 파일(.xlsx, .xls)을 끌어다 놓아주세요.');
          }
        }}
        className={`relative bg-white p-8 rounded-2xl shadow-lg border-2 transition-all duration-200 ${
          isDraggingExcel 
            ? 'border-blue-500 bg-blue-50/70 scale-[1.02] ring-4 ring-blue-100' 
            : state.excelFile 
              ? 'border-green-500' 
              : 'border-gray-100 hover:border-blue-300'
        }`}
      >
        <div className="flex flex-col items-center text-center space-y-4">
          <div className={`p-4 rounded-full transition-transform ${isDraggingExcel ? 'scale-110' : ''} ${state.excelFile ? 'bg-green-100 text-green-600' : 'bg-blue-50 text-blue-600'}`}>
            <FileSpreadsheet size={48} />
          </div>
          <h3 className="text-xl font-bold text-gray-800">1. 엑셀 양식 업로드</h3>
          <p className="text-gray-500 text-sm">
            '성명'과 '1' 표시가 포함된 안전관리비/보호구 지급대장 엑셀 파일(.xlsx, .xls)을 업로드하세요.
          </p>
          
          <label className="cursor-pointer bg-gray-900 text-white px-6 py-3 rounded-lg hover:bg-gray-800 transition-colors w-full">
            <span className="flex items-center justify-center gap-2">
              <Upload size={18} />
              {state.excelFile ? '파일 변경' : '엑셀 파일 선택'}
            </span>
            <input 
              key={`excel-input-${state.step}`} // Force reset input when step changes
              ref={excelInputRef}
              type="file" 
              accept=".xlsx,.xls,.csv,.xlsm,.xlsb,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" 
              className="hidden" 
              onChange={handleExcelUpload} 
            />
          </label>

          {/* Drag & Drop 유도 영역 */}
          <div className={`w-full py-2.5 px-3 border border-dashed rounded-lg text-xs font-medium transition-colors ${
            isDraggingExcel ? 'border-blue-500 text-blue-700 bg-blue-100/60' : 'border-gray-200 text-gray-500 bg-gray-50'
          }`}>
            {isDraggingExcel ? '🚀 여기에 엑셀 파일을 놓으세요!' : '🖱️ 탐색기에서 엑셀 파일을 여기로 끌어다 놓으셔도 됩니다.'}
          </div>
          
          {state.excelFile && (
            <div className="flex items-center gap-2 text-green-600 text-sm font-medium">
              <CheckCircle size={16} />
              {state.excelFile.name} 로드됨
            </div>
          )}
        </div>
      </div>

      {/* Signature Upload Card (Drag & Drop + 폴더 지원) */}
      <div 
        onDragOver={(e) => {
          e.preventDefault();
          e.stopPropagation();
          if (!isDraggingSig) setIsDraggingSig(true);
        }}
        onDragEnter={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setIsDraggingSig(true);
        }}
        onDragLeave={(e) => {
          e.preventDefault();
          e.stopPropagation();
          if (!e.currentTarget.contains(e.relatedTarget as Node)) {
            setIsDraggingSig(false);
          }
        }}
        onDrop={async (e) => {
          e.preventDefault();
          e.stopPropagation();
          setIsDraggingSig(false);
          const files = await extractFilesFromDataTransfer(e.dataTransfer);
          
          const excelFile = files.find(f => /\.(xlsx|xls|csv|xlsm|xlsb)$/i.test(f.name));
          if (excelFile && files.length === 1) {
            setToast({ msg: '엑셀 파일을 감지했습니다. 왼쪽 1번 엑셀 양식 카드에 넣어주세요.', type: 'info' });
            return;
          }

          if (files.length === 0) {
            setError('폴더가 비어있거나 읽을 수 있는 이미지 파일이 없습니다.');
            return;
          }

          await processSignatureFiles(files);
        }}
        className={`relative bg-white p-8 rounded-2xl shadow-lg border-2 transition-all duration-200 ${
          isDraggingSig 
            ? 'border-purple-500 bg-purple-50/70 scale-[1.02] ring-4 ring-purple-100' 
            : state.signatures.size > 0 
              ? 'border-green-500' 
              : 'border-gray-100 hover:border-purple-300'
        }`}
      >
        <div className="flex flex-col items-center text-center space-y-4">
          <div className={`p-4 rounded-full transition-transform ${isDraggingSig ? 'scale-110' : ''} ${state.signatures.size > 0 ? 'bg-green-100 text-green-600' : 'bg-purple-50 text-purple-600'}`}>
            <ImageIcon size={48} />
          </div>
          <h3 className="text-xl font-bold text-gray-800">2. 서명 이미지 / 폴더 업로드</h3>
          <p className="text-gray-500 text-sm">
            수기 서명 이미지 파일(예: 홍길동_1.png) 또는 <strong>서명 폴더를 통째로</strong> 추가하세요.
          </p>
          
          {/* 개별 파일 추가 & 폴더 통째로 선택 버튼 그룹 */}
          <div className="grid grid-cols-2 gap-2 w-full">
            <label className="cursor-pointer bg-gray-900 text-white px-3 py-3 rounded-lg hover:bg-gray-800 transition-colors text-center">
              <span className="flex items-center justify-center gap-1.5 text-sm font-medium">
                <Upload size={16} />
                이미지 파일
              </span>
              <input 
                key={`sig-input-${state.step}`} // Force reset input when step changes
                ref={sigInputRef}
                type="file" 
                accept="image/*" 
                multiple 
                className="hidden" 
                onChange={handleSignatureUpload} 
              />
            </label>

            <label className="cursor-pointer bg-purple-700 text-white px-3 py-3 rounded-lg hover:bg-purple-800 transition-colors text-center">
              <span className="flex items-center justify-center gap-1.5 text-sm font-medium">
                <FolderUp size={16} />
                폴더 통째로
              </span>
              <input 
                key={`sig-folder-input-${state.step}`}
                ref={sigFolderInputRef}
                type="file" 
                {...({ webkitdirectory: "", directory: "" } as any)}
                multiple 
                className="hidden" 
                onChange={handleFolderUpload} 
              />
            </label>

            <button
              type="button"
              onClick={() => {
                setSignaturePadTargetWorker('');
                setIsSignaturePadOpen(true);
              }}
              className="col-span-2 cursor-pointer bg-indigo-50 border border-indigo-200 text-indigo-700 px-3 py-2.5 rounded-lg hover:bg-indigo-100 transition-colors text-center flex items-center justify-center gap-1.5 text-xs font-semibold"
            >
              <PenTool size={14} /> 현장 즉석 수기 서명 패드 (Canvas)
            </button>
          </div>

          {/* Drag & Drop 유도 영역 */}
          <div className={`w-full py-2.5 px-3 border border-dashed rounded-lg text-xs font-medium transition-colors ${
            isDraggingSig ? 'border-purple-500 text-purple-700 bg-purple-100/60' : 'border-gray-200 text-gray-500 bg-gray-50'
          }`}>
            {isDraggingSig ? '🚀 여기에 서명 폴더나 이미지들을 놓으세요!' : '📁 탐색기에서 서명 폴더를 통째로 여기로 끌어다 놓으셔도 됩니다.'}
          </div>
          
          {state.signatures.size > 0 && (
            <div className="w-full pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-1.5 text-emerald-700 text-xs font-semibold">
                <CheckCircle size={15} />
                <span>
                  {state.signatures.size}명 확인됨 ({Array.from(state.signatures.values()).flat().length}개 파일)
                  {vaultRestoredCount > 0 && ' · 보관함 자동 복원'}
                </span>
              </div>
              <button
                type="button"
                onClick={handleClearVault}
                className="inline-flex items-center gap-1 px-2 py-1 rounded-md text-[11px] font-medium text-rose-600 hover:bg-rose-50 transition-colors"
                title="브라우저에 자동 저장된 서명 보관함을 비웁니다"
              >
                <Trash2 size={12} /> 보관함 비우기
              </button>
            </div>
          )}
        </div>
      </div>
      </div>

      <div className="flex justify-center pt-2">
        <button 
          onClick={() => runAutoMatch()}
          disabled={!state.excelFile || state.signatures.size === 0 || processing}
          className="bg-indigo-600 text-white px-10 py-4 rounded-2xl text-base font-bold shadow-lg shadow-indigo-200 hover:bg-indigo-700 hover:shadow-xl disabled:opacity-45 disabled:shadow-none disabled:cursor-not-allowed flex items-center gap-3 transition-all"
        >
          {processing ? <RefreshCw className="animate-spin" size={20} /> : <ArrowRight size={20} />}
          자동 매칭 및 미리보기 시작
        </button>
      </div>
    </div>
  );

  const renderPreviewStep = () => {
    if (!state.sheetData) return null;

    return (
      <SignatureWorkspace
        previewLoading={previewLoading}
        previewModel={previewModel}
        processing={processing}
        exportFormat={exportFormat}
        variationStrength={variationStrength}
        batchCount={batchCount}
        batchProgress={batchProgress}
        maskResidentNumber={state.maskResidentNumber}
        signatureScalePreset={state.signatureScalePreset}
        placementMode={state.placementMode}
        detectedSignatureColLabel={detectedSignatureColLabel}
        enhanceSignatureInk={state.enhanceSignatureInk}
        availableSheets={state.availableSheets}
        activeSheetIndex={state.activeSheetIndex}
        unmatchedWorkerCount={unmatchedWorkerCount}
        selectableColumns={selectableColumns}
        workerRowMap={workerRowMap}
        canUndo={historyPast.length > 0}
        canRedo={historyFuture.length > 0}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onToggleColumnSignatures={handleToggleColumnSignatures}
        onToggleCellSignature={handleToggleCellSignature}
        onClearAllAssignments={handleClearAllAssignments}
        onVariationStrengthChange={setVariationStrength}
        onBatchCountChange={(value) => setBatchCount(Math.max(1, Math.min(50, value || 1)))}
        onExportFormatChange={setExportFormat}
        onMaskResidentNumberChange={handleMaskResidentNumberChange}
        onSignatureScalePresetChange={handleSignatureScalePresetChange}
        onPlacementModeChange={handlePlacementModeChange}
        onEnhanceSignatureInkChange={handleEnhanceSignatureInkChange}
        onSelectSheet={handleSelectSheet}
        onOpenManualMatch={() => setIsManualMatchOpen(true)}
        onOpenSignaturePad={() => {
          setSignaturePadTargetWorker('');
          setIsSignaturePadOpen(true);
        }}
        onAutoMatch={() => runAutoMatch()}
        onSingleExport={() => handleExport(false)}
        onBatchZipExport={handleBatchZipExport}
        onCancelBatchExport={handleCancelBatchExport}
        isBatchCancelable={!!batchAbortRef.current && processing}
        onStartOver={handleReset}
        assignmentCount={state.assignments.size}
        rowCount={state.sheetData.rows.length}
      />
    );
  };

  const renderExportStep = () => (
    <div className="flex flex-col items-center justify-center h-[60vh] space-y-6">
      <div className="bg-green-100 p-6 rounded-full text-green-600 mb-4 animate-bounce-slow">
        <CheckCircle size={64} />
      </div>
      <h2 className="text-3xl font-bold text-gray-800">내보내기 완료!</h2>
      <div className="inline-flex items-center gap-2 px-4 py-1.5 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-700 text-xs font-semibold">
        <ShieldCheck size={15} />
        OOXML 표준 무결성 검증 완료 · 원본 서식 100% 보존 · 서명 {state.assignments.size}건 반영
      </div>
      <p className="text-gray-500 max-w-md text-center">
        파일 다운로드가 시작되었습니다.<br/>
        다른 버전(새로운 무작위 서명 배치)이 필요하시면 아래 버튼을 눌러주세요.
      </p>
      <div className="flex gap-4 mt-8">
        <button 
          onClick={handleBackToPreview}
          className="bg-white border border-gray-300 text-gray-700 px-6 py-3 rounded-xl font-medium hover:bg-gray-50 flex items-center gap-2"
        >
          <Settings size={18} />
          편집 화면으로 돌아가기
        </button>
        <button 
          onClick={() => handleExport(true)}
          className="bg-indigo-600 text-white px-6 py-3 rounded-xl font-medium hover:bg-indigo-700 flex items-center gap-2 shadow-lg"
        >
          <Copy size={18} />
          다른 랜덤 버전 즉시 다운로드
        </button>
        <button
          onClick={handleReset}
          className="bg-white border border-slate-300 text-slate-700 px-6 py-3 rounded-xl font-medium hover:bg-slate-50 flex items-center gap-2"
        >
          <RotateCcw size={18} />
          처음으로 돌아가기
        </button>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans">
      {/* Toast */}
      {toast && (
        <div className="fixed top-20 right-4 z-[100] bg-gray-800 text-white px-6 py-3 rounded-lg shadow-xl flex items-center gap-3 animate-fade-in-down">
          {toast.type === 'success' ? <CheckCircle size={20} className="text-green-400"/> : <Settings size={20} className="text-blue-400"/>}
          {toast.msg}
        </div>
      )}

      {/* Header */}
      <header className="bg-white/90 backdrop-blur-md border-b border-slate-200/80 sticky top-0 z-50">
        <div className="max-w-[1600px] mx-auto px-4 h-16 flex items-center justify-between gap-4">
          <div className="flex items-center gap-2.5">
            <div className="bg-gradient-to-br from-indigo-600 to-violet-600 text-white p-1.5 rounded-xl shadow-sm">
              <CheckCircle size={20} strokeWidth={2.5} />
            </div>
            <div>
              <h1 className="text-lg font-extrabold tracking-tight text-slate-900 leading-none">
                SMART<span className="text-indigo-600">SIGN</span>PRO
              </h1>
              <span className="text-[10px] font-semibold text-slate-400 tracking-wide">
                PSI AI STUDIO · v2.0
              </span>
            </div>
          </div>

          {/* 3-Step Visual Progress Stepper */}
          <nav className="hidden md:flex items-center gap-2 text-xs font-semibold">
            {[
              { id: 'upload', num: 1, label: '파일·서명 준비' },
              { id: 'preview', num: 2, label: '서명 배치·미리보기' },
              { id: 'export', num: 3, label: '무손실 내보내기' },
            ].map((s, idx) => {
              const stepOrder = { upload: 0, preview: 1, export: 2 };
              const curIdx = stepOrder[state.step];
              const isActive = state.step === s.id;
              const isDone = curIdx > idx;
              return (
                <React.Fragment key={s.id}>
                  <div
                    className={`flex items-center gap-1.5 px-3 py-1.5 rounded-full transition-colors ${
                      isActive
                        ? 'bg-indigo-600 text-white shadow-sm'
                        : isDone
                          ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                          : 'bg-slate-100 text-slate-400'
                    }`}
                  >
                    <span className={`w-4 h-4 rounded-full flex items-center justify-center text-[10px] font-bold ${
                      isActive ? 'bg-white text-indigo-600' : isDone ? 'bg-emerald-600 text-white' : 'bg-slate-300 text-white'
                    }`}>
                      {s.num}
                    </span>
                    <span>{s.label}</span>
                  </div>
                  {idx < 2 && <ArrowRight size={13} className="text-slate-300" />}
                </React.Fragment>
              );
            })}
          </nav>

          <div className="flex items-center gap-3">
            <span className="hidden lg:inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 text-[11px] font-medium">
              <Lock size={11} className="text-emerald-600" /> 100% 로컬 처리
            </span>
            <button 
              onClick={() => setShowGuide(true)}
              className="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-indigo-50 text-slate-700 hover:text-indigo-700 font-semibold text-xs flex items-center gap-1.5 transition-colors"
            >
              <HelpCircle size={15} />
              이용 가이드
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="w-full">
        {error && (
          <div className="max-w-4xl mx-auto mt-6 bg-red-50 border-l-4 border-red-500 text-red-800 px-6 py-4 rounded flex items-start gap-4 shadow-sm">
            <AlertCircle size={24} className="flex-shrink-0 mt-0.5" />
            <div className="flex-1">
              <p className="font-semibold text-base mb-1">오류 발생</p>
              <p className="text-sm whitespace-pre-wrap">{error}</p>
            </div>
            <button 
              onClick={() => setError(null)} 
              className="ml-auto text-red-600 hover:text-red-800 flex-shrink-0"
            >
              <X size={20} />
            </button>
          </div>
        )}

        {state.step === 'upload' && renderUploadStep()}
        {state.step === 'preview' && renderPreviewStep()}
        {state.step === 'export' && renderExportStep()}
      </main>

      {/* Loading Overlay */}
      {processing && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center backdrop-blur-sm">
          <div className="bg-white p-8 rounded-2xl shadow-2xl flex flex-col items-center">
            <RefreshCw className="animate-spin text-indigo-600 mb-4" size={48} />
            <h3 className="text-lg font-semibold text-gray-900">처리 중...</h3>
            <p className="text-sm text-gray-500 mt-2">
              {state.step === 'upload' && '파일을 분석하고 있습니다'}
              {state.step === 'preview' && '서명을 무작위로 배치하고 있습니다'}
              {state.step === 'export' && '엑셀 파일을 생성하고 있습니다'}
            </p>
            <div className="w-48 h-1 bg-gray-200 rounded-full mt-4 overflow-hidden">
              <div className="h-full bg-indigo-600 animate-pulse"></div>
            </div>
          </div>
        </div>
      )}

      {/* Guide Modal */}
      {showGuide && (
        <div className="fixed inset-0 bg-black/60 z-[60] flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl overflow-hidden max-h-[90vh] flex flex-col">
            <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-gray-50">
              <h3 className="text-xl font-bold text-gray-900 flex items-center gap-2">
                <HelpCircle className="text-indigo-600" /> 이용 가이드 및 워크플로우
              </h3>
              <button onClick={() => setShowGuide(false)} className="text-gray-400 hover:text-gray-700">
                <X size={24} />
              </button>
            </div>
            
            <div className="p-0 overflow-y-auto custom-scrollbar bg-gray-50">
              <div className="bg-white p-8 border-b">
                <div className="flex flex-col md:flex-row justify-between items-center gap-4 text-center">
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-blue-100 rounded-2xl flex items-center justify-center text-blue-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <FileSpreadsheet size={32} />
                    </div>
                    <div className="font-bold text-gray-800">1. 엑셀 업로드</div>
                  </div>
                  <ArrowRight className="text-gray-300 hidden md:block" />
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-purple-100 rounded-2xl flex items-center justify-center text-purple-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <ImageIcon size={32} />
                    </div>
                    <div className="font-bold text-gray-800">2. 서명 업로드</div>
                  </div>
                  <ArrowRight className="text-gray-300 hidden md:block" />
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-indigo-100 rounded-2xl flex items-center justify-center text-indigo-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <Settings size={32} />
                    </div>
                    <div className="font-bold text-gray-800">3. 자동 매칭</div>
                  </div>
                  <ArrowRight className="text-gray-300 hidden md:block" />
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-green-100 rounded-2xl flex items-center justify-center text-green-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <Download size={32} />
                    </div>
                    <div className="font-bold text-gray-800">4. 엑셀 다운로드</div>
                  </div>
                </div>
              </div>
              <div className="p-8 space-y-8">
                <div className="space-y-4">
                  <h4 className="font-bold text-gray-900">📝 서명 파일명 규칙 (필수)</h4>
                  <p className="text-sm text-gray-600">서명 이미지 파일명은 반드시 <span className="bg-gray-200 px-2 py-0.5 rounded font-mono">사람이름_숫자.확장자</span> 형태여야 합니다.</p>
                  <ul className="text-sm text-gray-600 space-y-1 ml-4">
                    <li>✅ 올바른 예: <code className="bg-green-100 px-1 text-green-800">홍길동_1.png</code>, <code className="bg-green-100 px-1 text-green-800">김철수_2.jpg</code></li>
                    <li>❌ 잘못된 예: <code className="bg-red-100 px-1 text-red-800">홍길동.png</code>, <code className="bg-red-100 px-1 text-red-800">signature.jpg</code></li>
                  </ul>
                </div>

                <div className="space-y-4">
                  <h4 className="font-bold text-gray-900">📊 엑셀 파일 형식 요구사항</h4>
                  <ul className="text-sm text-gray-600 space-y-2">
                    <li>• <strong>필수 열:</strong> '성명' 또는 '이름' 열이 있어야 합니다</li>
                    <li>• <strong>서명 기호:</strong> 서명이 필요한 곳에 '1', '(1)', '1.', '1)' 중 하나 입력</li>
                    <li>• <strong>파일 형식:</strong> Microsoft Excel (.xlsx) 형식만 지원</li>
                    <li>• <strong>파일 크기:</strong> 5MB 이하 권장 (이미지 포함 시)</li>
                  </ul>
                </div>

                <div className="space-y-4">
                  <h4 className="font-bold text-gray-900">🎯 주요 기능</h4>
                  <ul className="text-sm text-gray-600 space-y-2">
                    <li>• <strong>자동 매칭:</strong> 엑셀의 이름과 서명 파일명을 자동으로 매칭</li>
                    <li>• <strong>무작위 배치:</strong> 같은 사람의 서명이 다른 버전으로 무작위 배치</li>
                    <li>• <strong>다중 버전:</strong> 원본 파일을 보존하고 여러 버전 생성 가능</li>
                    <li>• <strong>스타일 보존:</strong> 원본 엑셀의 모든 포맷팅과 스타일이 유지됨</li>
                  </ul>
                </div>

                <div className="bg-blue-50 p-4 rounded-lg text-sm text-blue-900 space-y-2">
                  <strong>💡 팁:</strong>
                  <ul className="space-y-1">
                    <li>• 같은 사람의 서명이 많을수록 더 자연스러운 무작위 배치가 가능합니다</li>
                    <li>• '다른 랜덤 버전 즉시 다운로드' 버튼으로 빠르게 새 버전을 생성할 수 있습니다</li>
                    <li>• 서명 이미지는 PNG, JPG 형식을 권장합니다</li>
                  </ul>
                </div>
              </div>
            </div>
            
            <div className="p-4 border-t border-gray-100 bg-white flex justify-end">
              <button onClick={() => setShowGuide(false)} className="bg-indigo-600 text-white px-8 py-3 rounded-lg font-medium hover:bg-indigo-700 shadow-lg">알겠습니다</button>
            </div>
          </div>
        </div>
      )}

      {/* 현장 즉석 수기 서명 패드 모달 */}
      <SignaturePadModal
        isOpen={isSignaturePadOpen}
        onClose={() => setIsSignaturePadOpen(false)}
        defaultWorkerName={signaturePadTargetWorker}
        onSave={handleSaveInstantSignature}
      />

      {/* 외국인 근로자 및 미매칭 서명 수동 연결 모달 */}
      <ManualMatchModal
        isOpen={isManualMatchOpen}
        onClose={() => setIsManualMatchOpen(false)}
        workerItems={workerItems}
        signatures={state.signatures}
        maskResidentNumber={state.maskResidentNumber}
        onApplyMatches={handleApplyManualMatches}
        onOpenSignaturePadForWorker={handleOpenSignaturePadForWorker}
      />
    </div>
  );
}