import React, { useState, useEffect, useRef } from 'react';
import { Upload, FileSpreadsheet, Image as ImageIcon, CheckCircle, RotateCcw, Download, Settings, RefreshCw, AlertCircle, HelpCircle, X, ArrowRight, FileText, MousePointer2, Copy, FileDown } from 'lucide-react';
import { parseExcelFile, autoMatchSignatures, generateFinalExcel, normalizeName } from './services/excelService';
import { buildSheetPreviewModel, exportToPDF, exportToPNG, SheetPreviewModel } from './services/alternativeExportService';
import { AppState, SignatureFile, SheetData, SignatureAssignment } from './types';
import SignatureWorkspace from './components/SignatureWorkspace';
import { buildBatchExcelZip, BatchExportProgress } from './services/batchExportService';

// Factory function to ensure fresh state on reset
const getInitialState = (): AppState => ({
  step: 'upload',
  excelFile: null,
  excelBuffer: null,
  sheetData: null,
  signatures: new Map(),
  assignments: new Map(),
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
  
  // Refs to clear file inputs
  const excelInputRef = useRef<HTMLInputElement>(null);
  const sigInputRef = useRef<HTMLInputElement>(null);

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
          true
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
  }, [state.step, state.excelBuffer, state.assignments, state.signatures]);

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

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Validate file type
    if (!file.name.endsWith('.xlsx') && !file.type.includes('spreadsheet')) {
      setError("XLSX 파일만 지원합니다. 파일 확장명을 확인해주세요.");
      if (excelInputRef.current) excelInputRef.current.value = '';
      return;
    }

    // Check file size (Warning if > 5MB)
    if (file.size > 5 * 1024 * 1024) {
      if (!window.confirm(`선택하신 엑셀 파일의 용량이 큽니다 (${(file.size / 1024 / 1024).toFixed(1)}MB).\n파일 내에 이미지가 많거나 행이 매우 많으면 'Out of Memory' 오류가 발생할 수 있습니다.\n\n계속 진행하시겠습니까?`)) {
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

      const sheetData = await parseExcelFile(buffer);
      
      if (sheetData.rows.length === 0) {
        throw new Error("데이터가 없는 파일입니다. 성명 열과 데이터가 포함된 파일을 확인해주세요.");
      }

      setState(prev => ({ ...prev, excelFile: file, excelBuffer: buffer, sheetData, step: 'upload' }));
      setToast({ msg: `${file.name} 로드됨 (${sheetData.rows.length}개 행)`, type: 'success' });
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "알 수 없는 오류";
      setError(`엑셀 파일 읽기 실패: ${errorMsg}`);
      console.error('Excel upload error:', err);
    } finally {
      setProcessing(false);
    }
  };

  const handleSignatureUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setProcessing(true);
    const newSignatures = new Map<string, SignatureFile[]>(state.signatures);
    let count = 0;
    const failedFiles: string[] = [];

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      let objectUrl: string | null = null;
      // Only images
      if (!file.type.startsWith('image/')) {
        failedFiles.push(`${file.name} (이미지 파일이 아님)`);
        continue;
      }

      // File size check (warning if > 2MB per image)
      if (file.size > 2 * 1024 * 1024) {
        failedFiles.push(`${file.name} (2MB 초과 - 압축 권장)`);
        continue;
      }

      try {
        /**
         * 메모리 최적화:
         * - 우선 createImageBitmap으로 크기를 읽어 브라우저 디코더 리소스를 즉시 해제한다.
         * - fallback 경로에서는 임시 Object URL을 만들고 즉시 revoke 한다.
         */
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
    
    const toastMsg = failedFiles.length > 0 
      ? `${count}개 추가됨${failedFiles.length > 0 ? ` (${failedFiles.length}개 제외됨: ${failedFiles.slice(0, 2).join(', ')}${failedFiles.length > 2 ? '...' : ''})` : ''}`
      : `${count}개의 서명이 추가되었습니다.`;
    
    setToast({ msg: toastMsg, type: count > 0 ? 'success' : 'info' });
  };

  const runAutoMatch = () => {
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
        const assignments = autoMatchSignatures(state.sheetData, state.signatures, { variationStrength });
        setState(prev => ({ ...prev, assignments, step: 'preview' }));
        setProcessing(false);
        
        if (assignments.size === 0) {
          setError("매칭된 서명이 없습니다. 엑셀 파일에 '성명' 열과 서명 기호(1)가 있는지 확인해주세요.");
          setToast({ msg: '⚠️ 매칭 실패', type: 'info' });
        } else {
          const signatureCount = new Set(Array.from(assignments.values()).map(a => a.signatureBaseName)).size;
          setToast({ msg: `✅ ${assignments.size}개 위치에 ${signatureCount}명의 서명이 배치되었습니다`, type: 'success' });
        }
    }, 100);
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
        assignmentsToUse = autoMatchSignatures(state.sheetData, state.signatures, { variationStrength });
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
        await exportToPDF(state.excelBuffer, assignmentsToUse, state.signatures, filename);
        
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
        await exportToPNG(state.excelBuffer, assignmentsToUse, state.signatures, filename);
        
        const elapsed = performance.now() - startTime;
        console.log(`========== [내보내기 결과] ==========`);
        console.log(`PNG 생성 완료`);
        console.log(`소요 시간: ${elapsed.toFixed(1)}ms`);
        
        setState(prev => ({ ...prev, step: 'export' }));
        setToast({ msg: `✅ PNG 이미지가 생성되었습니다: ${filename}`, type: 'success' });
        setError(null);
      } else {
        // Excel 내보내기 (기본)
        const blob = await generateFinalExcel(state.excelBuffer, assignmentsToUse, state.signatures);
        
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

    cleanupBlobUrls(state.signatures);
    revokeAllTrackedObjectUrls();

    setState(getInitialState());
    setPreviewModel(null);
    setPreviewLoading(false);
    setBatchProgress(null);
    setProcessing(false);
    setVariationStrength(70);
    setBatchCount(5);
    setExportFormat('excel');
    setError(null);
    setToast({ msg: '새 파일 작업을 시작할 수 있도록 초기화되었습니다.', type: 'info' });

    if (excelInputRef.current) excelInputRef.current.value = '';
    if (sigInputRef.current) sigInputRef.current.value = '';
  };

  const handleReset = () => {
    if (window.confirm("정말로 처음부터 다시 시작하시겠습니까?\n모든 데이터가 초기화됩니다.")) {
      startOver();
    }
  };

  const handleBackToPreview = () => {
     setState(prev => ({ ...prev, step: 'preview' }));
  };

  // --- Render Steps ---

  const renderUploadStep = () => (
    <div className="grid grid-cols-1 md:grid-cols-2 gap-8 w-full max-w-4xl mx-auto mt-10 px-4">
      {/* Excel Upload Card */}
      <div className={`bg-white p-8 rounded-2xl shadow-lg border-2 ${state.excelFile ? 'border-green-500' : 'border-gray-100'}`}>
        <div className="flex flex-col items-center text-center space-y-4">
          <div className={`p-4 rounded-full ${state.excelFile ? 'bg-green-100 text-green-600' : 'bg-blue-50 text-blue-600'}`}>
            <FileSpreadsheet size={48} />
          </div>
          <h3 className="text-xl font-bold text-gray-800">1. 엑셀 양식 업로드</h3>
          <p className="text-gray-500 text-sm">'성명'과 '1' 표시가 포함된 안전관리비/보호구 지급대장 엑셀 파일을 업로드하세요.</p>
          
          <label className="cursor-pointer bg-gray-900 text-white px-6 py-3 rounded-lg hover:bg-gray-800 transition-colors w-full">
            <span className="flex items-center justify-center gap-2">
              <Upload size={18} />
              {state.excelFile ? '파일 변경' : '엑셀 파일 선택'}
            </span>
            <input 
              key={`excel-input-${state.step}`} // Force reset input when step changes
              ref={excelInputRef}
              type="file" 
              accept=".xlsx" 
              className="hidden" 
              onChange={handleExcelUpload} 
            />
          </label>
          
          {state.excelFile && (
            <div className="flex items-center gap-2 text-green-600 text-sm font-medium">
              <CheckCircle size={16} />
              {state.excelFile.name} 로드됨
            </div>
          )}
        </div>
      </div>

      {/* Signature Upload Card */}
      <div className={`bg-white p-8 rounded-2xl shadow-lg border-2 ${state.signatures.size > 0 ? 'border-green-500' : 'border-gray-100'}`}>
        <div className="flex flex-col items-center text-center space-y-4">
          <div className={`p-4 rounded-full ${state.signatures.size > 0 ? 'bg-green-100 text-green-600' : 'bg-purple-50 text-purple-600'}`}>
            <ImageIcon size={48} />
          </div>
          <h3 className="text-xl font-bold text-gray-800">2. 서명 이미지 업로드</h3>
          <p className="text-gray-500 text-sm">모든 수기 서명 이미지 파일을 업로드하세요 (예: 홍길동_1.png).</p>
          
          <label className="cursor-pointer bg-gray-900 text-white px-6 py-3 rounded-lg hover:bg-gray-800 transition-colors w-full">
            <span className="flex items-center justify-center gap-2">
              <Upload size={18} />
              이미지 추가
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
          
          {state.signatures.size > 0 && (
            <div className="flex items-center gap-2 text-green-600 text-sm font-medium">
              <CheckCircle size={16} />
              {state.signatures.size}명의 서명 확인됨 ({Array.from(state.signatures.values()).flat().length}개 파일)
            </div>
          )}
        </div>
      </div>

      <div className="md:col-span-2 flex justify-center mt-4 pb-10">
        <button 
          onClick={runAutoMatch}
          disabled={!state.excelFile || state.signatures.size === 0 || processing}
          className="bg-indigo-600 text-white px-10 py-4 rounded-xl text-lg font-bold shadow-xl hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-3 transition-all"
        >
          {processing ? <RefreshCw className="animate-spin" /> : <Settings />}
          자동 매칭 시작
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
        onVariationStrengthChange={setVariationStrength}
        onBatchCountChange={(value) => setBatchCount(Math.max(1, Math.min(50, value || 1)))}
        onExportFormatChange={setExportFormat}
        onAutoMatch={runAutoMatch}
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
      <header className="bg-white border-b border-gray-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-orange-500 text-white p-1.5 rounded-lg">
              <CheckCircle size={20} strokeWidth={3} />
            </div>
            <h1 className="text-xl font-bold tracking-tight text-gray-900">
              Safety<span className="text-orange-600">Sign</span>Pro
            </h1>
          </div>
          <div className="flex items-center gap-4">
            <button 
              onClick={() => setShowGuide(true)}
              className="text-gray-600 hover:text-indigo-600 font-medium text-sm flex items-center gap-1.5 transition-colors"
            >
              <HelpCircle size={18} />
              이용 가이드
            </button>
            <div className="text-sm text-gray-400 font-medium hidden sm:block">
              개인보호구 지급대장 자동화 v1.0
            </div>
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
    </div>
  );
}