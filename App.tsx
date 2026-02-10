import React, { useState, useEffect, useRef } from 'react';
import { Upload, FileSpreadsheet, Image as ImageIcon, CheckCircle, RotateCcw, Download, Settings, RefreshCw, AlertCircle, HelpCircle, X, ArrowRight, FileText, MousePointer2, Copy } from 'lucide-react';
import { parseExcelFile, autoMatchSignatures, generateFinalExcel, normalizeName } from './services/excelService';
import { AppState, SignatureFile, SheetData, SignatureAssignment } from './types';

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
  const [error, setError] = useState<string | null>(null);
  const [showGuide, setShowGuide] = useState(false);
  const [toast, setToast] = useState<{msg: string, type: 'success' | 'info'} | null>(null);
  
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

      const objectUrl = URL.createObjectURL(file);

      try {
        const getImageDims = () => new Promise<{w: number, h: number}>((resolve) => {
          const img = new Image();
          img.onload = () => resolve({ w: img.width, h: img.height });
          img.onerror = () => {
            URL.revokeObjectURL(objectUrl);
            resolve({ w: 100, h: 50 });
          };
          img.src = objectUrl;
        });

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
          URL.revokeObjectURL(objectUrl);
          continue;
        }
        
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
          URL.revokeObjectURL(objectUrl);
          failedFiles.push(`${file.name} (중복)`);
        }
      } catch (err) {
        console.error('Image upload error:', err);
        URL.revokeObjectURL(objectUrl);
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
        const assignments = autoMatchSignatures(state.sheetData, state.signatures);
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
    
    try {
      let assignmentsToUse = state.assignments;
      if (isRetry && state.sheetData) {
        assignmentsToUse = autoMatchSignatures(state.sheetData, state.signatures);
        setState(prev => ({ ...prev, assignments: assignmentsToUse }));
      }

      console.log(`[내보내기 시작] 서명 개수: ${assignmentsToUse.size}, 파일 크기: ${state.excelBuffer.byteLength} bytes`);
      
      const blob = await generateFinalExcel(state.excelBuffer, assignmentsToUse, state.signatures);
      
      const elapsed = performance.now() - startTime;
      console.log(`[내보내기 완료] 소요 시간: ${elapsed.toFixed(1)}ms, 파일 크기: ${blob.size} bytes`);
      
      if (!blob || blob.size === 0) {
        throw new Error("생성된 파일이 비어있습니다 (0 bytes)");
      }

      if (blob.size < 50) {
        console.error(`❌ 파일 크기 이상: ${blob.size} bytes - 파일이 손상됨`);
        throw new Error(`생성된 파일이 너무 작습니다 (${blob.size} bytes). 제너레이터 로그를 확인해주세요.`);
      }

      const url = URL.createObjectURL(blob);
      console.log(`[다운로드 준비] Object URL 생성됨: ${url.substring(0, 50)}...`);
      
      const a = document.createElement('a');
      a.href = url;
      
      // Timestamp to avoid filename collision
      const timestamp = new Date().toISOString().slice(11,19).replace(/:/g,'');
      const filename = `서명완료_${timestamp}_${state.excelFile?.name || 'output.xlsx'}`;
      
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      
      console.log(`[다운로드 시작] 파일명: ${filename}`);
      
      // Clean up after a delay to allow download to start
      setTimeout(() => {
        URL.revokeObjectURL(url);
        console.log(`[메모리 정리] Object URL 해제됨`);
      }, 100);
      
      setState(prev => ({ ...prev, step: 'export' }));
      setToast({ msg: `✅ 파일이 생성되었습니다: ${filename} (${(blob.size / 1024).toFixed(1)}KB)`, type: 'success' });
      setError(null);
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : "알 수 없는 오류";
      const fullError = `${errorMsg}\n\n디버그 정보:\n- 엑셀 버퍼: ${state.excelBuffer?.byteLength || 0} bytes\n- 배치된 서명: ${state.assignments.size}개`;
      
      console.error(`[내보내기 실패] ${fullError}`);
      console.error("Full error object:", err);
      
      setError(`엑셀 파일 생성 실패: ${errorMsg}\n\n해결 방법:\n1. 브라우저 콘솔 로그 확인\n2. 파일 크기를 줄여보기\n3. 이미지 해상도 낮추기`);
    } finally {
      setProcessing(false);
    }
  };

  const cleanupBlobUrls = (signatures: Map<string, SignatureFile[]>) => {
    signatures.forEach(list => {
      list.forEach(s => {
        try {
          URL.revokeObjectURL(s.previewUrl);
        } catch (e) {
          console.warn("Failed to revoke object URL:", e);
        }
      });
    });
  };

  const handleReset = () => {
    if (window.confirm("정말로 처음부터 다시 시작하시겠습니까?\n모든 데이터가 초기화됩니다.")) {
      // Cleanup existing blob URLs
      cleanupBlobUrls(state.signatures);
      
      setState(getInitialState());
      setError(null);
      setToast({ msg: '초기화되었습니다.', type: 'info' });
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
      <div className="flex flex-col h-[calc(100vh-100px)]">
        {/* Toolbar */}
        <div className="bg-white p-4 shadow-sm border-b flex justify-between items-center z-10 flex-wrap gap-3">
          <div className="flex items-center gap-4 flex-wrap">
            <h2 className="text-lg sm:text-xl font-bold text-gray-800">미리보기 및 편집</h2>
            <span className="bg-blue-100 text-blue-800 text-xs px-3 py-1 rounded-full font-medium">
              {state.assignments.size}개 배치 / {state.sheetData.rows.length}행
            </span>
          </div>
          <div className="flex gap-2 flex-wrap">
             <button onClick={runAutoMatch} className="px-3 py-2 text-xs sm:text-sm text-gray-600 hover:bg-gray-100 rounded-lg flex items-center gap-1 whitespace-nowrap">
              <RefreshCw size={16} /> 재설정
            </button>
            <button onClick={handleReset} className="px-3 py-2 text-xs sm:text-sm text-red-600 hover:bg-red-50 rounded-lg flex items-center gap-1 whitespace-nowrap">
              <RotateCcw size={16} /> 초기화
            </button>
            <button 
              onClick={() => handleExport(false)}
              disabled={processing}
              className="bg-green-600 text-white px-4 py-2 text-xs sm:text-sm rounded-lg font-semibold hover:bg-green-700 flex items-center gap-1 shadow-md disabled:opacity-50 whitespace-nowrap"
            >
              {processing ? <RefreshCw className="animate-spin" size={16} /> : <Download size={16} />}
              다운로드
            </button>
          </div>
        </div>

        {/* Table View with optimizations for mobile */}
        <div className="flex-1 overflow-auto bg-gray-100 p-2 sm:p-8 custom-scrollbar relative">
          <div className="bg-white shadow-xl rounded-sm overflow-hidden inline-block min-w-full">
            <table className="border-collapse w-full table-auto sm:table-fixed">
              <tbody>
                {state.sheetData.rows.map((row) => (
                  <tr key={row.index} className="h-12 sm:h-10 border-b border-gray-200 hover:bg-gray-50">
                    {/* Render cells optimized for mobile */}
                    {row.cells.slice(0, 12).map((cell) => {  // Limit columns for mobile performance
                      const assignKey = `${cell.row}:${cell.col}`;
                      const assignment = state.assignments.get(assignKey);
                      
                      let sigImgUrl = null;
                      if (assignment) {
                         const sigs = state.signatures.get(assignment.signatureBaseName);
                         const sig = sigs?.find(s => s.variant === assignment.signatureVariantId);
                         sigImgUrl = sig?.previewUrl;
                      }

                      return (
                        <td 
                          key={cell.address} 
                          className={`border-r border-gray-200 px-2 py-1 text-xs sm:text-sm relative min-w-[60px] sm:min-w-[80px] ${assignment ? 'bg-blue-50/30' : ''}`}
                          title={`값: ${cell.value}`}
                        >
                          <div className="relative w-full h-full min-h-[40px] sm:min-h-[30px] flex items-center">
                            <span className="z-0 text-gray-400 select-none truncate max-w-full text-xs">
                              {cell.value}
                            </span>
                            
                            {/* Overlay Signature with better sizing */}
                            {assignment && sigImgUrl && (
                              <div 
                                className="absolute inset-0 z-10 flex items-center justify-center cursor-pointer group"
                              >
                                <img 
                                  src={sigImgUrl} 
                                  alt="sig" 
                                  className="pointer-events-none drop-shadow-sm mix-blend-multiply transition-transform duration-300"
                                  style={{
                                    transform: `rotate(${assignment.rotation}deg) scale(${assignment.scale}) translate(${assignment.offsetX}px, ${assignment.offsetY}px)`,
                                    maxWidth: '130%', 
                                    maxHeight: '130%',
                                    objectFit: 'contain'
                                  }}
                                />
                                <div className="hidden group-hover:block absolute -top-8 left-0 bg-black text-white text-xs p-1 rounded whitespace-nowrap z-20 text-xs">
                                  {assignment.signatureVariantId.substring(0, 15)}
                                </div>
                              </div>
                            )}
                          </div>
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {state.sheetData.rows[0]?.cells.length > 12 && (
            <p className="text-xs text-gray-500 mt-2 text-center">📱 모바일에서는 처음 12개 열만 표시됩니다</p>
          )}
        </div>
      </div>
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
      </div>
      <button 
        onClick={handleReset}
        className="text-gray-400 hover:text-gray-600 underline text-sm mt-4"
      >
        처음으로 돌아가기 (파일 초기화)
      </button>
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