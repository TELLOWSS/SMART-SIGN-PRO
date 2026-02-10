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

    // Check file size (Warning if > 5MB)
    if (file.size > 5 * 1024 * 1024) {
      if (!window.confirm(`선택하신 엑셀 파일의 용량이 큽니다 (${(file.size / 1024 / 1024).toFixed(1)}MB).\n파일 내에 이미지가 많거나 행이 매우 많으면 'Out of Memory' 오류가 발생할 수 있습니다.\n\n계속 진행하시겠습니까?`)) {
        if (excelInputRef.current) excelInputRef.current.value = '';
        return;
      }
    }

    try {
      setProcessing(true);
      const buffer = await file.arrayBuffer();
      const sheetData = await parseExcelFile(buffer);
      setState(prev => ({ ...prev, excelFile: file, excelBuffer: buffer, sheetData }));
      setError(null);
    } catch (err) {
      setError("엑셀 파일을 읽는 중 오류가 발생했습니다. 유효한 .xlsx 파일인지 확인해주세요.");
      console.error(err);
    } finally {
      setProcessing(false);
      // Don't clear input immediately here to allow user to see the file name
    }
  };

  const handleSignatureUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setProcessing(true);
    const newSignatures = new Map<string, SignatureFile[]>(state.signatures);
    let count = 0;

    // Process files in chunks to avoid UI freeze, but here we just loop async
    // Using Blob URLs is fast enough to loop directly
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      // Only images
      if (!file.type.startsWith('image/')) continue;

      // --- MEMORY FIX: Use Blob URL instead of reading file content ---
      const objectUrl = URL.createObjectURL(file);

      // We need dimensions to calculate aspect ratio later, but we load it lightly
      const getImageDims = () => new Promise<{w: number, h: number}>((resolve) => {
         const img = new Image();
         img.onload = () => resolve({ w: img.width, h: img.height });
         img.onerror = () => resolve({ w: 100, h: 50 }); // Fallback
         img.src = objectUrl;
      });

      const { w, h } = await getImageDims();

      // Parse name logic:
      const fileNameNoExt = file.name.substring(0, file.name.lastIndexOf('.'));
      const lastUnderscoreIdx = fileNameNoExt.lastIndexOf('_');
      
      let baseNameString = fileNameNoExt;
      if (lastUnderscoreIdx > 0) {
        baseNameString = fileNameNoExt.substring(0, lastUnderscoreIdx);
      }

      const baseName = normalizeName(baseNameString);
      
      const sigFile: SignatureFile = {
        name: baseName,
        variant: file.name,
        previewUrl: objectUrl, // Store lightweight URL
        width: w,
        height: h
      };

      const list: SignatureFile[] = newSignatures.get(sigFile.name) || [];
      // Avoid duplicates
      if (!list.find(s => s.variant === sigFile.variant)) {
        list.push(sigFile);
        newSignatures.set(sigFile.name, list);
        count++;
      } else {
        // If duplicate, revoke the new URL to save memory
        URL.revokeObjectURL(objectUrl);
      }
    }

    setState(prev => ({ ...prev, signatures: newSignatures }));
    setProcessing(false);
    if (sigInputRef.current) sigInputRef.current.value = '';
    setToast({ msg: `${count}개의 서명이 추가되었습니다.`, type: 'success' });
  };

  const runAutoMatch = () => {
    if (!state.sheetData) return;
    setProcessing(true);
    // Timeout to allow UI to show processing state
    setTimeout(() => {
        const assignments = autoMatchSignatures(state.sheetData, state.signatures);
        setState(prev => ({ ...prev, assignments, step: 'preview' }));
        setProcessing(false);
        setToast({ msg: '서명이 무작위로 재배치되었습니다.', type: 'info' });
    }, 100);
  };

  const handleExport = async (isRetry: boolean = false) => {
    if (!state.excelBuffer) return;
    setProcessing(true);
    try {
      // If retrying/generating new variation, we might want to re-roll matching first if requested?
      // But user might just want to export the CURRENT preview.
      // If calling from Export screen "Generate Another", we should probably re-roll first.
      
      let assignmentsToUse = state.assignments;
      if (isRetry && state.sheetData) {
        assignmentsToUse = autoMatchSignatures(state.sheetData, state.signatures);
        setState(prev => ({ ...prev, assignments: assignmentsToUse }));
      }

      const blob = await generateFinalExcel(state.excelBuffer, assignmentsToUse, state.signatures);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      
      // Timestamp to avoid filename collision
      const timestamp = new Date().toISOString().slice(11,19).replace(/:/g,'');
      const filename = `서명완료_${timestamp}_${state.excelFile?.name || 'output.xlsx'}`;
      
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      
      setState(prev => ({ ...prev, step: 'export' }));
      setToast({ msg: `파일이 생성되었습니다: ${filename}`, type: 'success' });
    } catch (err) {
      console.error(err);
      setError("엑셀 파일 생성에 실패했습니다. 메모리가 부족하여 브라우저가 중단되었을 수 있습니다. 작업을 나누거나 이미지를 줄여주세요.");
    } finally {
      setProcessing(false);
    }
  };

  const handleReset = () => {
    if (window.confirm("정말로 처음부터 다시 시작하시겠습니까?\n모든 데이터가 초기화됩니다.")) {
      // Cleanup existing blob URLs
      state.signatures.forEach(list => {
        list.forEach(s => URL.revokeObjectURL(s.previewUrl));
      });
      
      setState(getInitialState());
      // Refs will be reset when component re-renders due to key prop
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
        <div className="bg-white p-4 shadow-sm border-b flex justify-between items-center z-10">
          <div className="flex items-center gap-4">
            <h2 className="text-xl font-bold text-gray-800">미리보기 및 편집</h2>
            <span className="bg-blue-100 text-blue-800 text-xs px-3 py-1 rounded-full font-medium">
              {state.assignments.size}개 서명 배치됨
            </span>
          </div>
          <div className="flex gap-3">
             <button onClick={runAutoMatch} className="px-4 py-2 text-sm text-gray-600 hover:bg-gray-100 rounded-lg flex items-center gap-2">
              <RefreshCw size={16} /> 무작위 재설정
            </button>
            <button onClick={handleReset} className="px-4 py-2 text-sm text-red-600 hover:bg-red-50 rounded-lg flex items-center gap-2">
              <RotateCcw size={16} /> 처음부터 다시
            </button>
            <button 
              onClick={() => handleExport(false)}
              disabled={processing}
              className="bg-green-600 text-white px-6 py-2 rounded-lg font-semibold hover:bg-green-700 flex items-center gap-2 shadow-md"
            >
              {processing ? <RefreshCw className="animate-spin" size={18} /> : <Download size={18} />}
              최종 엑셀 내보내기
            </button>
          </div>
        </div>

        {/* Table View */}
        <div className="flex-1 overflow-auto bg-gray-100 p-8 custom-scrollbar relative">
          <div className="bg-white shadow-xl rounded-sm overflow-hidden inline-block min-w-full">
            <table className="border-collapse w-full table-fixed">
              <tbody>
                {state.sheetData.rows.map((row) => (
                  <tr key={row.index} className="h-10 border-b border-gray-200 hover:bg-gray-50">
                    {/* Render cells. Assuming max 20 columns for performance or dynamic */}
                    {row.cells.map((cell) => {
                      const assignKey = `${cell.row}:${cell.col}`;
                      const assignment = state.assignments.get(assignKey);
                      
                      // Find signature image url if assigned
                      let sigImgUrl = null;
                      if (assignment) {
                         const sigs = state.signatures.get(assignment.signatureBaseName);
                         const sig = sigs?.find(s => s.variant === assignment.signatureVariantId);
                         sigImgUrl = sig?.previewUrl;
                      }

                      return (
                        <td 
                          key={cell.address} 
                          className={`border-r border-gray-200 px-2 py-1 text-sm relative min-w-[80px] ${assignment ? 'bg-blue-50/30' : ''}`}
                          title={`값: ${cell.value} | 행: ${cell.row}`}
                        >
                          <div className="relative w-full h-full min-h-[30px] flex items-center">
                            <span className="z-0 text-gray-400 select-none truncate max-w-full">
                              {cell.value}
                            </span>
                            
                            {/* Overlay Signature */}
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
                                    maxWidth: '120%', 
                                    maxHeight: '120%'
                                  }}
                                />
                                <div className="hidden group-hover:block absolute -top-8 left-0 bg-black text-white text-xs p-1 rounded whitespace-nowrap z-20">
                                  {assignment.signatureVariantId}
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
          <div className="max-w-4xl mx-auto mt-6 bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-lg flex items-center gap-3">
            <AlertCircle size={20} />
            {error}
            <button onClick={() => setError(null)} className="ml-auto text-sm underline">닫기</button>
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
            <h3 className="text-lg font-semibold">처리 중...</h3>
            <p className="text-sm text-gray-500">잠시만 기다려주세요</p>
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
                <p className="text-gray-600">서명 이미지 파일명은 <code>홍길동_1.png</code>, <code>홍길동_2.png</code> 와 같이 설정해주세요.</p>
                <div className="bg-gray-100 p-4 rounded-lg text-sm text-gray-700">
                  <strong>TIP:</strong><br/>
                  - 엑셀 파일은 Microsoft Excel 표준 형식을 권장합니다.<br/>
                  - '다른 랜덤 버전 즉시 다운로드' 버튼을 사용하면, 같은 파일에 대해 서명 배치를 새로 무작위로 섞어서 즉시 다운로드할 수 있습니다.
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