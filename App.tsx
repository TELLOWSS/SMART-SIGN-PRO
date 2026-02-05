import React, { useState, useEffect, useRef } from 'react';
import { Upload, FileSpreadsheet, Image as ImageIcon, CheckCircle, RotateCcw, Download, Settings, RefreshCw, AlertCircle, HelpCircle, X, ArrowRight, FileText, MousePointer2 } from 'lucide-react';
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
  
  // Refs to clear file inputs
  const excelInputRef = useRef<HTMLInputElement>(null);
  const sigInputRef = useRef<HTMLInputElement>(null);

  // --- Handlers ---

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

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
      // Reset input value to allow re-uploading the same file if needed
      if (excelInputRef.current) excelInputRef.current.value = '';
    }
  };

  const handleSignatureUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setProcessing(true);
    // Explicitly type the Map to prevent inference issues
    const newSignatures = new Map<string, SignatureFile[]>(state.signatures);
    let count = 0;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      // Only images
      if (!file.type.startsWith('image/')) continue;

      const reader = new FileReader();
      const loadPromise = new Promise<SignatureFile>((resolve) => {
        reader.onload = (evt) => {
          const img = new Image();
          img.onload = () => {
            // Parse name logic:
            // "HongGilDong_1.png" -> Base: "HongGilDong", Variant: "HongGilDong_1.png"
            // "HongGilDong.png" -> Base: "HongGilDong"
            // Handle multiple underscores: "Hong_Gil_Dong_1.png" -> "Hong_Gil_Dong"
            
            const fileNameNoExt = file.name.substring(0, file.name.lastIndexOf('.'));
            const lastUnderscoreIdx = fileNameNoExt.lastIndexOf('_');
            
            let baseNameString = fileNameNoExt;
            // If underscore exists, assume suffix (e.g. _1) is the variant ID
            if (lastUnderscoreIdx > 0) {
              baseNameString = fileNameNoExt.substring(0, lastUnderscoreIdx);
            }

            const baseName = normalizeName(baseNameString);
            
            resolve({
              name: baseName,
              variant: file.name, // unique ID including extension
              dataUrl: evt.target?.result as string,
              width: img.width,
              height: img.height
            });
          };
          img.src = evt.target?.result as string;
        };
        reader.readAsDataURL(file);
      });

      const sigFile = await loadPromise;
      const list: SignatureFile[] = newSignatures.get(sigFile.name) || [];
      // Avoid duplicates
      if (!list.find(s => s.variant === sigFile.variant)) {
        list.push(sigFile);
        newSignatures.set(sigFile.name, list);
        count++;
      }
    }

    setState(prev => ({ ...prev, signatures: newSignatures }));
    setProcessing(false);
    if (sigInputRef.current) sigInputRef.current.value = '';
  };

  const runAutoMatch = () => {
    if (!state.sheetData) return;
    const assignments = autoMatchSignatures(state.sheetData, state.signatures);
    setState(prev => ({ ...prev, assignments, step: 'preview' }));
  };

  const handleExport = async () => {
    if (!state.excelBuffer) return;
    setProcessing(true);
    try {
      const blob = await generateFinalExcel(state.excelBuffer, state.assignments, state.signatures);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `서명완료_${state.excelFile?.name || 'output.xlsx'}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setState(prev => ({ ...prev, step: 'export' }));
    } catch (err) {
      console.error(err);
      setError("엑셀 파일 생성에 실패했습니다.");
    } finally {
      setProcessing(false);
    }
  };

  const handleReset = () => {
    if (window.confirm("정말로 처음부터 다시 시작하시겠습니까?")) {
      setState(getInitialState());
      if (excelInputRef.current) excelInputRef.current.value = '';
      if (sigInputRef.current) sigInputRef.current.value = '';
    }
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
              onClick={handleExport}
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
                         sigImgUrl = sig?.dataUrl;
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
                                onClick={() => {
                                  // Manual toggle logic could go here
                                }}
                              >
                                <img 
                                  src={sigImgUrl} 
                                  alt="sig" 
                                  className="pointer-events-none drop-shadow-sm mix-blend-multiply"
                                  style={{
                                    transform: `rotate(${assignment.rotation}deg) scale(${assignment.scale}) translate(${assignment.offsetX}px, ${assignment.offsetY}px)`,
                                    maxWidth: '120%', // Allow slightly exceeding cell
                                    maxHeight: '120%'
                                  }}
                                />
                                {/* Hover tooltip */}
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
      <div className="bg-green-100 p-6 rounded-full text-green-600 mb-4">
        <CheckCircle size={64} />
      </div>
      <h2 className="text-3xl font-bold text-gray-800">내보내기 완료!</h2>
      <p className="text-gray-500 max-w-md text-center">
        무작위 서명이 포함된 엑셀 파일이 생성되어 다운로드되었습니다.<br/>
        파일 크기와 레이아웃은 100% 원본과 동일하게 유지됩니다.
      </p>
      <div className="flex gap-4 mt-8">
        <button 
          onClick={handleReset}
          className="bg-gray-800 text-white px-8 py-3 rounded-xl font-medium hover:bg-gray-700 flex items-center gap-2"
        >
          <RotateCcw size={18} />
          다른 파일 작업하기
        </button>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans">
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
            <p className="text-sm text-gray-500">데이터 분석 및 서명 매칭 중</p>
          </div>
        </div>
      )}

      {/* Guide Modal with Visual Workflow */}
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
              {/* Visual Workflow Banner */}
              <div className="bg-white p-8 border-b">
                <div className="flex flex-col md:flex-row justify-between items-center gap-4 text-center">
                  
                  {/* Step 1 */}
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-blue-100 rounded-2xl flex items-center justify-center text-blue-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <FileSpreadsheet size={32} />
                    </div>
                    <div className="font-bold text-gray-800">1. 엑셀 업로드</div>
                    <div className="text-xs text-gray-500 mt-1">이름 열 & '1' 마킹</div>
                  </div>

                  <ArrowRight className="text-gray-300 hidden md:block" />

                  {/* Step 2 */}
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-purple-100 rounded-2xl flex items-center justify-center text-purple-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <ImageIcon size={32} />
                    </div>
                    <div className="font-bold text-gray-800">2. 서명 업로드</div>
                    <div className="text-xs text-gray-500 mt-1">파일명: 홍길동_1.png</div>
                  </div>

                  <ArrowRight className="text-gray-300 hidden md:block" />

                  {/* Step 3 */}
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-indigo-100 rounded-2xl flex items-center justify-center text-indigo-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <Settings size={32} />
                    </div>
                    <div className="font-bold text-gray-800">3. 자동 매칭</div>
                    <div className="text-xs text-gray-500 mt-1">인쇄 최적화 안전 모드</div>
                  </div>

                  <ArrowRight className="text-gray-300 hidden md:block" />

                  {/* Step 4 */}
                  <div className="flex-1 flex flex-col items-center group">
                    <div className="w-16 h-16 bg-green-100 rounded-2xl flex items-center justify-center text-green-600 mb-3 shadow-sm group-hover:scale-110 transition-transform">
                      <Download size={32} />
                    </div>
                    <div className="font-bold text-gray-800">4. 엑셀 다운로드</div>
                    <div className="text-xs text-gray-500 mt-1">완벽한 매칭 & 내보내기</div>
                  </div>

                </div>
              </div>

              <div className="p-8 space-y-8">
                <section className="flex gap-4">
                  <div className="min-w-[40px] h-10 bg-gray-200 rounded-full flex items-center justify-center font-bold text-gray-600">1</div>
                  <div>
                    <h4 className="font-bold text-lg text-gray-900 mb-2">사전 준비 (Data Preparation)</h4>
                    <ul className="list-disc pl-5 space-y-2 text-sm text-gray-600">
                      <li><strong>엑셀 파일:</strong> '성명', '이름' 또는 'Name' 열이 필수입니다. 서명 위치에 숫자 <code>1</code>을 입력해두세요.</li>
                      <li><strong>서명 이미지:</strong> <code>이름_1.png</code>, <code>이름_2.png</code> 처럼 이름 뒤에 번호를 붙여 저장하세요. 프로그램이 자동으로 이름을 인식하고 랜덤으로 하나를 선택합니다.</li>
                    </ul>
                  </div>
                </section>

                <section className="flex gap-4">
                  <div className="min-w-[40px] h-10 bg-gray-200 rounded-full flex items-center justify-center font-bold text-gray-600">2</div>
                  <div>
                    <h4 className="font-bold text-lg text-gray-900 mb-2 text-orange-600">인쇄 안전 모드 (Safe Print Mode V2)</h4>
                    <p className="text-sm text-gray-600 mb-3">
                      인쇄 잘림 현상 분석 결과를 바탕으로 <strong>더욱 정밀한 안전 기준</strong>을 적용했습니다.
                    </p>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                      <div className="bg-white border rounded-lg p-3 text-center">
                        <div className="text-indigo-600 font-bold mb-1">회전</div>
                        <div className="text-xs text-gray-500">-8도 ~ +8도</div>
                      </div>
                      <div className="bg-white border rounded-lg p-3 text-center">
                        <div className="text-indigo-600 font-bold mb-1">크기</div>
                        <div className="text-xs text-gray-500">100% ~ 130%</div>
                      </div>
                      <div className="bg-white border border-orange-200 bg-orange-50 rounded-lg p-3 text-center">
                        <div className="text-orange-600 font-bold mb-1">위치</div>
                        <div className="text-xs text-gray-600 font-bold">상단 밀착 정렬 (하단 침범 방지)</div>
                      </div>
                    </div>
                  </div>
                </section>

                <section className="flex gap-4">
                  <div className="min-w-[40px] h-10 bg-gray-200 rounded-full flex items-center justify-center font-bold text-gray-600">3</div>
                  <div>
                    <h4 className="font-bold text-lg text-gray-900 mb-2">미리보기 및 내보내기</h4>
                    <p className="text-sm text-gray-600">
                      미리보기 화면에서 결과가 마음에 들지 않으면 <strong>'무작위 재설정'</strong> 버튼을 눌러보세요. 서명의 각도와 크기가 다시 랜덤하게 변경됩니다.
                    </p>
                  </div>
                </section>
              </div>
            </div>
            
            <div className="p-4 border-t border-gray-100 bg-white flex justify-end">
              <button 
                onClick={() => setShowGuide(false)}
                className="bg-indigo-600 text-white px-8 py-3 rounded-lg font-medium hover:bg-indigo-700 shadow-lg"
              >
                알겠습니다
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}