import React, { useRef, useState, useEffect } from 'react';
import { PenTool, RotateCcw, Check, X, Palette } from 'lucide-react';

interface SignaturePadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: (file: File, workerName: string) => void;
  defaultWorkerName?: string;
}

/**
 * 현장 즉석 수기 서명 패드 (Canvas)
 * - 터치 스크린(스마트폰, 태블릿) 및 PC 마우스 완벽 지원
 * - 부드러운 필압/곡선 렌더링
 * - 펜 색상(블랙, 블루) 및 굵기 선택 지원
 * - 투명 배경 PNG 생성 후 즉시 서명 라이브러리에 등록
 */
export default function SignaturePadModal({
  isOpen,
  onClose,
  onSave,
  defaultWorkerName = '',
}: SignaturePadModalProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [hasDrawn, setHasDrawn] = useState(false);
  const [workerName, setWorkerName] = useState(defaultWorkerName);
  const [penColor, setPenColor] = useState<'#0f172a' | '#1e3a8a'>('#0f172a');
  const [strokeWidth, setStrokeWidth] = useState<number>(3);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  useEffect(() => {
    if (isOpen) {
      setWorkerName(defaultWorkerName);
      setHasDrawn(false);
      setErrorMsg(null);
      // Give DOM time to render before clearing canvas
      setTimeout(clearCanvas, 50);
    }
  }, [isOpen, defaultWorkerName]);

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
    setErrorMsg(null);
  };

  const getCoordinates = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();

    if ('touches' in e) {
      const touch = e.touches[0];
      return {
        x: touch.clientX - rect.left,
        y: touch.clientY - rect.top,
      };
    } else {
      return {
        x: e.clientX - rect.left,
        y: e.clientY - rect.top,
      };
    }
  };

  const startDrawing = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const { x, y } = getCoordinates(e);
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.strokeStyle = penColor;
    ctx.lineWidth = strokeWidth;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    setIsDrawing(true);
    setHasDrawn(true);
  };

  const draw = (e: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>) => {
    if (!isDrawing) return;
    e.preventDefault();
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const { x, y } = getCoordinates(e);
    ctx.lineTo(x, y);
    ctx.stroke();
  };

  const stopDrawing = () => {
    if (!isDrawing) return;
    setIsDrawing(false);
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (ctx) ctx.closePath();
  };

  const handleSave = () => {
    const trimmedName = workerName.trim();
    if (!trimmedName) {
      setErrorMsg('근로자 성명을 입력해주세요.');
      return;
    }

    if (!hasDrawn) {
      setErrorMsg('서명란에 직접 서명해 주세요.');
      return;
    }

    const canvas = canvasRef.current;
    if (!canvas) return;

    canvas.toBlob((blob) => {
      if (!blob) {
        setErrorMsg('서명 이미지 생성에 실패했습니다.');
        return;
      }

      const file = new File(
        [blob],
        `${trimmedName}_직접서명_${Date.now()}.png`,
        { type: 'image/png' }
      );

      onSave(file, trimmedName);
      onClose();
    }, 'image/png');
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4 animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden border border-slate-200 flex flex-col">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-indigo-50 text-indigo-600 rounded-lg">
              <PenTool size={20} />
            </div>
            <div>
              <h3 className="text-base font-bold text-slate-900">현장 즉석 수기 서명 패드</h3>
              <p className="text-xs text-slate-500">터치나 마우스로 직접 서명하여 즉시 등록합니다.</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-slate-600 p-1.5 rounded-lg hover:bg-slate-100 transition-colors"
          >
            <X size={18} />
          </button>
        </div>

        {/* Body */}
        <div className="p-6 space-y-4">
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">
              근로자 성명 <span className="text-rose-500">*</span>
            </label>
            <input
              type="text"
              placeholder="예: 보반탄, 홍길동"
              value={workerName}
              onChange={(e) => {
                setWorkerName(e.target.value);
                if (errorMsg) setErrorMsg(null);
              }}
              className="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              autoFocus
            />
          </div>

          <div>
            <div className="flex items-center justify-between mb-1.5">
              <label className="text-xs font-semibold text-slate-700">서명 캔버스 (마우스/터치)</label>
              <div className="flex items-center gap-2">
                {/* 펜 색상 */}
                <div className="flex items-center gap-1 bg-slate-100 p-0.5 rounded-lg">
                  <button
                    type="button"
                    onClick={() => setPenColor('#0f172a')}
                    className={`w-5 h-5 rounded-md bg-slate-900 transition-transform ${penColor === '#0f172a' ? 'scale-110 ring-2 ring-indigo-500' : 'opacity-60'}`}
                    title="검정 펜"
                  />
                  <button
                    type="button"
                    onClick={() => setPenColor('#1e3a8a')}
                    className={`w-5 h-5 rounded-md bg-blue-900 transition-transform ${penColor === '#1e3a8a' ? 'scale-110 ring-2 ring-indigo-500' : 'opacity-60'}`}
                    title="파랑 펜"
                  />
                </div>

                {/* 펜 굵기 */}
                <div className="flex items-center gap-1 text-xs">
                  <button
                    type="button"
                    onClick={() => setStrokeWidth(2)}
                    className={`px-1.5 py-0.5 rounded border text-[11px] ${strokeWidth === 2 ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-slate-50 text-slate-600 border-slate-200'}`}
                  >
                    얇게
                  </button>
                  <button
                    type="button"
                    onClick={() => setStrokeWidth(3)}
                    className={`px-1.5 py-0.5 rounded border text-[11px] ${strokeWidth === 3 ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-slate-50 text-slate-600 border-slate-200'}`}
                  >
                    보통
                  </button>
                  <button
                    type="button"
                    onClick={() => setStrokeWidth(5)}
                    className={`px-1.5 py-0.5 rounded border text-[11px] ${strokeWidth === 5 ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-slate-50 text-slate-600 border-slate-200'}`}
                  >
                    굵게
                  </button>
                </div>
              </div>
            </div>

            <div className="relative border-2 border-dashed border-slate-300 rounded-xl bg-slate-50/50 overflow-hidden cursor-crosshair touch-none">
              <canvas
                ref={canvasRef}
                width={460}
                height={200}
                className="w-full h-[200px] block"
                onMouseDown={startDrawing}
                onMouseMove={draw}
                onMouseUp={stopDrawing}
                onMouseLeave={stopDrawing}
                onTouchStart={startDrawing}
                onTouchMove={draw}
                onTouchEnd={stopDrawing}
              />
              {!hasDrawn && (
                <div className="absolute inset-0 flex items-center justify-center pointer-events-none text-slate-400 text-sm">
                  이곳에 서명을 그려주세요 (터치 또는 마우스)
                </div>
              )}
            </div>
          </div>

          {errorMsg && (
            <p className="text-xs text-rose-600 font-medium bg-rose-50 p-2.5 rounded-lg border border-rose-200">
              ⚠️ {errorMsg}
            </p>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <button
            type="button"
            onClick={clearCanvas}
            className="px-3 py-2 rounded-xl border border-slate-300 text-slate-700 bg-white hover:bg-slate-50 text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <RotateCcw size={14} /> 다시 쓰기
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl text-slate-600 hover:text-slate-800 text-xs font-semibold hover:bg-slate-100 transition-colors"
            >
              취소
            </button>
            <button
              type="button"
              onClick={handleSave}
              className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
            >
              <Check size={14} /> 서명 등록 및 적용
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
