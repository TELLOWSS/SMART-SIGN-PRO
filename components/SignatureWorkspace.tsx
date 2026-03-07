import React from 'react';
import { Download, FileSpreadsheet, FileText, Image as ImageIcon, RefreshCw, PackageOpen, SlidersHorizontal } from 'lucide-react';
import { SheetPreviewModel } from '../services/alternativeExportService';

interface BatchProgress {
  current: number;
  total: number;
  phase: 'generate' | 'zip' | 'done';
  percent: number;
  mode?: 'standard' | 'high-volume';
  message?: string;
}

interface SignatureWorkspaceProps {
  previewLoading: boolean;
  previewModel: SheetPreviewModel | null;
  processing: boolean;
  exportFormat: 'excel' | 'pdf' | 'png';
  variationStrength: number;
  batchCount: number;
  batchProgress: BatchProgress | null;
  onVariationStrengthChange: (value: number) => void;
  onBatchCountChange: (value: number) => void;
  onExportFormatChange: (format: 'excel' | 'pdf' | 'png') => void;
  onAutoMatch: () => void;
  onSingleExport: () => void;
  onBatchZipExport: () => void;
  onCancelBatchExport: () => void;
  isBatchCancelable: boolean;
  assignmentCount: number;
  rowCount: number;
}

/**
 * 프리미엄 SaaS 스타일 워크스페이스
 * - 좌측: 컨트롤 패널(강도/배치 수량/액션)
 * - 우측: 실시간 라이브 프리뷰(병합셀 + 서명 오버레이)
 */
export default function SignatureWorkspace(props: SignatureWorkspaceProps) {
  const {
    previewLoading,
    previewModel,
    processing,
    exportFormat,
    variationStrength,
    batchCount,
    batchProgress,
    onVariationStrengthChange,
    onBatchCountChange,
    onExportFormatChange,
    onAutoMatch,
    onSingleExport,
    onBatchZipExport,
    onCancelBatchExport,
    isBatchCancelable,
    assignmentCount,
    rowCount,
  } = props;

  return (
    <div className="grid grid-cols-1 xl:grid-cols-[320px_1fr] gap-4 h-[calc(100vh-96px)] p-4 bg-slate-100">
      <aside className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5 flex flex-col gap-5">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">Control Panel</h2>
          <p className="text-sm text-slate-500 mt-1">서명 변형과 배치 생성을 제어합니다.</p>
        </div>

        <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
          <div className="flex items-center justify-between text-sm text-slate-700">
            <span className="flex items-center gap-2 font-medium"><SlidersHorizontal size={14} />서명 변형 강도</span>
            <span className="text-indigo-700 font-semibold">{variationStrength}</span>
          </div>
          <input
            type="range"
            min={0}
            max={100}
            value={variationStrength}
            onChange={(e) => onVariationStrengthChange(Number(e.target.value))}
            disabled={processing}
            className="w-full accent-indigo-600"
          />
          <p className="text-xs text-slate-500">값이 높을수록 회전/크기/위치 편차가 커집니다.</p>
        </div>

        <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3">
          <label className="text-sm font-medium text-slate-700 block">일괄 생성 파일 개수 (N)</label>
          <input
            type="number"
            min={1}
            max={50}
            value={batchCount}
            onChange={(e) => onBatchCountChange(Number(e.target.value || 1))}
            disabled={processing}
            className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
          />
          <p className="text-xs text-slate-500">권장: 1~20부 (대용량은 처리 시간 증가)</p>
        </div>

        <div className="flex flex-wrap gap-2">
          <button
            onClick={() => onExportFormatChange('excel')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1 ${
              exportFormat === 'excel' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'
            }`}
          >
            <FileSpreadsheet size={13} /> Excel
          </button>
          <button
            onClick={() => onExportFormatChange('pdf')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1 ${
              exportFormat === 'pdf' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'
            }`}
          >
            <FileText size={13} /> PDF
          </button>
          <button
            onClick={() => onExportFormatChange('png')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1 ${
              exportFormat === 'png' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'
            }`}
          >
            <ImageIcon size={13} /> PNG
          </button>
        </div>

        <div className="grid grid-cols-2 gap-2">
          <button
            onClick={onAutoMatch}
            disabled={processing}
            className="px-3 py-2 rounded-lg bg-slate-900 text-white text-sm font-medium hover:bg-slate-800 disabled:opacity-50 flex items-center justify-center gap-1"
          >
            <RefreshCw size={14} /> 재매칭
          </button>
          <button
            onClick={onSingleExport}
            disabled={processing}
            className="px-3 py-2 rounded-lg bg-emerald-600 text-white text-sm font-medium hover:bg-emerald-700 disabled:opacity-50 flex items-center justify-center gap-1"
          >
            <Download size={14} /> 단일 저장
          </button>
        </div>

        <button
          onClick={onBatchZipExport}
          disabled={processing || batchCount < 1}
          className="w-full px-4 py-3 rounded-xl bg-indigo-600 text-white font-semibold hover:bg-indigo-700 disabled:opacity-50 flex items-center justify-center gap-2 shadow-sm"
        >
          <PackageOpen size={16} /> 일괄 생성 및 ZIP 다운로드
        </button>

        {isBatchCancelable && (
          <button
            onClick={onCancelBatchExport}
            className="w-full px-4 py-2.5 rounded-xl bg-rose-50 text-rose-700 border border-rose-200 font-semibold hover:bg-rose-100 flex items-center justify-center gap-2"
          >
            작업 취소
          </button>
        )}

        {batchProgress && (
          <div className="bg-indigo-50 border border-indigo-100 rounded-xl p-3">
            <div className="flex items-center justify-between text-xs text-indigo-900 mb-2">
              <span>
                {batchProgress.phase === 'zip' ? 'ZIP 압축 중' : '파일 생성 중'}
                {batchProgress.mode === 'high-volume' ? ' · 고용량 모드' : ''}
              </span>
              <span className="font-semibold">{Math.round(batchProgress.percent)}%</span>
            </div>
            <div className="w-full h-2 rounded-full bg-indigo-100 overflow-hidden">
              <div className="h-full bg-indigo-600 transition-all" style={{ width: `${batchProgress.percent}%` }} />
            </div>
            <p className="text-[11px] text-indigo-800 mt-2">
              {batchProgress.message || (batchProgress.phase === 'zip'
                ? `압축 진행 중... (${batchProgress.current}/${batchProgress.total})`
                : `${batchProgress.current}/${batchProgress.total} 파일 생성 완료`)}
            </p>
          </div>
        )}

        <div className="mt-auto bg-slate-50 rounded-xl border border-slate-200 p-3 text-xs text-slate-600 space-y-1">
          <p>배치된 서명: <span className="font-semibold text-slate-800">{assignmentCount}개</span></p>
          <p>프리뷰 행 수: <span className="font-semibold text-slate-800">{rowCount}행</span></p>
        </div>
      </aside>

      <section className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
        <div className="px-5 py-3 border-b border-slate-200 bg-slate-50">
          <h3 className="text-sm font-semibold text-slate-800">Live Preview</h3>
          <p className="text-xs text-slate-500">병합 셀 구조와 서명 오버레이 상태를 실시간으로 확인합니다.</p>
        </div>

        <div className="flex-1 overflow-auto p-4 bg-slate-100">
          <div className="inline-block min-w-full bg-white rounded-lg border border-slate-200 shadow-sm overflow-hidden">
            {previewLoading && (
              <div className="p-10 text-center text-slate-500 text-sm">미리보기 렌더링 중입니다...</div>
            )}

            {!previewLoading && previewModel && (
              <table className="border-collapse w-full table-auto">
                <tbody>
                  {previewModel.rows.map((row) => (
                    <tr key={`preview-row-${row.row}`} className="border-b border-slate-200 hover:bg-slate-50/40">
                      {row.cells.map((cell) => {
                        if (cell.hidden) return null;

                        return (
                          <td
                            key={cell.key}
                            rowSpan={cell.rowSpan}
                            colSpan={cell.colSpan}
                            className="border-r border-slate-200 px-2 py-1 text-xs sm:text-sm relative min-w-[64px]"
                            style={{
                              fontFamily: cell.style.fontFamily,
                              fontSize: cell.style.fontSize,
                              fontWeight: cell.style.fontWeight,
                              fontStyle: cell.style.fontStyle,
                              textAlign: cell.style.textAlign,
                              verticalAlign: cell.style.verticalAlign,
                            }}
                          >
                            <div className="relative w-full h-full min-h-[34px] flex items-center justify-center">
                              <span className="z-0 text-slate-700 whitespace-pre-wrap break-words">{cell.text}</span>

                              {cell.signature && (
                                <div className="absolute inset-0 z-10 flex items-center justify-center pointer-events-none">
                                  <img
                                    src={cell.signature.src}
                                    alt="signature-preview"
                                    className="drop-shadow-sm mix-blend-multiply"
                                    style={{
                                      transform: cell.signature.transform,
                                      opacity: cell.signature.opacity,
                                      maxWidth: '130%',
                                      maxHeight: '130%',
                                      objectFit: 'contain',
                                    }}
                                  />
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
            )}

            {!previewLoading && !previewModel && (
              <div className="p-10 text-center text-slate-500 text-sm">미리보기를 생성할 수 없습니다.</div>
            )}
          </div>
        </div>
      </section>
    </div>
  );
}
