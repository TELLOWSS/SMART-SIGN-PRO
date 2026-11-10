import React from 'react';
import {
  Download,
  FileSpreadsheet,
  FileText,
  Image as ImageIcon,
  RefreshCw,
  PackageOpen,
  SlidersHorizontal,
  ShieldCheck,
  Scaling,
  Users,
  PenTool,
  Layers,
  AlertTriangle,
} from 'lucide-react';
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
  maskResidentNumber: boolean;
  signatureScalePreset: number;
  availableSheets: string[];
  activeSheetIndex: number;
  unmatchedWorkerCount: number;
  onVariationStrengthChange: (value: number) => void;
  onBatchCountChange: (value: number) => void;
  onExportFormatChange: (format: 'excel' | 'pdf' | 'png') => void;
  onMaskResidentNumberChange: (value: boolean) => void;
  onSignatureScalePresetChange: (value: number) => void;
  onSelectSheet: (index: number) => void;
  onOpenManualMatch: () => void;
  onOpenSignaturePad: () => void;
  onAutoMatch: () => void;
  onSingleExport: () => void;
  onBatchZipExport: () => void;
  onCancelBatchExport: () => void;
  isBatchCancelable: boolean;
  onStartOver: () => void;
  assignmentCount: number;
  rowCount: number;
}

/**
 * 프리미엄 SaaS 스타일 워크스페이스
 * - 좌측: 컨트롤 패널 (마스킹, 크기 프리셋, 미매칭 수동연결, 즉석 패드, 강도, 배치 수량, 액션)
 * - 우측: 다중 워크시트 탭 + 실시간 라이브 프리뷰(병합셀 + 서명 오버레이)
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
    maskResidentNumber,
    signatureScalePreset,
    availableSheets,
    activeSheetIndex,
    unmatchedWorkerCount,
    onVariationStrengthChange,
    onBatchCountChange,
    onExportFormatChange,
    onMaskResidentNumberChange,
    onSignatureScalePresetChange,
    onSelectSheet,
    onOpenManualMatch,
    onOpenSignaturePad,
    onAutoMatch,
    onSingleExport,
    onBatchZipExport,
    onCancelBatchExport,
    isBatchCancelable,
    onStartOver,
    assignmentCount,
    rowCount,
  } = props;

  return (
    <div className="grid grid-cols-1 xl:grid-cols-[340px_1fr] gap-4 h-[calc(100vh-96px)] p-4 bg-slate-100">
      <aside className="bg-white rounded-2xl shadow-sm border border-slate-200 p-5 flex flex-col gap-4 overflow-y-auto">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">Control Panel</h2>
          <p className="text-xs text-slate-500 mt-0.5">서식 보존 및 실무 커스텀 옵션을 설정합니다.</p>
        </div>

        {/* 1. 개인정보 보호 (주민번호 마스킹) */}
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2">
          <div className="flex items-center justify-between">
            <span className="flex items-center gap-1.5 text-xs font-semibold text-slate-800">
              <ShieldCheck size={15} className="text-emerald-600" /> 주민번호 뒷자리 마스킹
            </span>
            <label className="relative inline-flex items-center cursor-pointer">
              <input
                type="checkbox"
                checked={maskResidentNumber}
                onChange={(e) => onMaskResidentNumberChange(e.target.checked)}
                disabled={processing}
                className="sr-only peer"
              />
              <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-emerald-600"></div>
            </label>
          </div>
          <p className="text-[11px] text-slate-500 leading-tight">
            내국인(<span className="font-mono text-slate-700">670305-1******</span>) 및 외국인등록번호 뒷자리를 마스킹합니다.
          </p>
        </div>

        {/* 2. 서명란 맞춤 크기 프리셋 */}
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2.5">
          <div className="flex items-center justify-between text-xs text-slate-800">
            <span className="flex items-center gap-1.5 font-semibold">
              <Scaling size={14} className="text-indigo-600" /> 서명 크기 맞춤
            </span>
            <span className="text-indigo-700 font-bold">{Math.round(signatureScalePreset * 100)}%</span>
          </div>

          <div className="grid grid-cols-3 gap-1 text-[11px]">
            <button
              type="button"
              onClick={() => onSignatureScalePresetChange(0.85)}
              disabled={processing}
              className={`py-1.5 px-2 rounded-lg font-medium transition-colors ${
                Math.abs(signatureScalePreset - 0.85) < 0.02
                  ? 'bg-indigo-600 text-white font-semibold shadow-sm'
                  : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
              }`}
            >
              컴팩트 85%
            </button>
            <button
              type="button"
              onClick={() => onSignatureScalePresetChange(1.0)}
              disabled={processing}
              className={`py-1.5 px-2 rounded-lg font-medium transition-colors ${
                Math.abs(signatureScalePreset - 1.0) < 0.02
                  ? 'bg-indigo-600 text-white font-semibold shadow-sm'
                  : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
              }`}
            >
              표준 100%
            </button>
            <button
              type="button"
              onClick={() => onSignatureScalePresetChange(1.15)}
              disabled={processing}
              className={`py-1.5 px-2 rounded-lg font-medium transition-colors ${
                Math.abs(signatureScalePreset - 1.15) < 0.02
                  ? 'bg-indigo-600 text-white font-semibold shadow-sm'
                  : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
              }`}
            >
              확대 115%
            </button>
          </div>

          <input
            type="range"
            min={70}
            max={130}
            step={5}
            value={Math.round(signatureScalePreset * 100)}
            onChange={(e) => onSignatureScalePresetChange(Number(e.target.value) / 100)}
            disabled={processing}
            className="w-full accent-indigo-600"
          />
        </div>

        {/* 3. 외국인 근로자 및 미매칭 관리 & 현장 즉석 서명 */}
        <div className="grid grid-cols-2 gap-2">
          <button
            type="button"
            onClick={onOpenManualMatch}
            disabled={processing}
            className="p-2.5 rounded-xl border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-900 text-xs font-semibold flex flex-col items-center justify-center gap-1 transition-colors text-center relative"
          >
            <Users size={16} className="text-indigo-600" />
            <span>수동 서명 매칭</span>
            {unmatchedWorkerCount > 0 && (
              <span className="absolute -top-1.5 -right-1.5 px-1.5 py-0.5 rounded-full bg-amber-500 text-white text-[10px] font-bold shadow-sm flex items-center gap-0.5">
                <AlertTriangle size={9} /> {unmatchedWorkerCount}
              </span>
            )}
          </button>

          <button
            type="button"
            onClick={onOpenSignaturePad}
            disabled={processing}
            className="p-2.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-800 text-xs font-semibold flex flex-col items-center justify-center gap-1 transition-colors text-center shadow-sm"
          >
            <PenTool size={16} className="text-indigo-600" />
            <span>현장 즉석 서명</span>
          </button>
        </div>

        {/* 4. 서명 변형 강도 */}
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2">
          <div className="flex items-center justify-between text-xs text-slate-700">
            <span className="flex items-center gap-1.5 font-medium"><SlidersHorizontal size={13} />서명 변형 강도</span>
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
          <p className="text-[11px] text-slate-500">회전, 크기, 위치의 자연스러운 편차 폭</p>
        </div>

        {/* 5. 일괄 생성 개수 */}
        <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2">
          <label className="text-xs font-medium text-slate-700 block">일괄 생성 파일 개수 (N)</label>
          <input
            type="number"
            min={1}
            max={50}
            value={batchCount}
            onChange={(e) => onBatchCountChange(Number(e.target.value || 1))}
            disabled={processing}
            className="w-full rounded-lg border border-slate-300 px-3 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500"
          />
        </div>

        {/* 6. 출력 포맷 */}
        <div className="flex flex-wrap gap-1.5">
          <button
            onClick={() => onExportFormatChange('excel')}
            className={`flex-1 px-3 py-1.5 rounded-lg text-xs font-medium flex items-center justify-center gap-1 ${
              exportFormat === 'excel' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'
            }`}
          >
            <FileSpreadsheet size={13} /> Excel
          </button>
          <button
            onClick={() => onExportFormatChange('pdf')}
            className={`flex-1 px-3 py-1.5 rounded-lg text-xs font-medium flex items-center justify-center gap-1 ${
              exportFormat === 'pdf' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'
            }`}
          >
            <FileText size={13} /> PDF
          </button>
          <button
            onClick={() => onExportFormatChange('png')}
            className={`flex-1 px-3 py-1.5 rounded-lg text-xs font-medium flex items-center justify-center gap-1 ${
              exportFormat === 'png' ? 'bg-indigo-600 text-white' : 'bg-slate-100 text-slate-700'
            }`}
          >
            <ImageIcon size={13} /> PNG
          </button>
        </div>

        {/* 7. 주요 액션 버튼 */}
        <div className="grid grid-cols-2 gap-2">
          <button
            onClick={onAutoMatch}
            disabled={processing}
            className="px-3 py-2 rounded-xl bg-slate-900 text-white text-xs font-semibold hover:bg-slate-800 disabled:opacity-50 flex items-center justify-center gap-1.5 transition-colors"
          >
            <RefreshCw size={13} /> 재매칭
          </button>
          <button
            onClick={onSingleExport}
            disabled={processing}
            className="px-3 py-2 rounded-xl bg-emerald-600 text-white text-xs font-semibold hover:bg-emerald-700 disabled:opacity-50 flex items-center justify-center gap-1.5 shadow-sm transition-colors"
          >
            <Download size={13} /> 단일 저장
          </button>
        </div>

        <button
          onClick={onBatchZipExport}
          disabled={processing || batchCount < 1}
          className="w-full px-4 py-2.5 rounded-xl bg-indigo-600 text-white text-xs font-bold hover:bg-indigo-700 disabled:opacity-50 flex items-center justify-center gap-2 shadow-sm transition-colors"
        >
          <PackageOpen size={15} /> 일괄 생성 및 ZIP 다운로드
        </button>

        <button
          onClick={onStartOver}
          className="w-full px-4 py-2 rounded-xl border border-slate-300 text-slate-700 bg-white hover:bg-slate-50 text-xs font-medium flex items-center justify-center gap-1.5 transition-colors"
        >
          <RefreshCw size={13} /> 처음으로 돌아가기
        </button>

        {isBatchCancelable && (
          <button
            onClick={onCancelBatchExport}
            className="w-full px-4 py-2 rounded-xl bg-rose-50 text-rose-700 border border-rose-200 text-xs font-semibold hover:bg-rose-100 flex items-center justify-center gap-1.5"
          >
            작업 취소
          </button>
        )}

        {batchProgress && (
          <div className="bg-indigo-50 border border-indigo-100 rounded-xl p-3">
            <div className="flex items-center justify-between text-xs text-indigo-900 mb-1.5">
              <span>
                {batchProgress.phase === 'zip' ? 'ZIP 압축 중' : '파일 생성 중'}
                {batchProgress.mode === 'high-volume' ? ' · 고용량 모드' : ''}
              </span>
              <span className="font-semibold">{Math.round(batchProgress.percent)}%</span>
            </div>
            <div className="w-full h-2 rounded-full bg-indigo-100 overflow-hidden">
              <div className="h-full bg-indigo-600 transition-all" style={{ width: `${batchProgress.percent}%` }} />
            </div>
            <p className="text-[11px] text-indigo-800 mt-1.5">
              {batchProgress.message || (batchProgress.phase === 'zip'
                ? `압축 진행 중... (${batchProgress.current}/${batchProgress.total})`
                : `${batchProgress.current}/${batchProgress.total} 파일 생성 완료`)}
            </p>
          </div>
        )}

        <div className="mt-auto bg-slate-50 rounded-xl border border-slate-200 p-2.5 text-[11px] text-slate-600 space-y-0.5">
          <p>배치된 서명: <span className="font-semibold text-slate-800">{assignmentCount}개</span></p>
          <p>프리뷰 행 수: <span className="font-semibold text-slate-800">{rowCount}행</span></p>
        </div>
      </aside>

      <section className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
        {/* Header with multi-sheet tabs */}
        <div className="px-5 py-3 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
          <div>
            <h3 className="text-sm font-semibold text-slate-800">Live Preview</h3>
            <p className="text-xs text-slate-500">병합 셀 구조와 서명 오버레이 상태를 실시간으로 확인합니다.</p>
          </div>

          {/* 다중 워크시트(월별/공종별 탭) 선택 */}
          {availableSheets && availableSheets.length > 1 && (
            <div className="flex items-center gap-1.5 overflow-x-auto max-w-xl pb-0.5">
              <span className="text-xs font-semibold text-slate-500 flex items-center gap-1 mr-1">
                <Layers size={13} /> 시트:
              </span>
              {availableSheets.map((sheetName, idx) => (
                <button
                  key={`sheet-tab-${idx}-${sheetName}`}
                  type="button"
                  onClick={() => onSelectSheet(idx)}
                  disabled={processing}
                  className={`px-3 py-1 rounded-lg text-xs font-medium whitespace-nowrap transition-colors ${
                    activeSheetIndex === idx
                      ? 'bg-indigo-600 text-white font-semibold shadow-sm'
                      : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  {sheetName}
                </button>
              ))}
            </div>
          )}
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
