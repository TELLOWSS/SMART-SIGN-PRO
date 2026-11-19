import React, { useState, useMemo } from 'react';
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
  Crosshair,
  Sparkles,
  Check,
  Trash2,
  MousePointerClick,
  Search,
  ZoomIn,
  ZoomOut,
  Undo2,
  Redo2,
  Filter,
  X,
} from 'lucide-react';
import { SheetPreviewModel } from '../services/alternativeExportService';
import { SignaturePlacementMode } from '../types';

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
  placementMode: SignaturePlacementMode;
  detectedSignatureColLabel: string | null;
  enhanceSignatureInk: boolean;
  availableSheets: string[];
  activeSheetIndex: number;
  unmatchedWorkerCount: number;
  selectableColumns: { col: number; label: string; isSignatureCol: boolean; assignedCount: number }[];
  workerRowMap: Map<number, { workerName: string; nameCol: number }>;
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onToggleColumnSignatures: (col: number, label: string) => void;
  onToggleCellSignature: (row: number, col: number) => void;
  onClearAllAssignments: () => void;
  onVariationStrengthChange: (value: number) => void;
  onBatchCountChange: (value: number) => void;
  onExportFormatChange: (format: 'excel' | 'pdf' | 'png') => void;
  onMaskResidentNumberChange: (value: boolean) => void;
  onSignatureScalePresetChange: (value: number) => void;
  onPlacementModeChange: (mode: SignaturePlacementMode) => void;
  onEnhanceSignatureInkChange: (value: boolean) => void;
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
 * SMART SIGN PRO v2.0 플래그십 워크스페이스
 * - 좌측: 그룹화된 컨트롤 패널 + 하단 고정 Primary Export Dock
 * - 우측: 다중 시트 탭 + 검색/미서명 필터/확대축소/Undo·Redo 툴바 + 품목 일괄 칩 바 + 클릭-서명 라이브 프리뷰
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
    placementMode,
    detectedSignatureColLabel,
    enhanceSignatureInk,
    availableSheets,
    activeSheetIndex,
    unmatchedWorkerCount,
    selectableColumns,
    workerRowMap,
    canUndo,
    canRedo,
    onUndo,
    onRedo,
    onToggleColumnSignatures,
    onToggleCellSignature,
    onClearAllAssignments,
    onVariationStrengthChange,
    onBatchCountChange,
    onExportFormatChange,
    onMaskResidentNumberChange,
    onSignatureScalePresetChange,
    onPlacementModeChange,
    onEnhanceSignatureInkChange,
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

  const [previewSearch, setPreviewSearch] = useState('');
  const [onlyUnsignedRows, setOnlyUnsignedRows] = useState(false);
  const [zoomPercent, setZoomPercent] = useState(100);

  const totalWorkersCount = workerRowMap.size;

  // 라이브 프리뷰 행 필터링 (헤더 행은 항상 유지, 근로자 데이터 행만 검색/미서명 필터 적용)
  const visibleRows = useMemo(() => {
    if (!previewModel) return [];
    const q = previewSearch.trim().toLowerCase();
    if (!q && !onlyUnsignedRows) return previewModel.rows;

    return previewModel.rows.filter((row) => {
      const workerInfo = workerRowMap.get(row.row);
      if (!workerInfo) return true; // 제목/헤더 행은 항상 유지

      if (onlyUnsignedRows) {
        const hasAnySig = row.cells.some(c => !!c.signature);
        if (hasAnySig) return false;
      }

      if (q) {
        const rowText = row.cells.map(c => c.text || '').join(' ').toLowerCase();
        return rowText.includes(q);
      }

      return true;
    });
  }, [previewModel, previewSearch, onlyUnsignedRows, workerRowMap]);

  const adjustZoom = (delta: number) => {
    setZoomPercent(prev => Math.max(70, Math.min(130, prev + delta)));
  };

  return (
    <div className="grid grid-cols-1 xl:grid-cols-[352px_1fr] gap-3.5 h-[calc(100vh-64px)] p-3.5 bg-transparent animate-fade-up">
      {/* LEFT STUDIO CONTROL RAIL */}
      <aside className="ceramic-card rounded-3xl flex flex-col overflow-hidden">
        {/* Scrollable Settings Area */}
        <div className="p-4 flex-1 overflow-y-auto space-y-4">
          {/* Minimal Studio KPI Strip */}
          <div className="grid grid-cols-3 gap-2">
            <div className="bg-slate-900/[0.03] border border-slate-200/70 rounded-2xl px-2.5 py-2.5 text-center">
              <div className="text-[10px] font-semibold text-slate-400">근로자</div>
              <div className="text-[15px] font-extrabold text-slate-900 tracking-tight">{totalWorkersCount}<span className="text-[11px] font-semibold text-slate-400 ml-0.5">명</span></div>
            </div>
            <div className="bg-indigo-50/70 border border-indigo-200/60 rounded-2xl px-2.5 py-2.5 text-center">
              <div className="text-[10px] font-semibold text-indigo-600/80">배치 서명</div>
              <div className="text-[15px] font-extrabold text-indigo-700 tracking-tight">{assignmentCount}<span className="text-[11px] font-semibold text-indigo-400 ml-0.5">건</span></div>
            </div>
            <div className={`border rounded-2xl px-2.5 py-2.5 text-center ${unmatchedWorkerCount > 0 ? 'bg-amber-50/80 border-amber-200/80' : 'bg-emerald-50/80 border-emerald-200/80'}`}>
              <div className={`text-[10px] font-semibold ${unmatchedWorkerCount > 0 ? 'text-amber-700/80' : 'text-emerald-700/80'}`}>미서명</div>
              <div className={`text-[15px] font-extrabold tracking-tight ${unmatchedWorkerCount > 0 ? 'text-amber-700' : 'text-emerald-700'}`}>{unmatchedWorkerCount}<span className="text-[11px] font-semibold opacity-60 ml-0.5">명</span></div>
            </div>
          </div>

          {/* 1. 서명 배치 모드 (Apple/Linear Segmented Control) */}
          <div className="space-y-2 pt-1">
            <div className="flex items-center justify-between text-xs text-slate-900">
              <span className="flex items-center gap-1.5 font-bold">
                <Crosshair size={13} className="text-indigo-600" /> 서명 배치 모드
              </span>
              {detectedSignatureColLabel && (
                <span className="px-2 py-0.5 rounded-full bg-indigo-50 border border-indigo-200/70 text-indigo-700 text-[10px] font-bold">
                  [{detectedSignatureColLabel}] 열 감지
                </span>
              )}
            </div>

            <div className="grid grid-cols-3 gap-1 p-1 rounded-xl bg-slate-100/90 border border-slate-200/70 text-[11px]">
              <button
                type="button"
                onClick={() => onPlacementModeChange('auto')}
                disabled={processing}
                className={`py-1.5 px-1.5 rounded-lg font-semibold transition-all ${
                  placementMode === 'auto'
                    ? 'bg-white text-slate-950 shadow-sm ring-1 ring-slate-900/5'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                품목+서명
              </button>
              <button
                type="button"
                onClick={() => onPlacementModeChange('signatureColumnOnly')}
                disabled={processing}
                className={`py-1.5 px-1.5 rounded-lg font-semibold transition-all ${
                  placementMode === 'signatureColumnOnly'
                    ? 'bg-white text-slate-950 shadow-sm ring-1 ring-slate-900/5'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                서명 열만
              </button>
              <button
                type="button"
                onClick={() => onPlacementModeChange('placeholderCells')}
                disabled={processing}
                className={`py-1.5 px-1.5 rounded-lg font-semibold transition-all ${
                  placementMode === 'placeholderCells'
                    ? 'bg-white text-slate-950 shadow-sm ring-1 ring-slate-900/5'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                1 표시만
              </button>
            </div>
            <p className="text-[11px] text-slate-500 leading-snug">
              {placementMode === 'auto'
                ? detectedSignatureColLabel
                  ? `품목 칸 1번 + [${detectedSignatureColLabel}] 열 1번, 총 2번 자동 연동됩니다.`
                  : "품목/서명 칸에 자연스럽게 서명을 배치합니다."
                : placementMode === 'signatureColumnOnly'
                  ? "전용 '서명' 열에만 1번 서명을 배치합니다."
                  : "'1' 표시가 있는 셀에만 서명을 배치합니다."}
            </p>
          </div>

          {/* 2. 수동 매칭 & 현장 즉석 서명 */}
          <div className="grid grid-cols-2 gap-2">
            <button
              type="button"
              onClick={onOpenManualMatch}
              disabled={processing}
              className="p-2.5 rounded-xl border border-indigo-200 bg-white hover:bg-indigo-50 text-indigo-900 text-xs font-semibold flex flex-col items-center justify-center gap-1 transition-all shadow-sm relative"
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
              className="p-2.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-800 text-xs font-semibold flex flex-col items-center justify-center gap-1 transition-all shadow-sm"
            >
              <PenTool size={16} className="text-indigo-600" />
              <span>현장 즉석 서명</span>
            </button>
          </div>

          {/* 3. 품질 & 보안 토글 그룹 */}
          <div className="bg-slate-50 border border-slate-200/90 rounded-xl p-3 space-y-3">
            <div className="flex items-center justify-between gap-2">
              <div>
                <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-800">
                  <Sparkles size={14} className="text-indigo-600" /> 여백제거·진하게 보정
                </div>
                <p className="text-[10px] text-slate-500 mt-0.5">빈 여백 자동 크롭 및 볼펜 선명도 보강</p>
              </div>
              <label className="relative inline-flex items-center cursor-pointer flex-shrink-0">
                <input
                  type="checkbox"
                  checked={enhanceSignatureInk}
                  onChange={(e) => onEnhanceSignatureInkChange(e.target.checked)}
                  disabled={processing}
                  className="sr-only peer"
                />
                <div className="w-9 h-5 bg-slate-200 rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-indigo-600"></div>
              </label>
            </div>

            <div className="border-t border-slate-200/80 pt-2.5 flex items-center justify-between gap-2">
              <div>
                <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-800">
                  <ShieldCheck size={14} className="text-emerald-600" /> 주민번호 뒷자리 마스킹
                </div>
                <p className="text-[10px] text-slate-500 mt-0.5">내국인·외국인번호 개인정보 보호</p>
              </div>
              <label className="relative inline-flex items-center cursor-pointer flex-shrink-0">
                <input
                  type="checkbox"
                  checked={maskResidentNumber}
                  onChange={(e) => onMaskResidentNumberChange(e.target.checked)}
                  disabled={processing}
                  className="sr-only peer"
                />
                <div className="w-9 h-5 bg-slate-200 rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-emerald-600"></div>
              </label>
            </div>
          </div>

          {/* 4. 서명 크기 & 변형 강도 */}
          <div className="bg-slate-50 border border-slate-200/90 rounded-xl p-3 space-y-3">
            <div className="space-y-2">
              <div className="flex items-center justify-between text-xs text-slate-800">
                <span className="flex items-center gap-1.5 font-semibold">
                  <Scaling size={14} className="text-indigo-600" /> 서명 크기 맞춤
                </span>
                <span className="text-indigo-700 font-bold">{Math.round(signatureScalePreset * 100)}%</span>
              </div>
              <div className="grid grid-cols-3 gap-1 text-[11px]">
                {[
                  { label: '컴팩트 85%', val: 0.85 },
                  { label: '표준 100%', val: 1.0 },
                  { label: '확대 115%', val: 1.15 },
                ].map(p => (
                  <button
                    key={p.label}
                    type="button"
                    onClick={() => onSignatureScalePresetChange(p.val)}
                    disabled={processing}
                    className={`py-1 px-1.5 rounded-lg font-medium transition-colors ${
                      Math.abs(signatureScalePreset - p.val) < 0.02
                        ? 'bg-indigo-600 text-white font-semibold shadow-sm'
                        : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    {p.label}
                  </button>
                ))}
              </div>
              <input
                type="range"
                min={70}
                max={130}
                step={5}
                value={Math.round(signatureScalePreset * 100)}
                onChange={(e) => onSignatureScalePresetChange(Number(e.target.value) / 100)}
                disabled={processing}
                className="studio-slider"
              />
            </div>

            <div className="border-t border-slate-200/80 pt-2.5 space-y-1.5">
              <div className="flex items-center justify-between text-xs text-slate-700">
                <span className="flex items-center gap-1.5 font-medium"><SlidersHorizontal size={13} />필기 자연스러움 강도</span>
                <span className="text-indigo-700 font-semibold">{variationStrength}</span>
              </div>
              <input
                type="range"
                min={0}
                max={100}
                value={variationStrength}
                onChange={(e) => onVariationStrengthChange(Number(e.target.value))}
                disabled={processing}
                className="studio-slider"
              />
            </div>
          </div>

          {/* 5. 출력 포맷 & 일괄 수량 */}
          <div className="bg-slate-50 border border-slate-200/90 rounded-xl p-3 space-y-2.5">
            <div className="flex items-center justify-between gap-2">
              <span className="text-xs font-semibold text-slate-700">출력 형식</span>
              <div className="flex gap-1">
                {(['excel', 'pdf', 'png'] as const).map(fmt => (
                  <button
                    key={fmt}
                    type="button"
                    onClick={() => onExportFormatChange(fmt)}
                    className={`px-2.5 py-1 rounded-lg text-[11px] font-semibold flex items-center gap-1 transition-colors ${
                      exportFormat === fmt ? 'bg-indigo-600 text-white shadow-sm' : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
                    }`}
                  >
                    {fmt === 'excel' && <FileSpreadsheet size={12} />}
                    {fmt === 'pdf' && <FileText size={12} />}
                    {fmt === 'png' && <ImageIcon size={12} />}
                    {fmt.toUpperCase()}
                  </button>
                ))}
              </div>
            </div>

            <div className="flex items-center justify-between gap-2 pt-1 border-t border-slate-200/70">
              <label className="text-xs font-medium text-slate-700">ZIP 일괄 생성 (부)</label>
              <input
                type="number"
                min={1}
                max={50}
                value={batchCount}
                onChange={(e) => onBatchCountChange(Number(e.target.value || 1))}
                disabled={processing}
                className="w-20 rounded-lg border border-slate-300 bg-white px-2.5 py-1 text-xs font-semibold text-center focus:outline-none focus:ring-2 focus:ring-indigo-500"
              />
            </div>
          </div>

          {batchProgress && (
            <div className="bg-indigo-50 border border-indigo-200 rounded-xl p-3">
              <div className="flex items-center justify-between text-xs text-indigo-900 mb-1.5">
                <span className="font-semibold">
                  {batchProgress.phase === 'zip' ? 'ZIP 압축 중' : '파일 생성 중'}
                </span>
                <span className="font-bold">{Math.round(batchProgress.percent)}%</span>
              </div>
              <div className="w-full h-2 rounded-full bg-indigo-100 overflow-hidden">
                <div className="h-full bg-indigo-600 transition-all" style={{ width: `${batchProgress.percent}%` }} />
              </div>
              <p className="text-[11px] text-indigo-800 mt-1.5">
                {batchProgress.message || `${batchProgress.current}/${batchProgress.total} 완료`}
              </p>
            </div>
          )}
        </div>

        {/* Sticky Bottom Primary Export Dock */}
        <div className="p-3.5 border-t border-slate-200 bg-slate-50/90 space-y-2">
          <div className="grid grid-cols-2 gap-2">
            <button
              type="button"
              onClick={onAutoMatch}
              disabled={processing}
              className="px-3 py-2.5 rounded-xl bg-slate-900 text-white text-xs font-semibold hover:bg-slate-800 disabled:opacity-50 flex items-center justify-center gap-1.5 transition-colors"
            >
              <RefreshCw size={13} /> 랜덤 재매칭
            </button>
            <button
              type="button"
              onClick={onSingleExport}
              disabled={processing || assignmentCount === 0}
              className="px-3 py-2.5 rounded-xl bg-emerald-600 text-white text-xs font-bold hover:bg-emerald-700 disabled:opacity-50 flex items-center justify-center gap-1.5 shadow-sm transition-colors"
            >
              <Download size={14} /> 단일 저장
            </button>
          </div>

          <button
            type="button"
            onClick={onBatchZipExport}
            disabled={processing || batchCount < 1 || assignmentCount === 0}
            className="w-full px-4 py-2.5 rounded-xl bg-indigo-600 text-white text-xs font-bold hover:bg-indigo-700 disabled:opacity-50 flex items-center justify-center gap-2 shadow-sm transition-colors"
          >
            <PackageOpen size={15} /> {batchCount}부 일괄 생성 · ZIP 다운로드
          </button>

          {isBatchCancelable ? (
            <button
              type="button"
              onClick={onCancelBatchExport}
              className="w-full px-3 py-2 rounded-xl bg-rose-50 text-rose-700 border border-rose-200 text-xs font-semibold hover:bg-rose-100"
            >
              작업 취소
            </button>
          ) : (
            <button
              type="button"
              onClick={onStartOver}
              className="w-full px-3 py-1.5 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-200/60 text-[11px] font-medium flex items-center justify-center gap-1 transition-colors"
            >
              <RefreshCw size={12} /> 다른 파일로 새로 시작 ({rowCount}행)
            </button>
          )}
        </div>
      </aside>

      {/* RIGHT LIVE PREVIEW STUDIO */}
      <section className="ceramic-card rounded-3xl overflow-hidden flex flex-col">
        {/* Top Studio Toolbar: Title + Search + Filter + Zoom + Undo/Redo + Multi-Sheet */}
        <div className="px-5 py-3 border-b border-slate-200/80 bg-white/80 flex flex-wrap items-center justify-between gap-2.5">
          <div className="flex items-center gap-2 flex-wrap">
            <h3 className="text-sm font-bold text-slate-900">Live Preview</h3>
            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 text-[11px] font-medium border border-indigo-200/70">
              <MousePointerClick size={11} /> 칸 클릭으로 자유 지정
            </span>

            {/* 다중 워크시트 탭 */}
            {availableSheets && availableSheets.length > 1 && (
              <div className="flex items-center gap-1 ml-1 pl-2 border-l border-slate-300 overflow-x-auto max-w-md">
                <Layers size={13} className="text-slate-400 flex-shrink-0" />
                {availableSheets.map((sheetName, idx) => (
                  <button
                    key={`sheet-tab-${idx}-${sheetName}`}
                    type="button"
                    onClick={() => onSelectSheet(idx)}
                    disabled={processing}
                    className={`px-2.5 py-1 rounded-lg text-[11px] font-medium whitespace-nowrap transition-colors ${
                      activeSheetIndex === idx
                        ? 'bg-indigo-600 text-white font-semibold shadow-sm'
                        : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
                    }`}
                  >
                    {sheetName}
                  </button>
                ))}
              </div>
            )}
          </div>

          {/* Search, Unsigned Filter, Zoom, Undo/Redo */}
          <div className="flex items-center gap-1.5 flex-wrap">
            {/* 근로자 빠른 검색 */}
            <div className="relative">
              <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={previewSearch}
                onChange={(e) => setPreviewSearch(e.target.value)}
                placeholder="이름·직종 검색..."
                className="pl-7 pr-6 py-1 rounded-lg border border-slate-200 bg-white text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500 w-36"
              />
              {previewSearch && (
                <button
                  type="button"
                  onClick={() => setPreviewSearch('')}
                  className="absolute right-1.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <X size={12} />
                </button>
              )}
            </div>

            {/* 미서명 행만 필터 */}
            <button
              type="button"
              onClick={() => setOnlyUnsignedRows(v => !v)}
              className={`px-2.5 py-1 rounded-lg text-[11px] font-semibold flex items-center gap-1 border transition-colors ${
                onlyUnsignedRows
                  ? 'bg-amber-500 text-white border-amber-500 shadow-sm'
                  : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-100'
              }`}
              title="서명이 아직 없는 근로자 행만 모아보기"
            >
              <Filter size={11} /> 미서명만
            </button>

            {/* 확대 / 축소 */}
            <div className="inline-flex items-center bg-white border border-slate-200 rounded-lg overflow-hidden">
              <button
                type="button"
                onClick={() => adjustZoom(-10)}
                className="px-1.5 py-1 text-slate-600 hover:bg-slate-100"
                title="축소"
              >
                <ZoomOut size={13} />
              </button>
              <button
                type="button"
                onClick={() => setZoomPercent(100)}
                className="px-1.5 py-1 text-[11px] font-semibold text-slate-700 hover:bg-slate-100 min-w-[40px] text-center"
                title="100% 초기화"
              >
                {zoomPercent}%
              </button>
              <button
                type="button"
                onClick={() => adjustZoom(10)}
                className="px-1.5 py-1 text-slate-600 hover:bg-slate-100"
                title="확대"
              >
                <ZoomIn size={13} />
              </button>
            </div>

            {/* Undo / Redo */}
            <div className="inline-flex items-center bg-white border border-slate-200 rounded-lg overflow-hidden">
              <button
                type="button"
                onClick={onUndo}
                disabled={!canUndo || processing}
                className="px-2 py-1 text-slate-600 hover:bg-slate-100 disabled:opacity-35"
                title="실행 취소 (Ctrl+Z)"
              >
                <Undo2 size={13} />
              </button>
              <button
                type="button"
                onClick={onRedo}
                disabled={!canRedo || processing}
                className="px-2 py-1 text-slate-600 hover:bg-slate-100 disabled:opacity-35 border-l border-slate-200"
                title="다시 실행 (Ctrl+Y)"
              >
                <Redo2 size={13} />
              </button>
            </div>
          </div>
        </div>

        {/* 원하는 품목/서명 열 일괄 지정 칩 바 */}
        {selectableColumns && selectableColumns.length > 0 && (
          <div className="px-4 py-2 border-b border-slate-200 bg-white flex flex-wrap items-center gap-1.5">
            <span className="text-[11px] font-bold text-slate-600 mr-1 whitespace-nowrap">
              품목 일괄 지정:
            </span>
            {selectableColumns.map((colInfo) => {
              const isActive = colInfo.assignedCount > 0;
              return (
                <button
                  key={`col-chip-${colInfo.col}`}
                  type="button"
                  onClick={() => onToggleColumnSignatures(colInfo.col, colInfo.label)}
                  disabled={processing}
                  className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-medium transition-all border ${
                    isActive
                      ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm hover:bg-indigo-700'
                      : colInfo.isSignatureCol
                        ? 'bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100'
                        : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-indigo-50 hover:border-indigo-200 hover:text-indigo-700'
                  }`}
                  title={`클릭하여 [${colInfo.label}] 열 전체 근로자 서명 일괄 지정/해제`}
                >
                  {isActive && <Check size={12} />}
                  <span>{colInfo.label}</span>
                  {isActive && (
                    <span className="ml-0.5 px-1.5 py-0.2 rounded-full bg-white/20 text-[10px] font-bold">
                      {colInfo.assignedCount}
                    </span>
                  )}
                </button>
              );
            })}

            {assignmentCount > 0 && (
              <button
                type="button"
                onClick={onClearAllAssignments}
                disabled={processing}
                className="ml-auto inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-medium text-rose-600 bg-rose-50 border border-rose-200 hover:bg-rose-100 transition-colors"
                title="현재 배치된 모든 서명을 초기화하고 원하는 칸만 직접 선택합니다"
              >
                <Trash2 size={11} /> 전체 비우기
              </button>
            )}
          </div>
        )}

        {/* Zoomable Live Preview Table Canvas */}
        <div className="flex-1 overflow-auto p-5 studio-dot-grid">
          <div
            className="inline-block min-w-full bg-white rounded-2xl border border-slate-200/90 shadow-lg shadow-slate-900/5 overflow-hidden origin-top-left transition-transform"
            style={zoomPercent !== 100 ? { zoom: `${zoomPercent}%` } : undefined}
          >
            {previewLoading && (
              <div className="p-12 text-center text-slate-500 text-sm flex flex-col items-center gap-2">
                <RefreshCw size={22} className="animate-spin text-indigo-500" />
                <span>미리보기 렌더링 중입니다...</span>
              </div>
            )}

            {!previewLoading && previewModel && (
              <table className="border-collapse w-full table-auto">
                <tbody>
                  {visibleRows.map((row) => {
                    const workerInfo = workerRowMap.get(row.row);
                    return (
                      <tr key={`preview-row-${row.row}`} className="border-b border-slate-200 hover:bg-slate-50/50">
                        {row.cells.map((cell) => {
                          if (cell.hidden) return null;

                          const isInteractiveCell = !!workerInfo && cell.col !== workerInfo.nameCol;
                          const colLabel = selectableColumns.find(c => c.col === cell.col)?.label;

                          return (
                            <td
                              key={cell.key}
                              rowSpan={cell.rowSpan}
                              colSpan={cell.colSpan}
                              onClick={isInteractiveCell && !processing ? () => onToggleCellSignature(cell.row, cell.col) : undefined}
                              title={
                                isInteractiveCell
                                  ? `${workerInfo.workerName}${colLabel ? ` · [${colLabel}]` : ''} — 클릭하여 서명 ${cell.signature ? '해제' : '추가'}`
                                  : undefined
                              }
                              className={`border-r border-slate-200 px-2 py-1 text-xs sm:text-sm relative min-w-[64px] select-none ${
                                isInteractiveCell
                                  ? cell.signature
                                    ? 'cursor-pointer bg-indigo-50/35 hover:bg-rose-50/60 hover:ring-1 hover:ring-inset hover:ring-rose-300 transition-colors'
                                    : 'cursor-pointer hover:bg-indigo-50/70 hover:ring-1 hover:ring-inset hover:ring-indigo-300 transition-colors'
                                  : ''
                              }`}
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
                                        opacity: enhanceSignatureInk ? 0.95 : cell.signature.opacity,
                                        filter: enhanceSignatureInk ? 'contrast(1.35) brightness(0.82)' : undefined,
                                        maxWidth: '94%',
                                        maxHeight: '94%',
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
                    );
                  })}
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
