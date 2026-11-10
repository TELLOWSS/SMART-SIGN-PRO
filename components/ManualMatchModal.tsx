import React, { useState, useMemo } from 'react';
import { CheckCircle2, AlertCircle, X, Search, Sparkles, PenTool, Check } from 'lucide-react';
import { WorkerMatchItem, SignatureFile } from '../types';
import { normalizeName } from '../services/excelService';
import { maskIdentificationNumber } from '../services/excelUtils';

interface ManualMatchModalProps {
  isOpen: boolean;
  onClose: () => void;
  workerItems: WorkerMatchItem[];
  signatures: Map<string, SignatureFile[]>;
  maskResidentNumber: boolean;
  onApplyMatches: (manualAssignments: { rowNumber: number; signatureBaseName: string; signatureVariantId: string }[]) => void;
  onOpenSignaturePadForWorker: (workerName: string) => void;
}

/**
 * 외국인 근로자 및 미매칭 인원 수동 연결 모달
 * - 자동 매칭 실패 인원에 대한 원클릭 수동 연결 지원
 * - 이름 유사도 기반 스마트 추천
 * - 현장 즉석 서명 패드로의 즉시 연계 지원
 */
export default function ManualMatchModal({
  isOpen,
  onClose,
  workerItems,
  signatures,
  maskResidentNumber,
  onApplyMatches,
  onOpenSignaturePadForWorker,
}: ManualMatchModalProps) {
  const [filter, setFilter] = useState<'all' | 'unmatched' | 'matched'>('all');
  const [searchQuery, setSearchQuery] = useState('');
  
  // Local edit state: map rowNumber -> { baseName, variantId }
  const [selections, setSelections] = useState<Map<number, { baseName: string; variantId: string }>>(new Map());

  // Initialize/reset local state when modal opens
  React.useEffect(() => {
    if (isOpen) {
      const initialMap = new Map<number, { baseName: string; variantId: string }>();
      for (const item of workerItems) {
        if (item.isMatched && item.assignedSignatureName && item.assignedVariantId) {
          initialMap.set(item.rowNumber, {
            baseName: item.assignedSignatureName,
            variantId: item.assignedVariantId,
          });
        }
      }
      setSelections(initialMap);
      setFilter('all');
      setSearchQuery('');
    }
  }, [isOpen, workerItems]);

  // Flattened available signatures list
  const availableSignatureList = useMemo(() => {
    const list: { baseName: string; variant: string; previewUrl: string }[] = [];
    for (const [baseName, files] of signatures.entries()) {
      for (const f of files) {
        list.push({
          baseName,
          variant: f.variant,
          previewUrl: f.previewUrl,
        });
      }
    }
    return list;
  }, [signatures]);

  // Filtered workers
  const filteredWorkers = useMemo(() => {
    return workerItems.filter((item) => {
      const isCurrentlyMatched = selections.has(item.rowNumber);
      if (filter === 'unmatched' && isCurrentlyMatched) return false;
      if (filter === 'matched' && !isCurrentlyMatched) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().replace(/\s+/g, '');
        const matchName = item.workerName.toLowerCase().replace(/\s+/g, '').includes(q);
        const matchAffil = item.affiliation?.toLowerCase().replace(/\s+/g, '').includes(q);
        const matchJob = item.jobTitle?.toLowerCase().replace(/\s+/g, '').includes(q);
        return matchName || matchAffil || matchJob;
      }
      return true;
    });
  }, [workerItems, filter, searchQuery, selections]);

  // Smart fuzzy recommendation
  const handleAutoRecommend = () => {
    const nextMap = new Map(selections);
    let matchedCount = 0;

    for (const item of workerItems) {
      if (nextMap.has(item.rowNumber)) continue;

      const normName = normalizeName(item.workerName);
      if (!normName) continue;

      // Find best match in signatures
      let bestSig: { baseName: string; variant: string } | null = null;

      for (const sig of availableSignatureList) {
        const normSig = normalizeName(sig.baseName);
        if (normSig === normName || normName.includes(normSig) || normSig.includes(normName)) {
          bestSig = { baseName: sig.baseName, variant: sig.variant };
          break;
        }
      }

      if (bestSig) {
        nextMap.set(item.rowNumber, bestSig);
        matchedCount++;
      }
    }

    setSelections(nextMap);
  };

  const handleSelectSignature = (rowNumber: number, variantId: string) => {
    const nextMap = new Map(selections);
    if (!variantId) {
      nextMap.delete(rowNumber);
    } else {
      const sig = availableSignatureList.find(s => s.variant === variantId);
      if (sig) {
        nextMap.set(rowNumber, { baseName: sig.baseName, variantId: sig.variant });
      }
    }
    setSelections(nextMap);
  };

  const handleApply = () => {
    const assignmentsToApply: { rowNumber: number; signatureBaseName: string; signatureVariantId: string }[] = [];
    for (const [rowNumber, sel] of selections.entries()) {
      assignmentsToApply.push({
        rowNumber,
        signatureBaseName: sel.baseName,
        signatureVariantId: sel.variantId,
      });
    }
    onApplyMatches(assignmentsToApply);
    onClose();
  };

  if (!isOpen) return null;

  const totalWorkers = workerItems.length;
  const currentMatchedCount = selections.size;
  const currentUnmatchedCount = Math.max(0, totalWorkers - currentMatchedCount);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4 animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl shadow-2xl max-w-3xl w-full max-h-[90vh] overflow-hidden border border-slate-200 flex flex-col">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
          <div>
            <h3 className="text-base font-bold text-slate-900">외국인 근로자 및 미매칭 서명 수동 연결</h3>
            <p className="text-xs text-slate-500 mt-0.5">
              자동 매칭되지 않은 근로자에게 적합한 서명을 직접 지정하거나 즉석 서명을 등록합니다.
            </p>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-slate-600 p-1.5 rounded-lg hover:bg-slate-100 transition-colors"
          >
            <X size={18} />
          </button>
        </div>

        {/* Toolbar */}
        <div className="p-4 border-b border-slate-200 bg-white flex flex-wrap items-center justify-between gap-3">
          {/* Status summary tabs */}
          <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl text-xs font-medium">
            <button
              onClick={() => setFilter('all')}
              className={`px-3 py-1.5 rounded-lg transition-colors ${filter === 'all' ? 'bg-white text-indigo-700 shadow-sm font-semibold' : 'text-slate-600 hover:text-slate-900'}`}
            >
              전체 ({totalWorkers})
            </button>
            <button
              onClick={() => setFilter('unmatched')}
              className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${filter === 'unmatched' ? 'bg-amber-500 text-white shadow-sm font-semibold' : 'text-slate-600 hover:text-slate-900'}`}
            >
              미매칭 ({currentUnmatchedCount})
            </button>
            <button
              onClick={() => setFilter('matched')}
              className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${filter === 'matched' ? 'bg-emerald-600 text-white shadow-sm font-semibold' : 'text-slate-600 hover:text-slate-900'}`}
            >
              매칭 완료 ({currentMatchedCount})
            </button>
          </div>

          <div className="flex items-center gap-2">
            {/* Search */}
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                placeholder="이름/소속/직위 검색..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="pl-8 pr-3 py-1.5 rounded-xl border border-slate-200 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500 w-44"
              />
            </div>

            {/* Smart Auto Recommend */}
            <button
              type="button"
              onClick={handleAutoRecommend}
              className="px-3 py-1.5 rounded-xl bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold flex items-center gap-1.5 transition-colors"
            >
              <Sparkles size={13} /> 스마트 자동 매칭
            </button>
          </div>
        </div>

        {/* List Body */}
        <div className="flex-1 overflow-y-auto p-4 divide-y divide-slate-100">
          {filteredWorkers.length === 0 ? (
            <div className="p-12 text-center text-slate-400 text-sm">
              조건에 맞는 근로자가 없습니다.
            </div>
          ) : (
            filteredWorkers.map((item) => {
              const currentSel = selections.get(item.rowNumber);
              const isMatched = !!currentSel;
              const matchedSigObj = currentSel ? availableSignatureList.find(s => s.variant === currentSel.variantId) : null;
              const displayResident = maskResidentNumber
                ? maskIdentificationNumber(item.residentNumber)
                : item.residentNumber;

              return (
                <div key={`worker-row-${item.rowNumber}`} className="py-3 px-2 flex items-center justify-between gap-4 hover:bg-slate-50/70 rounded-xl transition-colors">
                  {/* Left: Worker Info */}
                  <div className="flex items-center gap-3 min-w-[200px]">
                    <div className="w-7 h-7 rounded-full bg-slate-100 flex items-center justify-center text-xs font-semibold text-slate-600">
                      {item.rowNumber}
                    </div>
                    <div>
                      <div className="flex items-center gap-2">
                        <span className="text-sm font-bold text-slate-900">{item.workerName}</span>
                        {item.jobTitle && (
                          <span className="text-[11px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-600 font-medium">
                            {item.jobTitle}
                          </span>
                        )}
                        {item.affiliation && (
                          <span className="text-[11px] px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-700 font-medium">
                            {item.affiliation}
                          </span>
                        )}
                      </div>
                      {displayResident && (
                        <p className="text-[11px] text-slate-400 mt-0.5">
                          ID: {displayResident}
                        </p>
                      )}
                    </div>
                  </div>

                  {/* Middle: Match Status */}
                  <div className="flex items-center gap-2">
                    {isMatched ? (
                      <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200">
                        <CheckCircle2 size={12} /> 매칭됨 ({currentSel.baseName})
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-semibold bg-amber-50 text-amber-700 border border-amber-200">
                        <AlertCircle size={12} /> 서명 미배정
                      </span>
                    )}

                    {matchedSigObj && (
                      <div className="w-10 h-7 rounded border border-slate-200 bg-white p-0.5 flex items-center justify-center overflow-hidden">
                        <img
                          src={matchedSigObj.previewUrl}
                          alt="preview"
                          className="max-h-full max-w-full object-contain mix-blend-multiply"
                        />
                      </div>
                    )}
                  </div>

                  {/* Right: Select & Action */}
                  <div className="flex items-center gap-2">
                    <select
                      value={currentSel?.variantId || ''}
                      onChange={(e) => handleSelectSignature(item.rowNumber, e.target.value)}
                      className="px-3 py-1.5 rounded-xl border border-slate-300 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white max-w-[180px]"
                    >
                      <option value="">(서명 선택 안함)</option>
                      {availableSignatureList.map((sig) => (
                        <option key={sig.variant} value={sig.variant}>
                          {sig.variant}
                        </option>
                      ))}
                    </select>

                    <button
                      type="button"
                      onClick={() => onOpenSignaturePadForWorker(item.workerName)}
                      className="px-2.5 py-1.5 rounded-xl border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold flex items-center gap-1 transition-colors"
                      title="이 근로자의 서명을 현장에서 즉석 패드로 그립니다"
                    >
                      <PenTool size={12} /> 직접 서명
                    </button>
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <p className="text-xs text-slate-500">
            총 <span className="font-semibold text-slate-800">{totalWorkers}명</span> 중 <span className="font-semibold text-emerald-600">{currentMatchedCount}명</span> 매칭됨
          </p>
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
              onClick={handleApply}
              className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
            >
              <Check size={14} /> 매칭 결과 저장 및 적용
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
