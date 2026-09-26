'use client';

import React from 'react';
import { ViewMode } from '@/types/schedule';
import { 
  Calendar as CalendarIcon, 
  Upload, 
  ChevronLeft, 
  ChevronRight, 
  Filter, 
  Sparkles,
  Activity
} from 'lucide-react';

interface MobileHeaderProps {
  currentDate: Date;
  onPrevDate: () => void;
  onNextDate: () => void;
  onToday: () => void;
  viewMode: ViewMode;
  onViewModeChange: (mode: ViewMode) => void;
  onOpenUpload: () => void;
  onlyMyShifts: boolean;
  onToggleOnlyMyShifts: () => void;
  targetUserName: string;
  hasFileLoaded: boolean;
  fileName?: string;
}

const VIEW_MODES: { key: ViewMode; label: string; shortLabel: string }[] = [
  { key: 'week', label: '주간 (Week)', shortLabel: '주' },
  { key: 'month', label: '1개월 (Month)', shortLabel: '1월' },
  { key: '3months', label: '3개월 (3M)', shortLabel: '3월' },
  { key: '6months', label: '6개월 (6M)', shortLabel: '6월' },
  { key: '1year', label: '1년 (Year)', shortLabel: '1년' },
];

export const MobileHeader: React.FC<MobileHeaderProps> = ({
  currentDate,
  onPrevDate,
  onNextDate,
  onToday,
  viewMode,
  onViewModeChange,
  onOpenUpload,
  onlyMyShifts,
  onToggleOnlyMyShifts,
  targetUserName,
  hasFileLoaded,
  fileName,
}) => {
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth() + 1;

  return (
    <header className="sticky top-0 z-30 pt-safe glass-header border-b border-slate-200/80 shadow-xs">
      {/* 1. 최상단 상태 정보 및 업로드 액션 바 */}
      <div className="px-4 pt-3 pb-2 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-600 flex items-center justify-center text-white shadow-md shadow-blue-500/20">
            <Activity className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-1.5">
              <span className="text-xs font-semibold text-blue-600 bg-blue-50 px-1.5 py-0.5 rounded-sm">
                응급의학과
              </span>
              <span className="text-xs text-slate-400 font-medium">근무표</span>
            </div>
            <h1 className="text-base font-bold text-slate-800 tracking-tight flex items-center gap-1">
              <span className="text-blue-600 underline decoration-amber-400 decoration-2 underline-offset-2">
                {targetUserName}
              </span>
              <span className="text-slate-600 text-sm font-normal">스케줄</span>
            </h1>
          </div>
        </div>

        {/* 우측 상단 액션: 엑셀 파일 업로드 버튼 */}
        <button
          onClick={onOpenUpload}
          className="flex items-center gap-1.5 bg-slate-900 hover:bg-slate-800 active:scale-95 text-white text-xs font-medium px-3 py-2 rounded-xl transition-all shadow-sm"
          title="엑셀 근무표 업로드 (.xlsx)"
        >
          <Upload className="w-3.5 h-3.5 text-blue-300" />
          <span>엑셀 업로드</span>
        </button>
      </div>

      {/* 2. 날짜 이동 네비게이션 & 타임프레임 셀렉터 */}
      <div className="px-4 py-2 flex items-center justify-between gap-2">
        {/* 월 이동 컨트롤러 */}
        <div className="flex items-center gap-1.5">
          <div className="flex items-center bg-white border border-slate-200 rounded-xl p-0.5 shadow-2xs">
            <button
              onClick={onPrevDate}
              aria-label="이전 기간"
              className="p-1.5 text-slate-600 hover:text-slate-900 active:bg-slate-100 rounded-lg transition-colors"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <button
              onClick={onToday}
              className="px-2 py-1 text-xs font-semibold text-slate-700 hover:text-blue-600 active:bg-slate-100 rounded-md transition-colors"
            >
              오늘
            </button>
            <button
              onClick={onNextDate}
              aria-label="다음 기간"
              className="p-1.5 text-slate-600 hover:text-slate-900 active:bg-slate-100 rounded-lg transition-colors"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>

          <div className="font-extrabold text-slate-800 text-sm tracking-tight">
            {year}년 {month}월
          </div>
        </div>

        {/* 뷰 모드 드롭다운 / 세그먼트 셀렉터 */}
        <div className="relative">
          <select
            value={viewMode}
            onChange={(e) => onViewModeChange(e.target.value as ViewMode)}
            className="appearance-none bg-white border border-slate-200 text-slate-800 font-semibold text-xs rounded-xl pl-3 pr-7 py-2 shadow-2xs focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            {VIEW_MODES.map((mode) => (
              <option key={mode.key} value={mode.key}>
                {mode.label}
              </option>
            ))}
          </select>
          <div className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400">
            <CalendarIcon className="w-3.5 h-3.5" />
          </div>
        </div>
      </div>

      {/* 3. 모바일 전용 빠른 필터 바: '박현우' 내 근무만 보기 토글 */}
      <div className="px-4 pb-2.5 flex items-center justify-between">
        <div className="flex items-center gap-1.5 text-2xs text-slate-500 truncate max-w-[200px]">
          <span className="inline-block w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
          <span className="truncate">{fileName || '기본 스케줄 데이터'}</span>
        </div>

        <button
          onClick={onToggleOnlyMyShifts}
          className={`flex items-center gap-1.5 text-xs font-semibold px-2.5 py-1.5 rounded-lg transition-all ${
            onlyMyShifts
              ? 'bg-amber-400 text-slate-950 shadow-sm shadow-amber-300 ring-2 ring-amber-300 scale-102'
              : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-50'
          }`}
        >
          <Sparkles className={`w-3.5 h-3.5 ${onlyMyShifts ? 'text-slate-950 fill-slate-950' : 'text-amber-500'}`} />
          <span>{targetUserName} 근무만</span>
        </button>
      </div>

      {/* 4. 타임프레임 스와이프 탭 바 (iOS 세그먼트 스타일) */}
      <div className="px-4 pb-2.5">
        <div className="grid grid-cols-5 gap-1 bg-slate-200/70 p-1 rounded-xl">
          {VIEW_MODES.map((mode) => {
            const isActive = viewMode === mode.key;
            return (
              <button
                key={mode.key}
                onClick={() => onViewModeChange(mode.key)}
                className={`py-1.5 text-xs font-semibold rounded-lg transition-all text-center ${
                  isActive
                    ? 'bg-white text-blue-700 shadow-sm font-bold scale-102'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {mode.shortLabel}
              </button>
            );
          })}
        </div>
      </div>
    </header>
  );
};
