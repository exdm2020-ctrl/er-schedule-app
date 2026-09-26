'use client';

import React, { useState } from 'react';
import { UserStats } from '@/types/schedule';
import { 
  Clock, 
  Sun, 
  Moon, 
  CalendarCheck, 
  ChevronDown, 
  ChevronUp, 
  Flame, 
  HelpCircle,
  Briefcase
} from 'lucide-react';

interface SummaryDashboardProps {
  stats: UserStats;
}

export const SummaryDashboard: React.FC<SummaryDashboardProps> = ({ stats }) => {
  const [isExpanded, setIsExpanded] = useState(false);

  return (
    <div className="mx-3 mt-3 mb-1 bg-white rounded-2xl border border-slate-200/90 shadow-2xs overflow-hidden transition-all">
      {/* 요약 헤더 (탭하면 펼쳐짐) */}
      <div 
        onClick={() => setIsExpanded(!isExpanded)}
        className="px-3.5 py-2.5 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors"
      >
        <div className="flex items-center gap-2">
          <div className="w-6 h-6 rounded-lg bg-amber-400 text-slate-950 flex items-center justify-center font-bold text-xs shadow-2xs">
            <Flame className="w-3.5 h-3.5 fill-slate-950" />
          </div>
          <div>
            <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
              <span>{stats.targetName} 이번 달 근무 요약</span>
              <span className="bg-amber-100 text-amber-900 font-extrabold text-2xs px-1.5 py-0.5 rounded-full">
                총 {stats.totalDays}일 ({stats.totalHours}시간)
              </span>
            </div>
          </div>
        </div>

        <button 
          type="button"
          aria-label={isExpanded ? "통계 접기" : "통계 펼치기"}
          className="text-slate-400 p-1"
        >
          {isExpanded ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
        </button>
      </div>

      {/* 미니 뱃지 띠 (접혀있을 때도 핵심 요약 노출) */}
      <div className="px-3.5 pb-2.5 pt-0.5 flex items-center gap-1.5 overflow-x-auto no-scrollbar">
        <div className="flex items-center gap-1 bg-sky-50 text-sky-800 px-2 py-0.5 rounded-md text-2xs font-semibold whitespace-nowrap">
          <Sun className="w-3 h-3 text-sky-600" />
          <span>데이 {stats.dayCount}회</span>
        </div>
        <div className="flex items-center gap-1 bg-emerald-50 text-emerald-800 px-2 py-0.5 rounded-md text-2xs font-semibold whitespace-nowrap">
          <Briefcase className="w-3 h-3 text-emerald-600" />
          <span>미드 {stats.midCount}회</span>
        </div>
        <div className="flex items-center gap-1 bg-teal-50 text-teal-800 px-2 py-0.5 rounded-md text-2xs font-semibold whitespace-nowrap">
          <HelpCircle className="w-3 h-3 text-teal-600" />
          <span>헬퍼 {stats.helperCount}회</span>
        </div>
        <div className="flex items-center gap-1 bg-indigo-50 text-indigo-800 px-2 py-0.5 rounded-md text-2xs font-semibold whitespace-nowrap">
          <Moon className="w-3 h-3 text-indigo-600" />
          <span>나이트 {stats.nightCount}회</span>
        </div>
        <div className="flex items-center gap-1 bg-rose-50 text-rose-800 px-2 py-0.5 rounded-md text-2xs font-semibold whitespace-nowrap">
          <CalendarCheck className="w-3 h-3 text-rose-600" />
          <span>주말/휴일 {stats.weekendHolidayCount}회</span>
        </div>
      </div>

      {/* 확장 상세 패널 */}
      {isExpanded && (
        <div className="px-3.5 pb-3 pt-2 border-t border-slate-100 bg-slate-50/50">
          <div className="grid grid-cols-2 gap-2 text-xs">
            <div className="bg-white p-2.5 rounded-xl border border-slate-200/80 shadow-2xs">
              <span className="text-slate-500 text-2xs block">평균 근무 시간</span>
              <span className="text-base font-bold text-slate-800">
                {stats.totalDays > 0 ? (stats.totalHours / stats.totalDays).toFixed(1) : 0}시간 / 일
              </span>
            </div>
            <div className="bg-white p-2.5 rounded-xl border border-slate-200/80 shadow-2xs">
              <span className="text-slate-500 text-2xs block">주말/공휴일 비중</span>
              <span className="text-base font-bold text-rose-600">
                {stats.totalDays > 0 ? Math.round((stats.weekendHolidayCount / stats.totalDays) * 100) : 0}%
              </span>
            </div>
          </div>
          <div className="mt-2 text-2xs text-slate-500 flex items-center justify-between">
            <span>* 평일: 데이(7h), 미드(10h), 나이트(7.5h)</span>
            <span>* 주말: 데이(8h), 헬퍼/미드(9h), 나이트(8h)</span>
          </div>
        </div>
      )}
    </div>
  );
};
