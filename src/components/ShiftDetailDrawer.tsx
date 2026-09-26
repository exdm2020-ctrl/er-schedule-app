'use client';

import React, { useState } from 'react';
import { DaySchedule, ShiftSlot } from '@/types/schedule';
import { 
  X, 
  Clock, 
  Calendar, 
  Copy, 
  Check, 
  Sparkles, 
  Sun, 
  Moon, 
  Briefcase, 
  HelpCircle,
  Share2,
  AlertCircle
} from 'lucide-react';

interface ShiftDetailDrawerProps {
  daySchedule: DaySchedule | null;
  onClose: () => void;
  targetUserName: string;
}

export const ShiftDetailDrawer: React.FC<ShiftDetailDrawerProps> = ({
  daySchedule,
  onClose,
  targetUserName,
}) => {
  const [copied, setCopied] = useState(false);

  if (!daySchedule) return null;

  const { date, dayOfWeek, dayNumber, isWeekend, isHoliday, holidayName, shifts, hasTargetUser } = daySchedule;

  const dayNames = ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'];
  const dayName = dayNames[dayOfWeek];

  // 클립보드 복사 텍스트 포맷팅
  const handleCopySchedule = () => {
    let text = `[응급의학과 근무표] ${date} (${dayName}${holidayName ? ` - ${holidayName}` : ''})\n`;
    shifts.forEach((s) => {
      text += `• ${s.label} (${s.timeRange}): ${s.workers.join(', ')}\n`;
    });
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-slate-900/60 backdrop-blur-xs transition-opacity animate-in fade-in duration-200">
      {/* 바깥 영역 클릭 시 닫기 */}
      <div className="absolute inset-0" onClick={onClose} />

      {/* 하단 시트 카드 */}
      <div className="relative w-full max-w-md bg-white rounded-t-3xl sm:rounded-2xl shadow-2xl p-5 pb-safe z-10 max-h-[85vh] overflow-y-auto animate-in slide-in-from-bottom-6 duration-200">
        {/* 상단 드래그 인디케이터 (iOS 스타일 핸들) */}
        <div className="w-12 h-1.5 bg-slate-300 rounded-full mx-auto mb-4" />

        {/* 헤더 정보 */}
        <div className="flex items-start justify-between pb-3 border-b border-slate-100">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xl font-black text-slate-900 tracking-tight">
                {date}
              </span>
              <span
                className={`text-sm font-bold px-2 py-0.5 rounded-md ${
                  dayOfWeek === 0 || isHoliday
                    ? 'bg-rose-100 text-rose-700'
                    : dayOfWeek === 6
                    ? 'bg-blue-100 text-blue-700'
                    : 'bg-slate-100 text-slate-700'
                }`}
              >
                {dayName}
              </span>
              {holidayName && (
                <span className="text-xs font-bold bg-rose-500 text-white px-2 py-0.5 rounded-full">
                  {holidayName}
                </span>
              )}
            </div>

            <p className="text-xs text-slate-500 mt-1 flex items-center gap-1.5">
              <span>총 {shifts.length}개 근무조</span>
              <span>•</span>
              <span>총 {shifts.reduce((acc, s) => acc + s.workers.length, 0)}명 근무</span>
              {hasTargetUser && (
                <>
                  <span>•</span>
                  <span className="text-amber-600 font-extrabold flex items-center gap-0.5">
                    <Sparkles className="w-3 h-3 fill-amber-500" />
                    {targetUserName} 근무일
                  </span>
                </>
              )}
            </p>
          </div>

          <button
            onClick={onClose}
            className="p-1.5 rounded-full text-slate-400 hover:text-slate-700 bg-slate-100 hover:bg-slate-200 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* 박현우 특별 배너 (근무 시) */}
        {hasTargetUser && (
          <div className="mt-3 p-3 bg-gradient-to-r from-amber-400 via-amber-300 to-yellow-400 rounded-2xl text-slate-950 shadow-sm border border-amber-400/80 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <div className="w-8 h-8 rounded-xl bg-slate-950 text-amber-300 flex items-center justify-center font-black">
                <Sparkles className="w-4 h-4 fill-amber-300" />
              </div>
              <div>
                <span className="text-xs font-bold block opacity-85">내 스케줄 안내</span>
                <span className="text-sm font-black">
                  {daySchedule.targetUserShifts.map((s) => `${s.label} (${s.timeRange})`).join(' + ')}
                </span>
              </div>
            </div>
            <span className="text-2xs font-extrabold bg-slate-950 text-white px-2 py-1 rounded-lg">
              {daySchedule.targetUserShifts.reduce((acc, s) => acc + s.hours, 0)}시간 근무
            </span>
          </div>
        )}

        {/* 근무조 상세 리스트 */}
        <div className="mt-4 space-y-2.5">
          <div className="text-xs font-bold text-slate-500 uppercase tracking-wider">
            {isWeekend || isHoliday ? '주말/공휴일 근무 편성' : '평일 정규 근무 편성'}
          </div>

          {shifts.length === 0 ? (
            <div className="p-8 text-center bg-slate-50 rounded-2xl border border-dashed border-slate-200 text-slate-400 text-sm">
              편성된 근무 일정이 없습니다.
            </div>
          ) : (
            shifts.map((shift, idx) => {
              const isMyShift = shift.hasTargetUser;

              let ShiftIcon = Sun;
              if (shift.type === 'NIGHT') ShiftIcon = Moon;
              else if (shift.type === 'HELPER') ShiftIcon = HelpCircle;
              else if (shift.type.includes('MID')) ShiftIcon = Briefcase;

              return (
                <div
                  key={shift.id || idx}
                  className={`p-3.5 rounded-2xl border transition-all ${
                    isMyShift
                      ? 'bg-amber-50/80 border-2 border-amber-400 shadow-sm ring-1 ring-amber-300'
                      : 'bg-white border-slate-200 shadow-2xs'
                  }`}
                >
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <div
                        className={`w-7 h-7 rounded-lg flex items-center justify-center ${
                          isMyShift
                            ? 'bg-amber-400 text-slate-950 font-bold'
                            : shift.type === 'DAY'
                            ? 'bg-sky-100 text-sky-700'
                            : shift.type === 'NIGHT'
                            ? 'bg-indigo-100 text-indigo-700'
                            : 'bg-emerald-100 text-emerald-700'
                        }`}
                      >
                        <ShiftIcon className="w-4 h-4" />
                      </div>
                      <div>
                        <h4 className="text-sm font-bold text-slate-800 flex items-center gap-1.5">
                          <span>{shift.label}</span>
                          {isMyShift && (
                            <span className="text-2xs font-extrabold bg-amber-400 text-slate-950 px-1.5 py-0.2 rounded-sm shadow-2xs">
                              내 근무
                            </span>
                          )}
                        </h4>
                        <div className="flex items-center gap-1.5 text-xs text-slate-500 mt-0.5">
                          <Clock className="w-3.5 h-3.5 text-slate-400" />
                          <span>{shift.timeRange}</span>
                          <span className="text-slate-300">|</span>
                          <span className="font-semibold text-slate-600">{shift.hours}시간</span>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* 근무자 태그 목록 */}
                  <div className="mt-3 pt-2.5 border-t border-slate-100 flex flex-wrap gap-1.5">
                    {shift.workers.map((worker, wIdx) => {
                      const isTarget = worker === targetUserName;
                      return (
                        <span
                          key={wIdx}
                          className={`text-xs font-semibold px-2.5 py-1 rounded-lg flex items-center gap-1 ${
                            isTarget
                              ? 'bg-[#FDE047] text-slate-950 font-black shadow-sm ring-1 ring-yellow-400'
                              : 'bg-slate-100 text-slate-700'
                          }`}
                        >
                          {isTarget && <Sparkles className="w-3 h-3 fill-slate-950" />}
                          {worker}
                        </span>
                      );
                    })}
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* 하단 액션 버튼 바 */}
        <div className="mt-5 pt-3 border-t border-slate-100 flex gap-2">
          <button
            onClick={handleCopySchedule}
            className="flex-1 flex items-center justify-center gap-2 py-3 px-4 rounded-xl bg-slate-900 hover:bg-slate-800 active:scale-98 text-white font-bold text-xs transition-all shadow-sm"
          >
            {copied ? (
              <>
                <Check className="w-4 h-4 text-emerald-400" />
                <span className="text-emerald-300">일정 복사 완료!</span>
              </>
            ) : (
              <>
                <Share2 className="w-4 h-4 text-blue-300" />
                <span>일정 복사 (공유)</span>
              </>
            )}
          </button>

          <button
            onClick={onClose}
            className="py-3 px-5 rounded-xl bg-slate-100 hover:bg-slate-200 active:bg-slate-300 text-slate-700 font-bold text-xs transition-colors"
          >
            닫기
          </button>
        </div>
      </div>
    </div>
  );
};
