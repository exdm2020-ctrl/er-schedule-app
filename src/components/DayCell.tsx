'use client';

import React from 'react';
import { DaySchedule, ShiftSlot } from '@/types/schedule';
import { Sparkles, Sun, Moon, Briefcase, HelpCircle } from 'lucide-react';

interface DayCellProps {
  daySchedule?: DaySchedule;
  dayNumber: number;
  dateStr: string;
  isCurrentMonth: boolean;
  isToday: boolean;
  isSelected: boolean;
  onlyMyShifts: boolean;
  targetUserName: string;
  onSelect: (day: DaySchedule) => void;
}

export const DayCell: React.FC<DayCellProps> = ({
  daySchedule,
  dayNumber,
  dateStr,
  isCurrentMonth,
  isToday,
  isSelected,
  onlyMyShifts,
  targetUserName,
  onSelect,
}) => {
  // 이전/다음 달 비활성화 셀
  if (!isCurrentMonth) {
    return (
      <div className="min-h-[88px] sm:min-h-[96px] p-0.5 bg-slate-100/40 border border-slate-200/40 rounded-xl opacity-30 pointer-events-none flex flex-col justify-between">
        <span className="text-[10px] font-semibold text-slate-400 pl-1">{dayNumber}</span>
      </div>
    );
  }

  const shifts = daySchedule?.shifts || [];
  const hasTargetUser = Boolean(daySchedule?.hasTargetUser);
  const isWeekend = daySchedule?.isWeekend;
  const isHoliday = daySchedule?.isHoliday;
  const holidayName = daySchedule?.holidayName;

  // 요일 및 휴일 색상 정의
  const isSunday = daySchedule?.dayOfWeek === 0;
  const isSaturday = daySchedule?.dayOfWeek === 6;

  let dateColor = 'text-slate-800';
  if (isSunday || isHoliday) {
    dateColor = 'text-rose-600 font-extrabold';
  } else if (isSaturday) {
    dateColor = 'text-blue-600 font-bold';
  }

  // '내 근무만 보기' 모드일 때 필터링
  const isDimmed = onlyMyShifts && !hasTargetUser;

  const handleClick = () => {
    if (daySchedule) {
      onSelect(daySchedule);
    }
  };

  return (
    <div
      onClick={handleClick}
      className={`relative min-h-[92px] sm:min-h-[102px] p-1 rounded-xl transition-all cursor-pointer flex flex-col justify-between select-none active:scale-[0.98] overflow-hidden ${
        isDimmed
          ? 'bg-slate-50/60 border border-slate-200/50 opacity-30'
          : hasTargetUser
          ? 'bg-amber-50/90 border-2 border-amber-400 shadow-md ring-2 ring-yellow-400/80 target-highlight-box'
          : isSelected
          ? 'bg-blue-50/60 border-2 border-blue-500 shadow-xs'
          : 'bg-white border border-slate-200 hover:border-slate-300 shadow-2xs'
      }`}
    >
      {/* 1. 상단 날짜 번호 및 휴일 배지 & 박현우 근무 표시 */}
      <div className="flex items-center justify-between gap-0.5 mb-1 w-full min-w-0">
        <div className="flex items-center gap-1 min-w-0">
          <span
            className={`inline-flex items-center justify-center text-[11px] font-black w-5 h-5 rounded-full shrink-0 ${dateColor} ${
              isToday ? 'bg-blue-600 text-white shadow-2xs' : ''
            }`}
          >
            {dayNumber}
          </span>

          {/* 공휴일/추석 등 한글 메모 배지 */}
          {holidayName && (
            <span className="text-[8.5px] font-black text-rose-600 bg-rose-50 border border-rose-200 px-1 py-0.2 rounded-xs truncate max-w-[48px] leading-none shrink-0">
              {holidayName}
            </span>
          )}
        </div>

        {/* '박현우' 근무일 상단 엠블럼 */}
        {hasTargetUser && (
          <div className="flex items-center gap-0.5 bg-[#FDE047] text-slate-950 font-black text-[8.5px] px-1.5 py-0.5 rounded-full shadow-xs border border-yellow-400 shrink-0">
            <Sparkles className="w-2.5 h-2.5 fill-slate-950" />
            <span>MY</span>
          </div>
        )}
      </div>

      {/* 2. 근무조 리스트: 각 근무조를 작은 블록(Flexbox)으로 만들고, flex-wrap & break-words로 글자 잘림 방지 */}
      <div className="flex-1 flex flex-col gap-1 overflow-hidden w-full min-w-0">
        {shifts.length === 0 ? (
          <div className="h-full flex items-center justify-center">
            <span className="text-[9px] text-slate-300">근무없음</span>
          </div>
        ) : (
          shifts.map((shift, idx) => {
            const hasMyName = shift.hasTargetUser;

            // 근무조 배지 색상
            let badgeBg = 'bg-slate-100 text-slate-700';
            if (shift.type === 'DAY') badgeBg = 'bg-sky-100 text-sky-800 font-bold';
            else if (shift.type === 'NIGHT') badgeBg = 'bg-indigo-100 text-indigo-800 font-bold';
            else if (shift.type === 'HELPER') badgeBg = 'bg-teal-100 text-teal-800 font-bold';
            else if (shift.type.includes('MID')) badgeBg = 'bg-emerald-100 text-emerald-800 font-bold';

            return (
              <div
                key={shift.id || idx}
                className={`w-full rounded-md p-0.5 flex flex-wrap items-center gap-0.5 text-[9.5px] leading-tight min-w-0 ${
                  hasMyName
                    ? 'bg-amber-100/90 border border-amber-300 ring-1 ring-amber-300/50'
                    : 'bg-slate-50 border border-slate-100'
                }`}
              >
                {/* 근무조 명칭 태그 (데이, 미드, 헬퍼, 나이트) */}
                <span
                  className={`text-[8px] font-black px-1 py-0.2 rounded-xs shrink-0 ${badgeBg}`}
                >
                  {shift.shortLabel.slice(0, 2)}
                </span>

                {/* 근무자 목록: flex-wrap과 break-words로 절대 화면 밖으로 잘리지 않고 자동 줄바꿈 */}
                <div className="flex flex-wrap items-center gap-0.5 min-w-0 flex-1">
                  {shift.workers.map((worker, wIdx) => {
                    const isTarget = worker === targetUserName;
                    return (
                      <React.Fragment key={wIdx}>
                        {isTarget ? (
                          // 박현우: 노란색 배경(#FDE047)에 굵은 글씨로 배지 처리!
                          <span className="inline-block bg-[#FDE047] text-slate-950 font-black text-[9px] px-1 py-0.2 rounded-sm border border-yellow-400 shadow-2xs tracking-tight break-words">
                            {worker}
                          </span>
                        ) : (
                          <span className="text-slate-600 text-[8.5px] font-medium break-words">
                            {worker}
                          </span>
                        )}
                        {wIdx < shift.workers.length - 1 && (
                          <span className="text-slate-300 text-[7px] select-none">/</span>
                        )}
                      </React.Fragment>
                    );
                  })}
                </div>
              </div>
            );
          })
        )}
      </div>

      {/* 3. 하단 여백 및 내 근무 요약 태그 */}
      {hasTargetUser && daySchedule && (
        <div className="mt-0.5 pt-0.5 border-t border-amber-200/80 flex items-center justify-between text-[8px] text-amber-900 font-extrabold">
          <span>{daySchedule.targetUserShifts.map(s => s.shortLabel).join('/')}</span>
          <span>{daySchedule.targetUserShifts.reduce((acc, s) => acc + s.hours, 0)}h</span>
        </div>
      )}
    </div>
  );
};
