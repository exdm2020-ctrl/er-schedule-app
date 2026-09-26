'use client';

import React from 'react';
import { DaySchedule } from '@/types/schedule';
import { Sparkles, Clock, Sun, Moon, Briefcase, HelpCircle } from 'lucide-react';

interface WeekTimelineViewProps {
  currentDate: Date;
  daysMap: Record<string, DaySchedule>;
  targetUserName: string;
  onlyMyShifts: boolean;
  onSelectDay: (day: DaySchedule) => void;
}

export const WeekTimelineView: React.FC<WeekTimelineViewProps> = ({
  currentDate,
  daysMap,
  targetUserName,
  onlyMyShifts,
  onSelectDay,
}) => {
  // 현재 날짜가 포함된 주의 일요일(0) 찾기
  const curr = new Date(currentDate);
  const day = curr.getDay();
  const sunday = new Date(curr);
  sunday.setDate(curr.getDate() - day);

  const weekDates: Date[] = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(sunday);
    d.setDate(sunday.getDate() + i);
    weekDates.push(d);
  }

  const dayLabels = ['일', '월', '화', '수', '목', '금', '토'];

  return (
    <div className="p-3 space-y-3">
      {weekDates.map((dateObj, i) => {
        const y = dateObj.getFullYear();
        const m = String(dateObj.getMonth() + 1).padStart(2, '0');
        const d = String(dateObj.getDate()).padStart(2, '0');
        const dateStr = `${y}-${m}-${d}`;
        const daySchedule = daysMap[dateStr];

        const isToday =
          new Date().toISOString().slice(0, 10) === dateStr;
        const hasTargetUser = Boolean(daySchedule?.hasTargetUser);
        const shifts = daySchedule?.shifts || [];
        const isDimmed = onlyMyShifts && !hasTargetUser;

        const isSunday = i === 0;
        const isSaturday = i === 6;
        const isHoliday = daySchedule?.isHoliday;

        return (
          <div
            key={dateStr}
            onClick={() => daySchedule && onSelectDay(daySchedule)}
            className={`p-3.5 rounded-2xl border transition-all cursor-pointer select-none active:scale-[0.99] ${
              isDimmed
                ? 'opacity-35 bg-white border-slate-200'
                : hasTargetUser
                ? 'bg-amber-50/80 border-2 border-amber-400 shadow-sm ring-1 ring-amber-300 target-highlight-box'
                : isToday
                ? 'bg-blue-50/60 border-2 border-blue-400 shadow-2xs'
                : 'bg-white border-slate-200 hover:border-slate-300 shadow-2xs'
            }`}
          >
            {/* 요일 헤더 */}
            <div className="flex items-center justify-between pb-2 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <div
                  className={`w-9 h-9 rounded-xl flex flex-col items-center justify-center font-bold ${
                    isToday
                      ? 'bg-blue-600 text-white shadow-xs'
                      : isSunday || isHoliday
                      ? 'bg-rose-100 text-rose-700'
                      : isSaturday
                      ? 'bg-blue-100 text-blue-700'
                      : 'bg-slate-100 text-slate-700'
                  }`}
                >
                  <span className="text-[10px] leading-none uppercase">{dayLabels[i]}</span>
                  <span className="text-sm leading-none mt-0.5">{dateObj.getDate()}</span>
                </div>

                <div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs font-bold text-slate-800">{dateStr}</span>
                    {daySchedule?.holidayName && (
                      <span className="text-[10px] font-bold text-rose-600 bg-rose-50 border border-rose-200 px-1.5 py-0.2 rounded-full">
                        {daySchedule.holidayName}
                      </span>
                    )}
                  </div>
                  <span className="text-2xs text-slate-400">
                    {shifts.length > 0
                      ? `${shifts.length}개 근무조 (${shifts.reduce((acc, s) => acc + s.workers.length, 0)}명)`
                      : '근무 일정 없음'}
                  </span>
                </div>
              </div>

              {hasTargetUser && (
                <div className="flex items-center gap-1 bg-gradient-to-r from-amber-400 to-yellow-400 text-slate-950 font-black text-2xs px-2.5 py-1 rounded-full shadow-2xs">
                  <Sparkles className="w-3 h-3 fill-slate-950" />
                  <span>{targetUserName} 근무</span>
                </div>
              )}
            </div>

            {/* 근무조 리스트 */}
            <div className="mt-2.5 space-y-1.5">
              {shifts.length === 0 ? (
                <div className="py-2 text-center text-xs text-slate-300">편성된 스케줄이 없습니다.</div>
              ) : (
                shifts.map((shift, sIdx) => {
                  const isMyShift = shift.hasTargetUser;

                  let ShiftIcon = Sun;
                  if (shift.type === 'NIGHT') ShiftIcon = Moon;
                  else if (shift.type === 'HELPER') ShiftIcon = HelpCircle;
                  else if (shift.type.includes('MID')) ShiftIcon = Briefcase;

                  return (
                    <div
                      key={shift.id || sIdx}
                      className={`p-2 rounded-xl border flex items-center justify-between gap-2 text-xs ${
                        isMyShift
                          ? 'bg-amber-100/90 border-amber-300 ring-1 ring-amber-300 font-bold text-slate-950'
                          : 'bg-slate-50 border-slate-200/70 text-slate-700'
                      }`}
                    >
                      <div className="flex items-center gap-1.5 shrink-0">
                        <span
                          className={`w-5 h-5 rounded-md flex items-center justify-center text-[10px] ${
                            isMyShift
                              ? 'bg-slate-950 text-amber-300'
                              : 'bg-white border border-slate-200 text-slate-600'
                          }`}
                        >
                          <ShiftIcon className="w-3 h-3" />
                        </span>
                        <span className="font-bold text-xs">{shift.label}</span>
                        <span className="text-2xs text-slate-400 flex items-center gap-0.5">
                          <Clock className="w-2.5 h-2.5" />
                          {shift.timeRange}
                        </span>
                      </div>

                      <div className="truncate text-right">
                        {shift.workers.map((w, wIdx) => (
                          <span
                            key={wIdx}
                            className={`inline-block ${
                              w === targetUserName
                                ? 'bg-amber-400 text-slate-950 px-1.5 py-0.5 rounded-md font-black shadow-2xs'
                                : 'text-slate-600'
                            }`}
                          >
                            {wIdx > 0 && <span className="text-slate-400 mx-1">/</span>}
                            {w}
                          </span>
                        ))}
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
};
