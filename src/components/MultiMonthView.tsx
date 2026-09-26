'use client';

import React from 'react';
import { DaySchedule, ScheduleData } from '@/types/schedule';
import { DayCell } from './DayCell';
import { Sparkles, Calendar } from 'lucide-react';

interface MultiMonthViewProps {
  startDate: Date;
  monthCount: number; // 3, 6, 12
  scheduleData: ScheduleData;
  targetUserName: string;
  onlyMyShifts: boolean;
  onSelectDay: (day: DaySchedule) => void;
  onJumpToMonth: (date: Date) => void;
}

export const MultiMonthView: React.FC<MultiMonthViewProps> = ({
  startDate,
  monthCount,
  scheduleData,
  targetUserName,
  onlyMyShifts,
  onSelectDay,
  onJumpToMonth,
}) => {
  const months: { year: number; month: number }[] = [];
  const startYear = startDate.getFullYear();
  const startMonth = startDate.getMonth(); // 0-indexed

  for (let i = 0; i < monthCount; i++) {
    const d = new Date(startYear, startMonth + i, 1);
    months.push({ year: d.getFullYear(), month: d.getMonth() + 1 });
  }

  const weekHeaders = ['일', '월', '화', '수', '목', '금', '토'];

  return (
    <div className="p-3 space-y-6">
      {months.map(({ year, month }) => {
        const yearMonthStr = `${year}-${String(month).padStart(2, '0')}`;
        const firstDayOfWeek = new Date(year, month - 1, 1).getDay();
        const daysInMonth = new Date(year, month, 0).getDate();

        // 이번 달 박현우 총 근무 일수 카운트
        let myShiftsInThisMonth = 0;
        for (let d = 1; d <= daysInMonth; d++) {
          const dateKey = `${yearMonthStr}-${String(d).padStart(2, '0')}`;
          if (scheduleData.days[dateKey]?.hasTargetUser) {
            myShiftsInThisMonth++;
          }
        }

        return (
          <div
            key={yearMonthStr}
            className="bg-white rounded-3xl p-3.5 border border-slate-200 shadow-2xs"
          >
            {/* 월 헤더 */}
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <div className="w-7 h-7 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center font-bold text-xs">
                  <Calendar className="w-3.5 h-3.5" />
                </div>
                <h3 className="font-extrabold text-sm text-slate-800">
                  {year}년 {month}월
                </h3>
              </div>

              <div className="flex items-center gap-2">
                {myShiftsInThisMonth > 0 && (
                  <span className="flex items-center gap-1 text-2xs font-black bg-[#FDE047] text-slate-950 px-2 py-0.5 rounded-full border border-yellow-400 shadow-2xs">
                    <Sparkles className="w-3 h-3 fill-slate-950" />
                    <span>{targetUserName} {myShiftsInThisMonth}일 근무</span>
                  </span>
                )}
                <button
                  onClick={() => onJumpToMonth(new Date(year, month - 1, 1))}
                  className="text-2xs text-blue-600 hover:underline font-semibold"
                >
                  월간뷰 이동
                </button>
              </div>
            </div>

            {/* 요일 헤더 */}
            <div className="grid grid-cols-7 gap-1 pt-2 pb-1.5 text-center text-2xs font-bold text-slate-400">
              {weekHeaders.map((day, idx) => (
                <div
                  key={day}
                  className={`${idx === 0 ? 'text-rose-500' : idx === 6 ? 'text-blue-500' : ''}`}
                >
                  {day}
                </div>
              ))}
            </div>

            {/* 그리드 날짜 셀 */}
            <div className="grid grid-cols-7 gap-1">
              {/* 이전 달 빈칸 */}
              {Array.from({ length: firstDayOfWeek }).map((_, idx) => (
                <div key={`empty-${idx}`} className="min-h-[70px] bg-slate-50/50 rounded-xl" />
              ))}

              {/* 해당 월 날짜들 */}
              {Array.from({ length: daysInMonth }).map((_, idx) => {
                const dayNum = idx + 1;
                const dateKey = `${yearMonthStr}-${String(dayNum).padStart(2, '0')}`;
                const daySchedule = scheduleData.days[dateKey];
                const isToday =
                  new Date().toISOString().slice(0, 10) === dateKey;

                return (
                  <DayCell
                    key={dateKey}
                    dayNumber={dayNum}
                    dateStr={dateKey}
                    daySchedule={daySchedule}
                    isCurrentMonth={true}
                    isToday={isToday}
                    isSelected={false}
                    onlyMyShifts={onlyMyShifts}
                    targetUserName={targetUserName}
                    onSelect={onSelectDay}
                  />
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
};
