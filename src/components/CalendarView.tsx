'use client';

import React, { useState } from 'react';
import { ScheduleData, DaySchedule, ViewMode } from '@/types/schedule';
import { DayCell } from './DayCell';
import { WeekTimelineView } from './WeekTimelineView';
import { MultiMonthView } from './MultiMonthView';

interface CalendarViewProps {
  currentDate: Date;
  scheduleData: ScheduleData;
  viewMode: ViewMode;
  onlyMyShifts: boolean;
  targetUserName: string;
  onSelectDay: (day: DaySchedule) => void;
  onJumpToMonth: (date: Date) => void;
}

export const CalendarView: React.FC<CalendarViewProps> = ({
  currentDate,
  scheduleData,
  viewMode,
  onlyMyShifts,
  targetUserName,
  onSelectDay,
  onJumpToMonth,
}) => {
  const [selectedDateStr, setSelectedDateStr] = useState<string | null>(null);

  const year = currentDate.getFullYear();
  const month = currentDate.getMonth() + 1; // 1-indexed

  // 1. 주간 뷰 모드(Week)
  if (viewMode === 'week') {
    return (
      <WeekTimelineView
        currentDate={currentDate}
        daysMap={scheduleData.days}
        targetUserName={targetUserName}
        onlyMyShifts={onlyMyShifts}
        onSelectDay={(day) => {
          setSelectedDateStr(day.date);
          onSelectDay(day);
        }}
      />
    );
  }

  // 2. 다중 월 뷰 (3M, 6M, 1Year)
  if (viewMode === '3months' || viewMode === '6months' || viewMode === '1year') {
    const monthCount = viewMode === '3months' ? 3 : viewMode === '6months' ? 6 : 12;
    return (
      <MultiMonthView
        startDate={currentDate}
        monthCount={monthCount}
        scheduleData={scheduleData}
        targetUserName={targetUserName}
        onlyMyShifts={onlyMyShifts}
        onSelectDay={(day) => {
          setSelectedDateStr(day.date);
          onSelectDay(day);
        }}
        onJumpToMonth={onJumpToMonth}
      />
    );
  }

  // 3. 1개월(Month) 표준 달력 뷰
  const firstDayOfWeek = new Date(year, month - 1, 1).getDay(); // 0(일) ~ 6(토)
  const daysInCurrentMonth = new Date(year, month, 0).getDate();
  const daysInPrevMonth = new Date(year, month - 1, 0).getDate();

  const weekHeaders = [
    { label: '일', isWeekend: true, color: 'text-rose-500' },
    { label: '월', isWeekend: false, color: 'text-slate-600' },
    { label: '화', isWeekend: false, color: 'text-slate-600' },
    { label: '수', isWeekend: false, color: 'text-slate-600' },
    { label: '목', isWeekend: false, color: 'text-slate-600' },
    { label: '금', isWeekend: false, color: 'text-slate-600' },
    { label: '토', isWeekend: true, color: 'text-blue-500' },
  ];

  // 달력 격자 셀 계산
  const calendarCells: {
    dayNumber: number;
    dateStr: string;
    isCurrentMonth: boolean;
  }[] = [];

  // 이전 달의 끝 날짜들
  for (let i = firstDayOfWeek - 1; i >= 0; i--) {
    const prevDayNum = daysInPrevMonth - i;
    const prevDate = new Date(year, month - 2, prevDayNum);
    const dateStr = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, '0')}-${String(prevDayNum).padStart(2, '0')}`;
    calendarCells.push({
      dayNumber: prevDayNum,
      dateStr,
      isCurrentMonth: false,
    });
  }

  // 이번 달의 날짜들
  for (let d = 1; d <= daysInCurrentMonth; d++) {
    const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    calendarCells.push({
      dayNumber: d,
      dateStr,
      isCurrentMonth: true,
    });
  }

  // 다음 달의 시작 날짜들 (7의 배수로 꽉 채움)
  const remainingCells = 7 - (calendarCells.length % 7);
  if (remainingCells < 7) {
    for (let d = 1; d <= remainingCells; d++) {
      const nextDate = new Date(year, month, d);
      const dateStr = `${nextDate.getFullYear()}-${String(nextDate.getMonth() + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      calendarCells.push({
        dayNumber: d,
        dateStr,
        isCurrentMonth: false,
      });
    }
  }

  const todayStr = new Date().toISOString().slice(0, 10);

  return (
    <div className="p-3">
      {/* 요일 헤더 */}
      <div className="grid grid-cols-7 gap-1 mb-1 text-center">
        {weekHeaders.map((header) => (
          <div
            key={header.label}
            className={`py-1 text-xs font-bold ${header.color}`}
          >
            {header.label}
          </div>
        ))}
      </div>

      {/* 달력 날짜 그리드 */}
      <div className="grid grid-cols-7 gap-1">
        {calendarCells.map((cell) => {
          const daySchedule = scheduleData.days[cell.dateStr];
          const isToday = cell.dateStr === todayStr;
          const isSelected = cell.dateStr === selectedDateStr;

          return (
            <DayCell
              key={cell.dateStr}
              dayNumber={cell.dayNumber}
              dateStr={cell.dateStr}
              daySchedule={daySchedule}
              isCurrentMonth={cell.isCurrentMonth}
              isToday={isToday}
              isSelected={isSelected}
              onlyMyShifts={onlyMyShifts}
              targetUserName={targetUserName}
              onSelect={(day) => {
                setSelectedDateStr(day.date);
                onSelectDay(day);
              }}
            />
          );
        })}
      </div>
    </div>
  );
};
