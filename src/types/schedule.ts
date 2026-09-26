export type ShiftType = 
  | 'DAY'         // 데이
  | 'MID'         // 미드
  | 'MID1'        // 미드1
  | 'MID2'        // 미드2
  | 'HELPER'      // 주말 헬퍼
  | 'NIGHT'       // 나이트
  | 'CUSTOM';

export interface ShiftSlot {
  id: string;
  type: ShiftType;
  label: string;           // "평일 데이", "주말 헬퍼", "미드1", "나이트" 등
  shortLabel: string;      // "데이", "미드1", "헬퍼", "나이트"
  timeRange: string;       // "08:00 - 15:00" 등
  hours: number;           // 근무 시간 (예: 7, 10, 7.5, 8, 9)
  workers: string[];       // 근무자 이름 목록
  hasTargetUser: boolean;  // '박현우' 포함 여부
  colorTheme: {
    bg: string;
    border: string;
    text: string;
    badgeBg: string;
  };
}

export interface DaySchedule {
  date: string;            // 'YYYY-MM-DD'
  dayOfWeek: number;       // 0: 일, 1: 월, ... 6: 토
  dayNumber: number;       // 1 ~ 31
  isWeekend: boolean;
  isHoliday: boolean;
  holidayName: string | null;
  rawText: string;         // 원본 엑셀 셀 내용
  shifts: ShiftSlot[];
  hasTargetUser: boolean;  // 박현우 포함 여부
  targetUserShifts: ShiftSlot[]; // 박현우가 참여한 근무
}

export type ViewMode = 'week' | 'month' | '3months' | '6months' | '1year';

export interface ScheduleData {
  title: string;
  fileName: string;
  yearMonth: string;       // 기본 연월 (예: '2024-10')
  startDate: string;
  endDate: string;
  days: Record<string, DaySchedule>; // 'YYYY-MM-DD' 키
  allWorkers: string[];
  parsedAt: string;
}

export interface UserStats {
  targetName: string;
  totalDays: number;
  totalHours: number;
  dayCount: number;
  midCount: number;
  helperCount: number;
  nightCount: number;
  weekendHolidayCount: number;
}
