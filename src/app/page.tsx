'use client';

import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  ChevronLeft, 
  ChevronRight, 
  Upload, 
  X, 
  Calendar as CalendarIcon, 
  Sparkles,
  Share2,
  Check,
  Download,
  AlertCircle
} from 'lucide-react';

// ==========================================
// 1. 타입 정의
// ==========================================
export type ShiftCode = 'D' | 'M1' | 'M2' | 'M' | 'H' | 'N';

export interface ShiftItem {
  code: ShiftCode;
  name: string;        // "데이", "미드1", "미드2", "주말 헬퍼" 등
  time: string;        // "08:00 - 15:00"
  workers: string[];   // ["김민준", "박현우"]
  hasTargetUser: boolean;
}

export interface ParsedDay {
  date: string;        // "2026-09-25"
  yearMonth: string;   // "2026-09"
  dayNum: number;      // 25
  dayOfWeek: number;   // 0(일) ~ 6(토)
  isWeekend: boolean;
  holidayNote: string | null;
  shifts: ShiftItem[];
  hasTargetUser: boolean;
  rawText: string;
}

export type ViewMode = '1month' | '3months' | '6months' | '1year';

const TARGET_USER_NAME = '박현우';

// ==========================================
// 2. 엑셀 파싱 핵심 알고리즘 (지정 규격 완벽 준수)
// ==========================================
function parseExcelData(fileBuffer: ArrayBuffer): ParsedDay[] {
  const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: false });
  if (!workbook.SheetNames.length) return [];

  // sheet[0] 읽기
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: '',
  });

  let currentYearMonth: string | null = null;
  const parsedData: ParsedDay[] = [];
  const processedDateKeys = new Set<string>();

  for (let r = 0; r < rawRows.length; r++) {
    const row = rawRows[r] || [];

    // [월 식별]: 셀 텍스트에 /(\d{4})\.(\d{1,2})월/ 매칭 시 (예: "2026.09월", "2026.9월")
    for (let c = 0; c < row.length; c++) {
      const cellStr = String(row[c] || '').trim();
      const monthMatch = cellStr.match(/(\d{4})\.(\d{1,2})월/);
      if (monthMatch) {
        const y = monthMatch[1];
        const m = String(parseInt(monthMatch[2], 10)).padStart(2, '0');
        currentYearMonth = `${y}-${m}`;
        break;
      }
    }

    if (!currentYearMonth) continue;

    // [요일 식별]: row에 "일", "월", "화"가 나타나면
    const rowStrArr = row.map(cell => String(cell || '').trim());
    const hasDaysOfWeek = rowStrArr.includes('일') && rowStrArr.includes('월') && rowStrArr.includes('화');

    if (hasDaysOfWeek) {
      // 다음 row는 '날짜', 그 다음 row는 '근무자' 행으로 인지
      let nextR = r + 1;

      while (nextR < rawRows.length) {
        const dateRow = rawRows[nextR] || [];
        const shiftRow = rawRows[nextR + 1] || [];

        // 새로운 월 헤더가 나오면 현재 월 주(Week) 파싱 중단
        const hasNextMonthHeader = dateRow.some(cell => String(cell || '').match(/(\d{4})\.(\d{1,2})월/));
        if (hasNextMonthHeader) break;

        // 다른 요일 행이 끼어있는 경우 스킵
        const isAnotherDayOfWeekRow = dateRow.some(c => String(c).trim() === '일') && dateRow.some(c => String(c).trim() === '월');
        if (isAnotherDayOfWeekRow) {
          nextR++;
          continue;
        }

        let foundValidDateInThisRow = false;

        for (let col = 0; col < 7; col++) {
          const dateCellVal = String(dateRow[col] || '').trim();
          // [날짜 추출]: 날짜 셀 텍스트(예: "25(추석)", "3")에서 정규식 /\d+/을 통해 숫자만 추출
          const dayMatch = dateCellVal.match(/\d+/);
          if (!dayMatch) continue;

          foundValidDateInThisRow = true;
          const dayNum = parseInt(dayMatch[0], 10);
          if (dayNum < 1 || dayNum > 31) continue;

          // 공휴일 메모 (예: "25(추석연휴)" -> "추석연휴")
          const memoMatch = dateCellVal.match(/\((.*?)\)/);
          const holidayNote = memoMatch ? memoMatch[1].trim() : null;

          const dateStr = `${currentYearMonth}-${String(dayNum).padStart(2, '0')}`;
          const [yStr, mStr] = currentYearMonth.split('-');
          const dateObj = new Date(parseInt(yStr, 10), parseInt(mStr, 10) - 1, dayNum);
          const dayOfWeek = dateObj.getDay();
          const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
          const isWeekendOrHol = isWeekend || Boolean(holidayNote);

          // [근무자 줄바꿈(\n) 분리 규칙 적용]
          const shiftCellVal = String(shiftRow[col] || '').trim();
          const lines = shiftCellVal.split(/\r?\n/).map(l => l.trim()).filter(Boolean);

          const shifts: ShiftItem[] = [];

          if (isWeekendOrHol) {
            // * 주말(토,일) 및 공휴일: 무조건 배열 길이 3이므로 [0]=D, [1]=M/H, [2]=N
            // [0] 주말 데이 (D)
            if (lines[0]) {
              const workers = lines[0].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
              shifts.push({
                code: 'D',
                name: '주말 데이',
                time: '08:00 - 16:00',
                workers,
                hasTargetUser: workers.includes(TARGET_USER_NAME),
              });
            }
            // [1] 주말 헬퍼/미드 (M/H)
            if (lines[1]) {
              const lineWorkers = lines[1].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
              if (lineWorkers.length >= 2) {
                const helperWorker = [lineWorkers[0]];
                const midWorker = lineWorkers.slice(1);
                shifts.push({
                  code: 'H',
                  name: '주말 헬퍼',
                  time: '14:00 - 23:00',
                  workers: helperWorker,
                  hasTargetUser: helperWorker.includes(TARGET_USER_NAME),
                });
                shifts.push({
                  code: 'M',
                  name: '주말 미드',
                  time: '15:00 - 24:00',
                  workers: midWorker,
                  hasTargetUser: midWorker.includes(TARGET_USER_NAME),
                });
              } else {
                shifts.push({
                  code: 'H',
                  name: '주말 헬퍼/미드',
                  time: '14:00 - 24:00',
                  workers: lineWorkers,
                  hasTargetUser: lineWorkers.includes(TARGET_USER_NAME),
                });
              }
            }
            // [2] 주말 나이트 (N)
            if (lines[2]) {
              const workers = lines[2].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
              shifts.push({
                code: 'N',
                name: '주말 나이트',
                time: '00:00 - 익일 08:00',
                workers,
                hasTargetUser: workers.includes(TARGET_USER_NAME),
              });
            }
          } else {
            // * 평일(월~금)
            if (lines.length >= 3) {
              // 배열 길이 3이면 [0]=D, [1]=M1/M2, [2]=N
              if (lines[0]) {
                const workers = lines[0].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
                shifts.push({
                  code: 'D',
                  name: '평일 데이',
                  time: '08:00 - 15:00',
                  workers,
                  hasTargetUser: workers.includes(TARGET_USER_NAME),
                });
              }
              if (lines[1]) {
                const midWorkers = lines[1].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
                if (midWorkers.length >= 2) {
                  const m1 = [midWorkers[0]];
                  const m2 = midWorkers.slice(1);
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드1',
                    time: '14:00 - 24:00',
                    workers: m1,
                    hasTargetUser: m1.includes(TARGET_USER_NAME),
                  });
                  shifts.push({
                    code: 'M2',
                    name: '평일 미드2',
                    time: '14:00 - 24:00',
                    workers: m2,
                    hasTargetUser: m2.includes(TARGET_USER_NAME),
                  });
                } else {
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드',
                    time: '14:00 - 24:00',
                    workers: midWorkers,
                    hasTargetUser: midWorkers.includes(TARGET_USER_NAME),
                  });
                }
              }
              if (lines[2]) {
                const workers = lines[2].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
                shifts.push({
                  code: 'N',
                  name: '평일 나이트',
                  time: '00:00 - 익일 07:30',
                  workers,
                  hasTargetUser: workers.includes(TARGET_USER_NAME),
                });
              }
            } else if (lines.length === 2) {
              // 배열 길이 2이면 [0]=M1/M2, [1]=N (데이 없음)
              if (lines[0]) {
                const midWorkers = lines[0].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
                if (midWorkers.length >= 2) {
                  const m1 = [midWorkers[0]];
                  const m2 = midWorkers.slice(1);
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드1',
                    time: '14:00 - 24:00',
                    workers: m1,
                    hasTargetUser: m1.includes(TARGET_USER_NAME),
                  });
                  shifts.push({
                    code: 'M2',
                    name: '평일 미드2',
                    time: '14:00 - 24:00',
                    workers: m2,
                    hasTargetUser: m2.includes(TARGET_USER_NAME),
                  });
                } else {
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드',
                    time: '14:00 - 24:00',
                    workers: midWorkers,
                    hasTargetUser: midWorkers.includes(TARGET_USER_NAME),
                  });
                }
              }
              if (lines[1]) {
                const workers = lines[1].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
                shifts.push({
                  code: 'N',
                  name: '평일 나이트',
                  time: '00:00 - 익일 07:30',
                  workers,
                  hasTargetUser: workers.includes(TARGET_USER_NAME),
                });
              }
            } else if (lines.length === 1) {
              const workers = lines[0].split(/[\/,]/).map(w => w.trim()).filter(Boolean);
              shifts.push({
                code: 'M1',
                name: '평일 근무',
                time: '14:00 - 24:00',
                workers,
                hasTargetUser: workers.includes(TARGET_USER_NAME),
              });
            }
          }

          if (!processedDateKeys.has(dateStr)) {
            processedDateKeys.add(dateStr);
            parsedData.push({
              date: dateStr,
              yearMonth: currentYearMonth,
              dayNum,
              dayOfWeek,
              isWeekend,
              holidayNote,
              shifts,
              hasTargetUser: shifts.some(s => s.hasTargetUser),
              rawText: shiftCellVal,
            });
          }
        }

        if (!foundValidDateInThisRow) {
          nextR++;
        } else {
          nextR += 2;
        }
      }

      r = nextR - 1;
    }
  }

  return parsedData;
}

// ==========================================
// 3. 초기 탑재용 2026년 9~11월 3개월 샘플 생성
// ==========================================
function generateInitialSampleData(): ParsedDay[] {
  const doctors = ['박현우', '김민준', '이서연', '정유진', '최준호', '윤도윤', '강예은', '임재현'];
  const months = [
    { ym: '2026-09', days: 30 },
    { ym: '2026-10', days: 31 },
    { ym: '2026-11', days: 30 },
  ];

  const list: ParsedDay[] = [];

  months.forEach(({ ym, days }) => {
    const [y, m] = ym.split('-').map(Number);

    for (let d = 1; d <= days; d++) {
      const dateStr = `${ym}-${String(d).padStart(2, '0')}`;
      const dateObj = new Date(y, m - 1, d);
      const dayOfWeek = dateObj.getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

      let holidayNote: string | null = null;
      if (ym === '2026-09') {
        if (d === 24) holidayNote = '추석연휴';
        if (d === 25) holidayNote = '추석';
        if (d === 26) holidayNote = '추석연휴';
      } else if (ym === '2026-10') {
        if (d === 3) holidayNote = '개천절';
        if (d === 9) holidayNote = '한글날';
      }

      const isWeekendOrHol = isWeekend || Boolean(holidayNote);
      const seed = (y * 365 + m * 31 + d) % doctors.length;
      const getDoc = (offset: number) => doctors[(seed + offset) % doctors.length];

      const shifts: ShiftItem[] = [];

      if (isWeekendOrHol) {
        // [0]=D, [1]=M/H, [2]=N
        const dWorkers = [getDoc(0)];
        const hWorkers = [getDoc(1)];
        const mWorkers = [getDoc(2)];
        const nWorkers = [getDoc(3)];

        shifts.push({ code: 'D', name: '주말 데이', time: '08:00 - 16:00', workers: dWorkers, hasTargetUser: dWorkers.includes(TARGET_USER_NAME) });
        shifts.push({ code: 'H', name: '주말 헬퍼', time: '14:00 - 23:00', workers: hWorkers, hasTargetUser: hWorkers.includes(TARGET_USER_NAME) });
        shifts.push({ code: 'M', name: '주말 미드', time: '15:00 - 24:00', workers: mWorkers, hasTargetUser: mWorkers.includes(TARGET_USER_NAME) });
        shifts.push({ code: 'N', name: '주말 나이트', time: '00:00 - 익일 08:00', workers: nWorkers, hasTargetUser: nWorkers.includes(TARGET_USER_NAME) });
      } else {
        const isNoDay = (d % 7 === 3 || d % 7 === 5);
        if (isNoDay) {
          // [0]=M1/M2, [1]=N
          const m1Workers = [getDoc(1)];
          const m2Workers = [getDoc(2)];
          const nWorkers = [getDoc(4)];
          shifts.push({ code: 'M1', name: '평일 미드1', time: '14:00 - 24:00', workers: m1Workers, hasTargetUser: m1Workers.includes(TARGET_USER_NAME) });
          shifts.push({ code: 'M2', name: '평일 미드2', time: '14:00 - 24:00', workers: m2Workers, hasTargetUser: m2Workers.includes(TARGET_USER_NAME) });
          shifts.push({ code: 'N', name: '평일 나이트', time: '00:00 - 익일 07:30', workers: nWorkers, hasTargetUser: nWorkers.includes(TARGET_USER_NAME) });
        } else {
          // [0]=D, [1]=M1/M2, [2]=N
          const dWorkers = [getDoc(0)];
          const m1Workers = [getDoc(1)];
          const m2Workers = [getDoc(2)];
          const nWorkers = [getDoc(3)];
          shifts.push({ code: 'D', name: '평일 데이', time: '08:00 - 15:00', workers: dWorkers, hasTargetUser: dWorkers.includes(TARGET_USER_NAME) });
          shifts.push({ code: 'M1', name: '평일 미드1', time: '14:00 - 24:00', workers: m1Workers, hasTargetUser: m1Workers.includes(TARGET_USER_NAME) });
          shifts.push({ code: 'M2', name: '평일 미드2', time: '14:00 - 24:00', workers: m2Workers, hasTargetUser: m2Workers.includes(TARGET_USER_NAME) });
          shifts.push({ code: 'N', name: '평일 나이트', time: '00:00 - 익일 07:30', workers: nWorkers, hasTargetUser: nWorkers.includes(TARGET_USER_NAME) });
        }
      }

      list.push({
        date: dateStr,
        yearMonth: ym,
        dayNum: d,
        dayOfWeek,
        isWeekend,
        holidayNote,
        shifts,
        hasTargetUser: shifts.some(s => s.hasTargetUser),
        rawText: shifts.map(s => `${s.code} ${s.workers.join('/')}`).join('\n'),
      });
    }
  });

  return list;
}

// ==========================================
// 4. 메인 컴포넌트 (100% 다크 모드)
// ==========================================
export default function ERSchedulePage() {
  // 상태 관리
  const [scheduleList, setScheduleList] = useState<ParsedDay[]>(() => generateInitialSampleData());
  const [currentYear, setCurrentYear] = useState<number>(2026);
  const [currentMonth, setCurrentMonth] = useState<number>(9); // 1-indexed
  const [viewMode, setViewMode] = useState<ViewMode>('1month');
  const [onlyMyShifts, setOnlyMyShifts] = useState(false);
  const [selectedDay, setSelectedDay] = useState<ParsedDay | null>(null);
  const [isUploadOpen, setIsUploadOpen] = useState(false);
  const [copied, setCopied] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // 날짜 맵 생성 (dateStr -> ParsedDay)
  const scheduleMap = useMemo(() => {
    const map = new Map<string, ParsedDay>();
    scheduleList.forEach(day => map.set(day.date, day));
    return map;
  }, [scheduleList]);

  // 이번 달 '박현우' 근무 통계
  const myStatsThisMonth = useMemo(() => {
    const ymPrefix = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
    let total = 0;
    let dCount = 0;
    let mCount = 0;
    let nCount = 0;

    scheduleList.forEach(item => {
      if (item.date.startsWith(ymPrefix) && item.hasTargetUser) {
        total++;
        item.shifts.forEach(s => {
          if (s.hasTargetUser) {
            if (s.code === 'D') dCount++;
            else if (s.code === 'N') nCount++;
            else mCount++;
          }
        });
      }
    });

    return { total, dCount, mCount, nCount };
  }, [scheduleList, currentYear, currentMonth]);

  // 이전/다음 월 이동
  const handlePrevMonth = () => {
    if (viewMode === '3months') {
      const prev = new Date(currentYear, currentMonth - 1 - 3, 1);
      setCurrentYear(prev.getFullYear());
      setCurrentMonth(prev.getMonth() + 1);
    } else if (viewMode === '6months') {
      const prev = new Date(currentYear, currentMonth - 1 - 6, 1);
      setCurrentYear(prev.getFullYear());
      setCurrentMonth(prev.getMonth() + 1);
    } else if (viewMode === '1year') {
      setCurrentYear(prev => prev - 1);
    } else {
      const prev = new Date(currentYear, currentMonth - 1 - 1, 1);
      setCurrentYear(prev.getFullYear());
      setCurrentMonth(prev.getMonth() + 1);
    }
  };

  const handleNextMonth = () => {
    if (viewMode === '3months') {
      const next = new Date(currentYear, currentMonth - 1 + 3, 1);
      setCurrentYear(next.getFullYear());
      setCurrentMonth(next.getMonth() + 1);
    } else if (viewMode === '6months') {
      const next = new Date(currentYear, currentMonth - 1 + 6, 1);
      setCurrentYear(next.getFullYear());
      setCurrentMonth(next.getMonth() + 1);
    } else if (viewMode === '1year') {
      setCurrentYear(next => next + 1);
    } else {
      const next = new Date(currentYear, currentMonth - 1 + 1, 1);
      setCurrentYear(next.getFullYear());
      setCurrentMonth(next.getMonth() + 1);
    }
  };

  // 파일 업로드 핸들러
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseExcelData(buffer);
      if (parsed.length > 0) {
        setScheduleList(parsed);
        // 첫 번째 감지된 날짜의 연/월로 이동
        const [y, m] = parsed[0].yearMonth.split('-').map(Number);
        setCurrentYear(y);
        setCurrentMonth(m);
        setIsUploadOpen(false);
      } else {
        alert('엑셀 파일에서 근무표 데이터를 찾을 수 없습니다. YYYY.MM월 형식 및 행 구조를 확인해주세요.');
      }
    } catch (err) {
      console.error(err);
      alert('엑셀 파싱 중 오류가 발생했습니다.');
    }
  };

  // 상세 일정 클립보드 복사
  const handleCopySchedule = () => {
    if (!selectedDay) return;
    const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
    let text = `[ER Schedule] ${selectedDay.date} (${dayNames[selectedDay.dayOfWeek]}${selectedDay.holidayNote ? ` ${selectedDay.holidayNote}` : ''})\n`;
    selectedDay.shifts.forEach(s => {
      text += `• ${s.code} (${s.name}, ${s.time}): ${s.workers.join(', ')}\n`;
    });
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  // 렌더링할 월 목록 계산 (viewMode에 따라 1개, 3개, 6개, 12개)
  const monthsToRender = useMemo(() => {
    const count = viewMode === '1month' ? 1 : viewMode === '3months' ? 3 : viewMode === '6months' ? 6 : 12;
    const arr: { year: number; month: number }[] = [];
    for (let i = 0; i < count; i++) {
      const d = new Date(currentYear, currentMonth - 1 + i, 1);
      arr.push({ year: d.getFullYear(), month: d.getMonth() + 1 });
    }
    return arr;
  }, [currentYear, currentMonth, viewMode]);

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-50 flex flex-col w-full max-w-md mx-auto shadow-2xl relative select-none">
      {/* ========================================================= */}
      {/* 1. 상단 심플 헤더: [🏥 ER Schedule] & [엑셀 업로드] 버튼 */}
      {/* ========================================================= */}
      <header className="sticky top-0 z-30 pt-safe bg-zinc-950/90 backdrop-blur-md border-b border-zinc-800/80 px-4 py-3 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <span className="text-lg">🏥</span>
          <h1 className="text-base font-extrabold tracking-tight text-white flex items-center gap-1.5">
            <span>ER Schedule</span>
            <span className="w-1.5 h-1.5 rounded-full bg-yellow-400 animate-pulse" />
          </h1>
        </div>

        <button
          onClick={() => setIsUploadOpen(true)}
          className="flex items-center gap-1.5 bg-zinc-900 hover:bg-zinc-800 active:scale-95 text-zinc-100 text-xs font-semibold px-3 py-1.5 rounded-lg border border-zinc-700 transition-all shadow-sm"
        >
          <Upload className="w-3.5 h-3.5 text-yellow-400" />
          <span>엑셀 업로드</span>
        </button>
      </header>

      {/* ========================================================= */}
      {/* 2. 달력 컨트롤: 연/월 이동, 뷰어 필터(1M, 3M, 6M, 1Y) */}
      {/* ========================================================= */}
      <div className="px-4 pt-3 pb-2 space-y-2 border-b border-zinc-900 bg-zinc-950">
        <div className="flex items-center justify-between">
          {/* 이전/다음 달 이동 화살표 */}
          <div className="flex items-center gap-1.5">
            <button
              onClick={handlePrevMonth}
              aria-label="이전 달"
              className="p-1.5 rounded-lg bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border border-zinc-800 transition-colors"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>

            <span className="text-sm font-black text-white tracking-tight px-1">
              {currentYear}년 {currentMonth}월
            </span>

            <button
              onClick={handleNextMonth}
              aria-label="다음 달"
              className="p-1.5 rounded-lg bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border border-zinc-800 transition-colors"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>

          {/* 타임프레임 필터 (1개월, 3개월, 6개월, 1년) */}
          <div className="flex items-center bg-zinc-900 p-0.5 rounded-lg border border-zinc-800">
            {(['1month', '3months', '6months', '1year'] as ViewMode[]).map(mode => {
              const label = mode === '1month' ? '1M' : mode === '3months' ? '3M' : mode === '6months' ? '6M' : '1Y';
              const isActive = viewMode === mode;
              return (
                <button
                  key={mode}
                  onClick={() => setViewMode(mode)}
                  className={`text-[11px] font-bold px-2 py-1 rounded-md transition-all ${
                    isActive
                      ? 'bg-zinc-800 text-yellow-400 shadow-xs'
                      : 'text-zinc-400 hover:text-zinc-200'
                  }`}
                >
                  {label}
                </button>
              );
            })}
          </div>
        </div>

        {/* 내 근무 요약 바 및 [박현우만 보기] 토글 */}
        <div className="flex items-center justify-between text-2xs pt-1 text-zinc-400">
          <div className="flex items-center gap-1.5">
            <span className="text-yellow-400 font-extrabold">{TARGET_USER_NAME} 이번 달:</span>
            <span>총 {myStatsThisMonth.total}회</span>
            <span className="text-zinc-600">|</span>
            <span className="text-sky-400">D {myStatsThisMonth.dCount}</span>
            <span className="text-emerald-400">M {myStatsThisMonth.mCount}</span>
            <span className="text-indigo-400">N {myStatsThisMonth.nCount}</span>
          </div>

          <button
            onClick={() => setOnlyMyShifts(!onlyMyShifts)}
            className={`flex items-center gap-1 px-2 py-0.5 rounded-md font-bold text-2xs transition-all border ${
              onlyMyShifts
                ? 'bg-yellow-400 text-black border-yellow-400 shadow-xs'
                : 'bg-zinc-900 text-zinc-300 border-zinc-800 hover:bg-zinc-800'
            }`}
          >
            <Sparkles className="w-3 h-3" />
            <span>박현우만</span>
          </button>
        </div>
      </div>

      {/* ========================================================= */}
      {/* 3. 달력 그리드 영역 (1개월 / 다중 월 연속 렌더링) */}
      {/* ========================================================= */}
      <main className="flex-1 p-3 space-y-6 pb-safe">
        {monthsToRender.map(({ year, month }) => {
          const ymStr = `${year}-${String(month).padStart(2, '0')}`;
          const firstDayOfWeek = new Date(year, month - 1, 1).getDay();
          const daysInMonth = new Date(year, month, 0).getDate();

          return (
            <div key={ymStr} className="space-y-1.5">
              {/* 다중 월 모드일 때 월 헤더 표시 */}
              {viewMode !== '1month' && (
                <div className="flex items-center justify-between px-1 pt-2">
                  <span className="text-xs font-black text-zinc-200">
                    {year}년 {month}월
                  </span>
                  <span className="text-2xs text-zinc-500 font-medium">
                    {daysInMonth}일
                  </span>
                </div>
              )}

              {/* 요일 헤더 (일 ~ 토) */}
              <div className="grid grid-cols-7 gap-1 text-center text-2xs font-extrabold text-zinc-400 pb-1">
                <span className="text-rose-500">일</span>
                <span>월</span>
                <span>화</span>
                <span>수</span>
                <span>목</span>
                <span>금</span>
                <span className="text-sky-400">토</span>
              </div>

              {/* 날짜 그리드 */}
              <div className="grid grid-cols-7 gap-1">
                {/* 1일 이전 빈칸 */}
                {Array.from({ length: firstDayOfWeek }).map((_, idx) => (
                  <div key={`empty-${idx}`} className="min-h-[82px] bg-zinc-950/40 rounded-lg border border-zinc-900/60" />
                ))}

                {/* 해당 월 날짜들 */}
                {Array.from({ length: daysInMonth }).map((_, idx) => {
                  const d = idx + 1;
                  const dateStr = `${ymStr}-${String(d).padStart(2, '0')}`;
                  const dayData = scheduleMap.get(dateStr);
                  const dateObj = new Date(year, month - 1, d);
                  const dayOfWeek = dateObj.getDay();
                  const isSun = dayOfWeek === 0;
                  const isSat = dayOfWeek === 6;
                  const isHol = Boolean(dayData?.holidayNote);

                  const hasTargetUser = Boolean(dayData?.hasTargetUser);
                  const isDimmed = onlyMyShifts && !hasTargetUser;

                  return (
                    <div
                      key={dateStr}
                      onClick={() => dayData && setSelectedDay(dayData)}
                      className={`min-h-[84px] p-1 rounded-lg border flex flex-col justify-between transition-all cursor-pointer overflow-hidden ${
                        isDimmed
                          ? 'opacity-25 bg-zinc-950 border-zinc-900'
                          : hasTargetUser
                          ? 'bg-zinc-900/95 border-yellow-400/90 ring-1 ring-yellow-400/50 shadow-md target-highlight-box'
                          : 'bg-zinc-900 border-zinc-800 hover:border-zinc-700'
                      }`}
                    >
                      {/* 상단 날짜 숫자 및 공휴일 메모 */}
                      <div className="flex items-center justify-between gap-0.5 mb-1 leading-none">
                        <span
                          className={`text-xs font-black ${
                            isSun || isHol ? 'text-rose-500' : isSat ? 'text-sky-400' : 'text-zinc-200'
                          }`}
                        >
                          {d}
                        </span>

                        {dayData?.holidayNote && (
                          <span className="text-[8px] font-bold text-rose-400 bg-rose-950/60 border border-rose-900/80 px-1 py-0.2 rounded-xs truncate max-w-[42px]">
                            {dayData.holidayNote}
                          </span>
                        )}
                      </div>

                      {/* [스케줄 표시 최적화]: 약어(D, M1, M2, N, H) 사용 및 한 줄 렌더링 (글자 잘림 방지) */}
                      <div className="flex-1 flex flex-col gap-0.5 overflow-hidden w-full">
                        {!dayData || dayData.shifts.length === 0 ? (
                          <span className="text-[9px] text-zinc-600 text-center py-2">-</span>
                        ) : (
                          dayData.shifts.map((shift, sIdx) => {
                            // 근무조 약어 배지 색상
                            let codeColor = 'text-zinc-400';
                            if (shift.code === 'D') codeColor = 'text-sky-400';
                            else if (shift.code === 'M1' || shift.code === 'M2' || shift.code === 'M') codeColor = 'text-emerald-400';
                            else if (shift.code === 'H') codeColor = 'text-cyan-400';
                            else if (shift.code === 'N') codeColor = 'text-indigo-400';

                            return (
                              <div
                                key={sIdx}
                                className="text-[10px] whitespace-nowrap overflow-hidden text-ellipsis flex items-center leading-tight min-w-0"
                              >
                                <span className={`font-black mr-1 text-[9px] shrink-0 ${codeColor}`}>
                                  {shift.code}
                                </span>

                                <span className="truncate">
                                  {shift.workers.map((worker, wIdx) => {
                                    const isTarget = worker === TARGET_USER_NAME;
                                    return (
                                      <React.Fragment key={wIdx}>
                                        {/* [박현우 하이라이트]: bg-yellow-400 text-black font-extrabold px-1 py-0.5 rounded shadow */}
                                        {isTarget ? (
                                          <span className="bg-yellow-400 text-black font-extrabold px-1 py-0.5 rounded shadow inline-block tracking-tight mx-0.5">
                                            {worker}
                                          </span>
                                        ) : (
                                          <span className="text-zinc-300">{worker}</span>
                                        )}
                                        {wIdx < shift.workers.length - 1 && (
                                          <span className="text-zinc-600 mx-0.5">/</span>
                                        )}
                                      </React.Fragment>
                                    );
                                  })}
                                </span>
                              </div>
                            );
                          })
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </main>

      {/* ========================================================= */}
      {/* 4. 상세 모달 (Dialog) - 날짜 클릭 시 다크 테마 팝업 */}
      {/* ========================================================= */}
      {selectedDay && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="absolute inset-0" onClick={() => setSelectedDay(null)} />

          <div className="relative w-full max-w-sm bg-zinc-900 border border-zinc-800 rounded-2xl shadow-2xl p-5 z-10 animate-in zoom-in-95 duration-150 text-zinc-100">
            {/* 상단 닫기 */}
            <div className="flex items-center justify-between pb-3 border-b border-zinc-800">
              <div>
                <div className="flex items-center gap-1.5">
                  <span className="text-base font-black text-white">
                    {selectedDay.date}
                  </span>
                  {selectedDay.holidayNote && (
                    <span className="text-xs font-bold text-rose-400 bg-rose-950/80 px-2 py-0.5 rounded-full border border-rose-900">
                      {selectedDay.holidayNote}
                    </span>
                  )}
                </div>
                <span className="text-2xs text-zinc-400">
                  {selectedDay.shifts.length}개 근무조 편성
                </span>
              </div>

              <button
                onClick={() => setSelectedDay(null)}
                className="p-1 rounded-lg text-zinc-400 hover:text-white bg-zinc-800 hover:bg-zinc-700 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* 박현우 당직 안내 배너 */}
            {selectedDay.hasTargetUser && (
              <div className="mt-3 p-2.5 bg-yellow-400/10 border border-yellow-400/40 rounded-xl flex items-center gap-2 text-yellow-400">
                <Sparkles className="w-4 h-4 fill-yellow-400 shrink-0" />
                <span className="text-xs font-extrabold">
                  {TARGET_USER_NAME} 당직 근무일입니다.
                </span>
              </div>
            )}

            {/* 근무조 상세 리스트 */}
            <div className="mt-3.5 space-y-2">
              {selectedDay.shifts.map((shift, idx) => {
                const isTarget = shift.hasTargetUser;
                return (
                  <div
                    key={idx}
                    className={`p-3 rounded-xl border text-xs ${
                      isTarget
                        ? 'bg-zinc-800/90 border-yellow-400/60 ring-1 ring-yellow-400/30'
                        : 'bg-zinc-950 border-zinc-800'
                    }`}
                  >
                    <div className="flex items-center justify-between mb-1.5">
                      <div className="flex items-center gap-1.5 font-bold">
                        <span className="bg-zinc-800 text-yellow-400 px-1.5 py-0.2 rounded text-2xs font-black">
                          {shift.code}
                        </span>
                        <span className="text-zinc-200">{shift.name}</span>
                      </div>
                      <span className="text-2xs text-zinc-400">{shift.time}</span>
                    </div>

                    <div className="flex flex-wrap items-center gap-1.5 pt-1 border-t border-zinc-800/60">
                      {shift.workers.map((worker, wIdx) => (
                        <span
                          key={wIdx}
                          className={`px-2 py-0.5 rounded text-xs font-semibold ${
                            worker === TARGET_USER_NAME
                              ? 'bg-yellow-400 text-black font-extrabold shadow-sm'
                              : 'bg-zinc-800 text-zinc-300'
                          }`}
                        >
                          {worker}
                        </span>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* 하단 공유/닫기 액션 */}
            <div className="mt-4 pt-3 border-t border-zinc-800 flex gap-2">
              <button
                onClick={handleCopySchedule}
                className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl bg-yellow-400 hover:bg-yellow-300 active:scale-98 text-black font-extrabold text-xs transition-all shadow-md"
              >
                {copied ? <Check className="w-3.5 h-3.5" /> : <Share2 className="w-3.5 h-3.5" />}
                <span>{copied ? '복사 완료!' : '일정 텍스트 복사'}</span>
              </button>
              <button
                onClick={() => setSelectedDay(null)}
                className="px-4 py-2.5 rounded-xl bg-zinc-800 hover:bg-zinc-700 text-zinc-300 text-xs font-bold transition-colors"
              >
                닫기
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================= */}
      {/* 5. 엑셀 파일 업로드 모달 */}
      {/* ========================================================= */}
      {isUploadOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="absolute inset-0" onClick={() => setIsUploadOpen(false)} />

          <div className="relative w-full max-w-sm bg-zinc-900 border border-zinc-800 rounded-2xl shadow-2xl p-5 z-10 text-zinc-100">
            <div className="flex items-center justify-between pb-3 border-b border-zinc-800">
              <div className="flex items-center gap-2">
                <CalendarIcon className="w-4 h-4 text-yellow-400" />
                <h3 className="text-sm font-extrabold text-white">근무표 엑셀 업로드</h3>
              </div>
              <button
                onClick={() => setIsUploadOpen(false)}
                className="p-1 rounded-lg text-zinc-400 hover:text-white bg-zinc-800 hover:bg-zinc-700 transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div
              onClick={() => fileInputRef.current?.click()}
              className="mt-4 border-2 border-dashed border-zinc-700 hover:border-yellow-400 bg-zinc-950/60 rounded-xl p-6 text-center cursor-pointer transition-colors"
            >
              <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileUpload}
                accept=".xlsx,.xls"
                className="hidden"
              />
              <Upload className="w-8 h-8 text-yellow-400 mx-auto mb-2 opacity-80" />
              <p className="text-xs font-bold text-zinc-200">
                .xlsx 파일을 선택하세요
              </p>
              <p className="text-2xs text-zinc-500 mt-1">
                9월, 10월, 11월 등 여러 달의 데이터가 한 시트에 있어도 자동 파싱됩니다.
              </p>
            </div>

            <div className="mt-3.5 p-2.5 bg-zinc-950 rounded-xl border border-zinc-800 text-2xs space-y-1 text-zinc-400">
              <p className="text-zinc-300 font-bold">파싱 알고리즘 안내:</p>
              <p>• 셀 내 `YYYY.MM월` 패턴으로 월을 자동 식별합니다.</p>
              <p>• 날짜 셀(`25(추석)`)에서 숫자만 정확히 추출합니다.</p>
              <p>• 줄바꿈(\n)에 따라 D / M1 / M2 / N / H로 자동 변환됩니다.</p>
            </div>

            <button
              onClick={() => {
                setScheduleList(generateInitialSampleData());
                setCurrentYear(2026);
                setCurrentMonth(9);
                setIsUploadOpen(false);
              }}
              className="w-full mt-3.5 py-2.5 rounded-xl bg-zinc-800 hover:bg-zinc-700 text-zinc-200 text-xs font-bold transition-colors"
            >
              2026년 9~11월 샘플 데이터로 리셋
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
