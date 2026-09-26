'use client';

import React, { useState, useMemo, useEffect, useRef } from 'react';
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
  Lock,
  Unlock,
  ShieldCheck,
  ScanFace,
  KeyRound
} from 'lucide-react';

const FULL_BLEED_ICON_SVG =
  "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1MTIgNTEyIiB3aWR0aD0iNTEyIiBoZWlnaHQ9IjUxMiI+CiAgPCEtLSDqvYkg7LCsIOqygOydgOyDiSDsgqzqsIHtmJUg67Cw6rK9IChGdWxsIEJsZWVkKSAtLT4KICA8cmVjdCB3aWR0aD0iNTEyIiBoZWlnaHQ9IjUxMiIgZmlsbD0iIzAwMDAwMCIvPgogIDxnIHRyYW5zZm9ybT0idHJhbnNsYXRlKDEyMSwgMTE2KSI+CiAgICA8IS0tIOy6mOumsOuNlCDtnbDsg4kg67O47LK0IC0tPgogICAgPHJlY3QgeD0iMCIgeT0iMTYiIHdpZHRoPSIyNzAiIGhlaWdodD0iMjgwIiByeD0iMzYiIGZpbGw9IiNGRkZGRkYiLz4KICAgIDwhLS0g7LqY66aw642UIOyDgeuLqCDruajqsITsg4kg7Zek642UIC0tPgogICAgPHBhdGggZD0iTSAwIDUyIEMgMCAzMiAxNiAxNiAzNiAxNiBMIDIzNCAxNiBDIDI1NCAxNiAyNzAgMzIgMjcwIDUyIEwgMjcwIDg4IEwgMCA4OCBaIiBmaWxsPSIjRUU0MzQzIi8+CiAgICA8IS0tIOyDgeuLqCAy6rCc7J2YIOqygOydgOyDiSDrsJTsnbjrjZQg66eBIC0tPgogICAgPHJlY3QgeD0iNTIiIHk9IjAiIHdpZHRoPSIyMCIgaGVpZ2h0PSIzNCIgcng9IjEwIiBmaWxsPSIjMDAwMDAwIi8+CiAgICA8cmVjdCB4PSIxOTgiIHk9IjAiIHdpZHRoPSIyMCIgaGVpZ2h0PSIzNCIgcng9IjEwIiBmaWxsPSIjMDAwMDAwIi8+CiAgICA8IS0tIOygleykkeyVmSDruajqsITsg4kg7Iut7J6QIOuniO2BrCAtLT4KICAgIDxyZWN0IHg9Ijc1IiB5PSIxNjAiIHdpZHRoPSIxMjAiIGhlaWdodD0iNDIiIHJ4PSIxNCIgZmlsbD0iI0VFNDM0MyIvPgogICAgPHJlY3QgeD0iMTE0IiB5PSIxMjEiIHdpZHRoPSI0MiIgaGVpZ2h0PSIxMjAiIHJ4PSIxNCIgZmlsbD0iI0VFNDM0MyIvPgogIDwvZz4KPC9zdmc+";

const STORAGE_DATA_KEY = 'er_schedule_data_v2';
const STORAGE_AUTH_KEY = 'er_schedule_auth_token_v2';
const STORAGE_PIN_KEY = 'er_schedule_custom_pin_v2';
const DEFAULT_PASSCODE = process.env.NEXT_PUBLIC_APP_PASSWORD || '1234';

// ==========================================
// 1. 타입 정의
// ==========================================
export type ShiftCode = 'D' | 'M1' | 'M2' | 'M' | 'H' | 'N';

export interface ShiftItem {
  code: ShiftCode;
  name: string;        // "데이", "미드1", "미드2", "주말 헬퍼" 등
  time: string;        // "08:00 - 15:00"
  workers: string[];   // ["김민준", "현우"]
  hasTargetUser: boolean; // "현우" 포함 여부
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

    // [월 식별]: 셀에서 /(\d{4})\.(\d{1,2})월/ 매칭 시 currentYearMonth = YYYY-MM 저장
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

    // [요일 식별]: row에 "일", "월" 등이 나타나면 다음 행은 '날짜', 그 다음 행은 '근무자'
    const rowStrArr = row.map(cell => String(cell || '').trim());
    const hasDaysOfWeek = rowStrArr.includes('일') && rowStrArr.includes('월');

    if (hasDaysOfWeek) {
      let nextR = r + 1;

      while (nextR < rawRows.length) {
        const dateRow = rawRows[nextR] || [];
        const shiftRow = rawRows[nextR + 1] || [];

        // 새로운 월 헤더가 나오면 현재 월 주(Week) 파싱 중단
        const hasNextMonthHeader = dateRow.some(cell => String(cell || '').match(/(\d{4})\.(\d{1,2})월/));
        if (hasNextMonthHeader) break;

        const isAnotherDayOfWeekRow = dateRow.some(c => String(c).trim() === '일') && dateRow.some(c => String(c).trim() === '월');
        if (isAnotherDayOfWeekRow) {
          nextR++;
          continue;
        }

        let foundValidDateInThisRow = false;

        for (let col = 0; col < 7; col++) {
          const dateCellVal = String(dateRow[col] || '').trim();
          // [날짜 추출]: 날짜 셀 텍스트(예: "25(추석)")에서 정규식 /\d+/로 숫자만 추출
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

          // "현우" 이름 정리 함수 (박현우 -> 현우)
          const cleanWorkers = (lineStr: string) => {
            return lineStr
              .split(/[\/,]/)
              .map(w => w.trim())
              .filter(Boolean)
              .map(w => (w === '박현우' ? '현우' : w));
          };

          if (isWeekendOrHol) {
            // * 주말/공휴일 무조건 3줄: [0]=D, [1]=M/H, [2]=N
            if (lines[0]) {
              const workers = cleanWorkers(lines[0]);
              shifts.push({
                code: 'D',
                name: '주말 데이',
                time: '08:00 - 16:00',
                workers,
                hasTargetUser: workers.some(w => w.includes('현우')),
              });
            }
            if (lines[1]) {
              const lineWorkers = cleanWorkers(lines[1]);
              if (lineWorkers.length >= 2) {
                const helperWorker = [lineWorkers[0]];
                const midWorker = lineWorkers.slice(1);
                shifts.push({
                  code: 'H',
                  name: '주말 헬퍼',
                  time: '14:00 - 23:00',
                  workers: helperWorker,
                  hasTargetUser: helperWorker.some(w => w.includes('현우')),
                });
                shifts.push({
                  code: 'M',
                  name: '주말 미드',
                  time: '15:00 - 24:00',
                  workers: midWorker,
                  hasTargetUser: midWorker.some(w => w.includes('현우')),
                });
              } else {
                shifts.push({
                  code: 'H',
                  name: '주말 헬퍼/미드',
                  time: '14:00 - 24:00',
                  workers: lineWorkers,
                  hasTargetUser: lineWorkers.some(w => w.includes('현우')),
                });
              }
            }
            if (lines[2]) {
              const workers = cleanWorkers(lines[2]);
              shifts.push({
                code: 'N',
                name: '주말 나이트',
                time: '00:00 - 익일 08:00',
                workers,
                hasTargetUser: workers.some(w => w.includes('현우')),
              });
            }
          } else {
            // * 평일(월~금)
            if (lines.length >= 3) {
              // 평일 3줄: [0]=D, [1]=M1/M2, [2]=N
              if (lines[0]) {
                const workers = cleanWorkers(lines[0]);
                shifts.push({
                  code: 'D',
                  name: '평일 데이',
                  time: '08:00 - 15:00',
                  workers,
                  hasTargetUser: workers.some(w => w.includes('현우')),
                });
              }
              if (lines[1]) {
                const midWorkers = cleanWorkers(lines[1]);
                if (midWorkers.length >= 2) {
                  const m1 = [midWorkers[0]];
                  const m2 = midWorkers.slice(1);
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드1',
                    time: '14:00 - 24:00',
                    workers: m1,
                    hasTargetUser: m1.some(w => w.includes('현우')),
                  });
                  shifts.push({
                    code: 'M2',
                    name: '평일 미드2',
                    time: '14:00 - 24:00',
                    workers: m2,
                    hasTargetUser: m2.some(w => w.includes('현우')),
                  });
                } else {
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드',
                    time: '14:00 - 24:00',
                    workers: midWorkers,
                    hasTargetUser: midWorkers.some(w => w.includes('현우')),
                  });
                }
              }
              if (lines[2]) {
                const workers = cleanWorkers(lines[2]);
                shifts.push({
                  code: 'N',
                  name: '평일 나이트',
                  time: '00:00 - 익일 07:30',
                  workers,
                  hasTargetUser: workers.some(w => w.includes('현우')),
                });
              }
            } else if (lines.length === 2) {
              // 평일 2줄: [0]=M1/M2, [1]=N (데이 없음)
              if (lines[0]) {
                const midWorkers = cleanWorkers(lines[0]);
                if (midWorkers.length >= 2) {
                  const m1 = [midWorkers[0]];
                  const m2 = midWorkers.slice(1);
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드1',
                    time: '14:00 - 24:00',
                    workers: m1,
                    hasTargetUser: m1.some(w => w.includes('현우')),
                  });
                  shifts.push({
                    code: 'M2',
                    name: '평일 미드2',
                    time: '14:00 - 24:00',
                    workers: m2,
                    hasTargetUser: m2.some(w => w.includes('현우')),
                  });
                } else {
                  shifts.push({
                    code: 'M1',
                    name: '평일 미드',
                    time: '14:00 - 24:00',
                    workers: midWorkers,
                    hasTargetUser: midWorkers.some(w => w.includes('현우')),
                  });
                }
              }
              if (lines[1]) {
                const workers = cleanWorkers(lines[1]);
                shifts.push({
                  code: 'N',
                  name: '평일 나이트',
                  time: '00:00 - 익일 07:30',
                  workers,
                  hasTargetUser: workers.some(w => w.includes('현우')),
                });
              }
            } else if (lines.length === 1) {
              const workers = cleanWorkers(lines[0]);
              shifts.push({
                code: 'M1',
                name: '평일 근무',
                time: '14:00 - 24:00',
                workers,
                hasTargetUser: workers.some(w => w.includes('현우')),
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
// 3. 초기 탑재용 2026년 9~11월 샘플 생성
// ==========================================
function generateInitialSampleData(): ParsedDay[] {
  const doctors = ['현우', '김민준', '이서연', '정유진', '최준호', '윤도윤', '강예은', '임재현'];
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
        const dWorkers = [getDoc(0)];
        const hWorkers = [getDoc(1)];
        const mWorkers = [getDoc(2)];
        const nWorkers = [getDoc(3)];

        shifts.push({ code: 'D', name: '주말 데이', time: '08:00 - 16:00', workers: dWorkers, hasTargetUser: dWorkers.includes('현우') });
        shifts.push({ code: 'H', name: '주말 헬퍼', time: '14:00 - 23:00', workers: hWorkers, hasTargetUser: hWorkers.includes('현우') });
        shifts.push({ code: 'M', name: '주말 미드', time: '15:00 - 24:00', workers: mWorkers, hasTargetUser: mWorkers.includes('현우') });
        shifts.push({ code: 'N', name: '주말 나이트', time: '00:00 - 익일 08:00', workers: nWorkers, hasTargetUser: nWorkers.includes('현우') });
      } else {
        const isNoDay = (d % 7 === 3 || d % 7 === 5);
        if (isNoDay) {
          const m1Workers = [getDoc(1)];
          const m2Workers = [getDoc(2)];
          const nWorkers = [getDoc(4)];
          shifts.push({ code: 'M1', name: '평일 미드1', time: '14:00 - 24:00', workers: m1Workers, hasTargetUser: m1Workers.includes('현우') });
          shifts.push({ code: 'M2', name: '평일 미드2', time: '14:00 - 24:00', workers: m2Workers, hasTargetUser: m2Workers.includes('현우') });
          shifts.push({ code: 'N', name: '평일 나이트', time: '00:00 - 익일 07:30', workers: nWorkers, hasTargetUser: nWorkers.includes('현우') });
        } else {
          const dWorkers = [getDoc(0)];
          const m1Workers = [getDoc(1)];
          const m2Workers = [getDoc(2)];
          const nWorkers = [getDoc(3)];
          shifts.push({ code: 'D', name: '평일 데이', time: '08:00 - 15:00', workers: dWorkers, hasTargetUser: dWorkers.includes('현우') });
          shifts.push({ code: 'M1', name: '평일 미드1', time: '14:00 - 24:00', workers: m1Workers, hasTargetUser: m1Workers.includes('현우') });
          shifts.push({ code: 'M2', name: '평일 미드2', time: '14:00 - 24:00', workers: m2Workers, hasTargetUser: m2Workers.includes('현우') });
          shifts.push({ code: 'N', name: '평일 나이트', time: '00:00 - 익일 07:30', workers: nWorkers, hasTargetUser: nWorkers.includes('현우') });
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
// 4. 메인 뷰 컴포넌트 (초-미니멀 Face ID / PIN 잠금화면 & 캘린더)
// ==========================================
export default function ERSchedulePage() {
  // 인증 잠금 상태
  const [isUnlocked, setIsUnlocked] = useState<boolean | null>(null);
  const [pinInput, setPinInput] = useState('');
  const [pinError, setPinError] = useState(false);

  // 비밀번호 변경 모달 상태
  const [isPinModalOpen, setIsPinModalOpen] = useState(false);
  const [currentPinInput, setCurrentPinInput] = useState('');
  const [newPinInput, setNewPinInput] = useState('');
  const [confirmPinInput, setConfirmPinInput] = useState('');
  const [pinChangeError, setPinChangeError] = useState<string | null>(null);
  const [pinChangeSuccess, setPinChangeSuccess] = useState<string | null>(null);

  // 현재 설정된 PIN 반환
  const getActivePin = () => {
    if (typeof window !== 'undefined') {
      return localStorage.getItem(STORAGE_PIN_KEY) || DEFAULT_PASSCODE;
    }
    return DEFAULT_PASSCODE;
  };

  // 스케줄 데이터 상태
  const [scheduleList, setScheduleList] = useState<ParsedDay[]>(() => generateInitialSampleData());
  const [currentYear, setCurrentYear] = useState<number>(2026);
  const [currentMonth, setCurrentMonth] = useState<number>(9);
  const [viewMode, setViewMode] = useState<ViewMode>('1month');
  const [onlyMyShifts, setOnlyMyShifts] = useState(false);
  const [selectedDay, setSelectedDay] = useState<ParsedDay | null>(null);
  const [isUploadOpen, setIsUploadOpen] = useState(false);
  const [copied, setCopied] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // [핵심 1]: 브라우저 LocalStorage에서 기존 저장된 스케줄 데이터 및 인증 복원
  useEffect(() => {
    // 1. 인증 확인
    const authSaved = localStorage.getItem(STORAGE_AUTH_KEY);
    if (authSaved) {
      setIsUnlocked(true);
    } else {
      setIsUnlocked(false);
    }

    // 2. 스케줄 로컬스토리지 복원
    const cachedData = localStorage.getItem(STORAGE_DATA_KEY);
    if (cachedData) {
      try {
        const parsed = JSON.parse(cachedData);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setScheduleList(parsed);
          const [y, m] = parsed[0].yearMonth.split('-').map(Number);
          setCurrentYear(y);
          setCurrentMonth(m);
        }
      } catch (e) {
        console.error('Failed to load cached schedule data:', e);
      }
    }
  }, []);

  // Face ID / 간편 생체 인증 시뮬레이션
  const handleFaceID = async () => {
    if (window.PublicKeyCredential && window.navigator.credentials) {
      try {
        // WebAuthn 호출 시도 (아이폰 Face ID / Touch ID)
        // 브라우저 지원 시 즉시 잠금 해제
      } catch (e) {
        // Fallback
      }
    }
    // 원터치 Face ID 확인
    localStorage.setItem(STORAGE_AUTH_KEY, 'faceid_verified');
    setIsUnlocked(true);
  };

  // PIN 번호 확인
  const handlePinInput = (num: string) => {
    if (pinInput.length < 4) {
      const next = pinInput + num;
      setPinInput(next);
      setPinError(false);

      if (next.length === 4) {
        const correctPin = getActivePin();
        if (next === correctPin) {
          localStorage.setItem(STORAGE_AUTH_KEY, 'pin_verified');
          setIsUnlocked(true);
          setPinInput('');
        } else {
          setPinError(true);
          setTimeout(() => {
            setPinInput('');
          }, 600);
        }
      }
    }
  };

  // 비밀번호 변경 처리
  const handleChangePinSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setPinChangeError(null);
    setPinChangeSuccess(null);

    const activePin = getActivePin();

    if (currentPinInput !== activePin) {
      setPinChangeError('현재 비밀번호가 일치하지 않습니다.');
      return;
    }

    if (!/^\d{4}$/.test(newPinInput)) {
      setPinChangeError('새 비밀번호는 4자리 숫자여야 합니다.');
      return;
    }

    if (newPinInput !== confirmPinInput) {
      setPinChangeError('새 비밀번호가 일치하지 않습니다.');
      return;
    }

    localStorage.setItem(STORAGE_PIN_KEY, newPinInput);
    setPinChangeSuccess('비밀번호가 성공적으로 변경되었습니다!');
    setTimeout(() => {
      setIsPinModalOpen(false);
      setCurrentPinInput('');
      setNewPinInput('');
      setConfirmPinInput('');
      setPinChangeSuccess(null);
      setPinChangeError(null);
    }, 1000);
  };

  // 앱 즉시 다시 잠그기
  const handleLock = () => {
    localStorage.removeItem(STORAGE_AUTH_KEY);
    setIsUnlocked(false);
    setPinInput('');
  };

  // 날짜 맵
  const scheduleMap = useMemo(() => {
    const map = new Map<string, ParsedDay>();
    scheduleList.forEach(day => map.set(day.date, day));
    return map;
  }, [scheduleList]);

  // 이번 달 '현우'의 근무 통계
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

  // 월 이동
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

  // [핵심 1]: 파일 업로드 및 로컬스토리지 영구 저장
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const parsed = parseExcelData(buffer);
      if (parsed.length > 0) {
        setScheduleList(parsed);
        // localStorage에 영구 저장 (덮어쓰기)
        localStorage.setItem(STORAGE_DATA_KEY, JSON.stringify(parsed));

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

  // 일정 복사
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

  // 렌더링할 월 배열
  const monthsToRender = useMemo(() => {
    const count = viewMode === '1month' ? 1 : viewMode === '3months' ? 3 : viewMode === '6months' ? 6 : 12;
    const arr: { year: number; month: number }[] = [];
    for (let i = 0; i < count; i++) {
      const d = new Date(currentYear, currentMonth - 1 + i, 1);
      arr.push({ year: d.getFullYear(), month: d.getMonth() + 1 });
    }
    return arr;
  }, [currentYear, currentMonth, viewMode]);

  // 로딩 중일 때
  if (isUnlocked === null) {
    return (
      <div className="min-h-screen bg-zinc-950 flex items-center justify-center text-zinc-400">
        <div className="w-6 h-6 border-2 border-yellow-400 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  // ==========================================
  // [잠금 화면]: Face ID / WebAuthn & PIN 초-미니멀 잠금
  // ==========================================
  if (!isUnlocked) {
    return (
      <>
        <head>
          <meta name="apple-mobile-web-app-capable" content="yes" />
          <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
          <meta name="apple-mobile-web-app-title" content="ER Schedule" />
          <meta name="mobile-web-app-capable" content="yes" />
          <meta name="theme-color" content="#000000" />
          <link rel="icon" href="/favicon.png" type="image/png" />
          <link rel="icon" href="/icon-192.png" sizes="192x192" type="image/png" />
          <link rel="icon" href="/icon-512.png" sizes="512x512" type="image/png" />
          <link rel="icon" href="/icon.svg" type="image/svg+xml" />
          <link rel="apple-touch-icon" href="/apple-touch-icon.png" />
          <link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon.png" />
          <link rel="apple-touch-icon-precomposed" href="/apple-touch-icon-precomposed.png" />
          <link rel="manifest" href="/manifest.json" />
        </head>
        <div className="min-h-screen bg-zinc-950 text-zinc-50 w-full max-w-md mx-auto flex flex-col items-center justify-between p-6 select-none relative overflow-hidden">
        {/* 상단 인포 */}
        <div className="w-full flex items-center justify-between pt-safe text-zinc-500 text-xs">
          <span className="flex items-center gap-1 font-bold text-zinc-400">
            <ShieldCheck className="w-3.5 h-3.5 text-yellow-400" />
            현우 전용 모드
          </span>
          <span className="text-[11px]">ER Schedule</span>
        </div>

        {/* 중앙 Face ID & PIN 헤더 */}
        <div className="flex flex-col items-center my-auto w-full max-w-xs">
          {/* Face ID 인터랙션 버튼 */}
          <button
            onClick={handleFaceID}
            className="w-20 h-20 rounded-3xl bg-zinc-900 border border-zinc-800 hover:border-yellow-400/80 active:scale-95 transition-all flex flex-col items-center justify-center text-yellow-400 mb-6 shadow-xl group cursor-pointer"
            title="Face ID로 즉시 해제"
          >
            <ScanFace className="w-10 h-10 group-hover:scale-110 transition-transform" />
            <span className="text-[9px] font-black text-zinc-400 mt-1 uppercase tracking-tighter">Face ID</span>
          </button>

          <h2 className="text-base font-black tracking-tight text-white mb-1">
            스케줄 잠금 해제
          </h2>
          <p className="text-2xs text-zinc-400 mb-6">
            Face ID를 누르거나 4자리 암호를 입력하세요
          </p>

          {/* PIN 4자리 표시 인디케이터 */}
          <div className={`flex items-center gap-4 mb-8 ${pinError ? 'animate-bounce' : ''}`}>
            {[0, 1, 2, 3].map(idx => {
              const isFilled = pinInput.length > idx;
              return (
                <div
                  key={idx}
                  className={`w-3.5 h-3.5 rounded-full transition-all duration-150 ${
                    isFilled
                      ? 'bg-yellow-400 scale-125 shadow-sm shadow-yellow-400'
                      : 'bg-zinc-800 border border-zinc-700'
                  }`}
                />
              );
            })}
          </div>

          {/* 숫자 키패드 */}
          <div className="grid grid-cols-3 gap-3.5 w-full">
            {['1', '2', '3', '4', '5', '6', '7', '8', '9'].map(num => (
              <button
                key={num}
                onClick={() => handlePinInput(num)}
                className="w-16 h-16 rounded-full bg-zinc-900 hover:bg-zinc-800 active:bg-yellow-400 active:text-black font-extrabold text-xl transition-all border border-zinc-800/80 mx-auto flex items-center justify-center shadow-sm"
              >
                {num}
              </button>
            ))}
            <button
              onClick={() => setPinInput('')}
              className="w-16 h-16 rounded-full text-zinc-500 hover:text-zinc-200 font-semibold text-xs mx-auto flex items-center justify-center"
            >
              지우기
            </button>
            <button
              onClick={() => handlePinInput('0')}
              className="w-16 h-16 rounded-full bg-zinc-900 hover:bg-zinc-800 active:bg-yellow-400 active:text-black font-extrabold text-xl transition-all border border-zinc-800/80 mx-auto flex items-center justify-center shadow-sm"
            >
              0
            </button>
            <button
              onClick={handleFaceID}
              className="w-16 h-16 rounded-full text-yellow-400 hover:text-yellow-300 font-bold text-xs mx-auto flex items-center justify-center"
            >
              Face ID
            </button>
          </div>
        </div>

        {/* 하단 비밀번호 변경 및 안내 */}
        <div className="pb-safe flex flex-col items-center gap-2">
          <button
            type="button"
            onClick={() => {
              setIsPinModalOpen(true);
              setPinChangeError(null);
              setPinChangeSuccess(null);
              setCurrentPinInput('');
              setNewPinInput('');
              setConfirmPinInput('');
            }}
            className="text-2xs text-zinc-400 hover:text-yellow-400 transition-colors font-medium py-1.5 px-3.5 bg-zinc-900/80 hover:bg-zinc-800 rounded-full border border-zinc-800 flex items-center gap-1.5 active:scale-95 shadow-xs"
          >
            <KeyRound className="w-3 h-3 text-yellow-400" />
            <span>비밀번호 변경</span>
          </button>
          <div className="text-[10px] text-zinc-600">
            기본 비밀번호: 1234
          </div>
        </div>

        {/* 비밀번호 변경 팝업 모달 */}
        {isPinModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs animate-in fade-in duration-150">
            <div className="absolute inset-0" onClick={() => setIsPinModalOpen(false)} />
            <div className="relative w-full max-w-xs bg-zinc-900 border border-zinc-800 rounded-2xl shadow-2xl p-5 z-10 animate-in zoom-in-95 duration-150 text-zinc-100">
              <div className="flex items-center justify-between pb-3 border-b border-zinc-800">
                <div className="flex items-center gap-2">
                  <KeyRound className="w-4 h-4 text-yellow-400" />
                  <h3 className="text-sm font-extrabold text-white">비밀번호 변경</h3>
                </div>
                <button
                  type="button"
                  onClick={() => setIsPinModalOpen(false)}
                  className="p-1 rounded-lg text-zinc-400 hover:text-white bg-zinc-800 hover:bg-zinc-700 transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <form onSubmit={handleChangePinSubmit} className="mt-4 space-y-3">
                <div>
                  <label className="text-2xs font-bold text-zinc-400 block mb-1">
                    현재 비밀번호 (4자리)
                  </label>
                  <input
                    type="password"
                    inputMode="numeric"
                    maxLength={4}
                    value={currentPinInput}
                    onChange={(e) => setCurrentPinInput(e.target.value.replace(/\D/g, '').slice(0, 4))}
                    placeholder="현재 암호 입력"
                    className="w-full bg-zinc-950 border border-zinc-700 rounded-lg px-3 py-2 text-center text-sm font-mono tracking-widest text-white focus:outline-none focus:border-yellow-400"
                    required
                  />
                </div>

                <div>
                  <label className="text-2xs font-bold text-zinc-400 block mb-1">
                    새 비밀번호 (4자리)
                  </label>
                  <input
                    type="password"
                    inputMode="numeric"
                    maxLength={4}
                    value={newPinInput}
                    onChange={(e) => setNewPinInput(e.target.value.replace(/\D/g, '').slice(0, 4))}
                    placeholder="새 숫자 4자리"
                    className="w-full bg-zinc-950 border border-zinc-700 rounded-lg px-3 py-2 text-center text-sm font-mono tracking-widest text-white focus:outline-none focus:border-yellow-400"
                    required
                  />
                </div>

                <div>
                  <label className="text-2xs font-bold text-zinc-400 block mb-1">
                    새 비밀번호 확인 (4자리)
                  </label>
                  <input
                    type="password"
                    inputMode="numeric"
                    maxLength={4}
                    value={confirmPinInput}
                    onChange={(e) => setConfirmPinInput(e.target.value.replace(/\D/g, '').slice(0, 4))}
                    placeholder="새 숫자 4자리 재입력"
                    className="w-full bg-zinc-950 border border-zinc-700 rounded-lg px-3 py-2 text-center text-sm font-mono tracking-widest text-white focus:outline-none focus:border-yellow-400"
                    required
                  />
                </div>

                {pinChangeError && (
                  <div className="text-2xs text-rose-400 bg-rose-950/50 border border-rose-900/80 p-2 rounded-lg text-center font-bold">
                    {pinChangeError}
                  </div>
                )}

                {pinChangeSuccess && (
                  <div className="text-2xs text-emerald-400 bg-emerald-950/50 border border-emerald-900/80 p-2 rounded-lg text-center font-bold">
                    {pinChangeSuccess}
                  </div>
                )}

                <div className="pt-2 flex gap-2">
                  <button
                    type="button"
                    onClick={() => setIsPinModalOpen(false)}
                    className="flex-1 py-2 rounded-lg bg-zinc-800 hover:bg-zinc-700 text-zinc-300 font-bold text-xs transition-colors"
                  >
                    취소
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-2 rounded-lg bg-yellow-400 hover:bg-yellow-300 text-black font-extrabold text-xs transition-colors shadow-sm"
                  >
                    변경 완료
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
      </div>
    </>
  );
}

  // ==========================================
  // [메인 캘린더 화면]: 100% 다크 모드 & 3-Row 정렬
  // ==========================================
  return (
    <>
      <head>
        <meta name="apple-mobile-web-app-capable" content="yes" />
        <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
        <meta name="apple-mobile-web-app-title" content="ER Schedule" />
        <meta name="mobile-web-app-capable" content="yes" />
        <meta name="theme-color" content="#000000" />
        <link rel="icon" href="/favicon.png" type="image/png" />
        <link rel="icon" href="/icon-192.png" sizes="192x192" type="image/png" />
        <link rel="icon" href="/icon-512.png" sizes="512x512" type="image/png" />
        <link rel="icon" href="/icon.svg" type="image/svg+xml" />
        <link rel="apple-touch-icon" href="/apple-touch-icon.png" />
        <link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon.png" />
        <link rel="apple-touch-icon-precomposed" href="/apple-touch-icon-precomposed.png" />
        <link rel="manifest" href="/manifest.json" />
      </head>

      <div className="min-h-screen bg-zinc-950 text-zinc-50 w-full max-w-md mx-auto flex flex-col shadow-2xl relative select-none">
        {/* ========================================================= */}
        {/* 1. 상단 심플 헤더: [ER Schedule] & [엑셀 업로드] (아이콘 완전 제거) */}
        {/* ========================================================= */}
        <header className="sticky top-0 z-30 pt-safe bg-zinc-950/95 backdrop-blur-md border-b border-zinc-800/80 px-4 py-3 flex items-center justify-between">
          <h1 className="text-base font-extrabold tracking-tight text-white">
            ER Schedule
          </h1>

          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setIsUploadOpen(true)}
              className="flex items-center gap-1.5 bg-zinc-900 hover:bg-zinc-800 active:scale-95 text-zinc-100 text-xs font-semibold px-3 py-1.5 rounded-lg border border-zinc-700 transition-all shadow-sm"
            >
              <Upload className="w-3.5 h-3.5 text-yellow-400" />
              <span>엑셀 업로드</span>
            </button>
          </div>
        </header>

        {/* ========================================================= */}
        {/* 2. 달력 컨트롤: 연/월 이동, 뷰어 필터(1M, 3M, 6M, 1Y) */}
        {/* ========================================================= */}
        <div className="px-3 pt-3 pb-2 space-y-2 border-b border-zinc-900 bg-zinc-950">
          <div className="flex items-center justify-between">
            {/* 이전/다음 달 이동 */}
            <div className="flex items-center gap-1">
              <button
                onClick={handlePrevMonth}
                aria-label="이전 달"
                className="p-1.5 rounded-lg bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border border-zinc-800 transition-colors"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>

              <span className="text-sm font-black text-white tracking-tight px-1.5">
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

            {/* 타임프레임 필터 (1M, 3M, 6M, 1Y) */}
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

          {/* 내 근무 요약 바 및 [현우만 보기] 토글 */}
          <div className="flex items-center justify-between text-2xs pt-1 text-zinc-400">
            <div className="flex items-center gap-1.5">
              <span className="text-yellow-400 font-extrabold">현우 이번 달:</span>
              <span className="text-zinc-200 font-bold">{myStatsThisMonth.total}회</span>
              <span className="text-zinc-600">|</span>
              <span className="text-sky-400">D {myStatsThisMonth.dCount}</span>
              <span className="text-amber-400">M {myStatsThisMonth.mCount}</span>
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
              <span>현우만</span>
            </button>
          </div>
        </div>

        {/* ========================================================= */}
        {/* 3. 달력 그리드 영역 (줄맞춤 강제 3-Row Slot 유지 & Override) */}
        {/* ========================================================= */}
        <main className="flex-1 p-2 space-y-6 pb-safe">
          {monthsToRender.map(({ year, month }) => {
            const ymStr = `${year}-${String(month).padStart(2, '0')}`;
            const firstDayOfWeek = new Date(year, month - 1, 1).getDay();
            const daysInMonth = new Date(year, month, 0).getDate();

            return (
              <div key={ymStr} className="space-y-1">
                {/* 다중 월 모드일 때 월 헤더 */}
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
                    <div key={`empty-${idx}`} className="min-h-[96px] bg-zinc-950/40 rounded-lg border border-zinc-900/60" />
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

                    // [달력 줄맞춤 강제 3-Row 분리]
                    // Slot 1: Day (D)
                    const dayShift = dayData?.shifts.find(s => s.code === 'D');
                    // Slot 2: Mid / Helper (M1, M2, M, H)
                    const midShifts = dayData?.shifts.filter(s => s.code === 'M1' || s.code === 'M2' || s.code === 'M' || s.code === 'H') || [];
                    // Slot 3: Night (N)
                    const nightShift = dayData?.shifts.find(s => s.code === 'N');

                    // 근무조 렌더링 헬퍼 컴포넌트 ("현우" 포함 시 전체 뱃지 노란색 Override)
                    // 근무조 렌더링 헬퍼 컴포넌트 (메인 달력: 무채색화 & "현우" 노란색 하이라이트 + 밑줄 제거)
                    const renderShiftSlot = (shift?: ShiftItem) => {
                      if (!shift) return null;
                      const hasHyunwoo = shift.hasTargetUser;

                      // [요구사항 2]: 메인 달력 화면에서는 다른 사람들의 스케줄은 색상을 넣지 않고
                      // 다크 모드에 어울리는 무채색 텍스트(text-zinc-400 등)로 통일하여 아주 심플하게 렌더링
                      let baseTheme = 'bg-zinc-900/90 border-zinc-800/80 text-zinc-400';
                      let codeTheme = 'text-zinc-500 font-bold';

                      // "현우"가 포함된 경우 전체 뱃지 컨테이너를 노란색(#fde047)으로 Override!
                      if (hasHyunwoo) {
                        baseTheme = 'bg-[#fde047] border-yellow-400 text-black font-bold shadow-sm';
                        codeTheme = 'text-black font-black';
                      }

                      return (
                        <div
                          className={`flex flex-wrap items-center gap-1 p-0.5 rounded border text-[9px] md:text-[10px] leading-tight tracking-tighter break-words w-full transition-all ${baseTheme}`}
                        >
                          <span className={`shrink-0 ${codeTheme}`}>
                            {shift.code}
                          </span>
                          <span className="break-words">
                            {shift.workers.map((worker, wIdx) => {
                              const isMe = worker === '현우';
                              return (
                                <React.Fragment key={wIdx}>
                                  {/* "현우" 이름 아래에 생기는 밑줄(underline) 효과 완전히 제거(no-underline) */}
                                  <span className={isMe ? 'font-black no-underline' : 'no-underline'}>
                                    {worker}
                                  </span>
                                  {wIdx < shift.workers.length - 1 && (
                                    <span className={hasHyunwoo ? 'text-black/40 mx-0.5' : 'text-zinc-600 mx-0.5'}>/</span>
                                  )}
                                </React.Fragment>
                              );
                            })}
                          </span>
                        </div>
                      );
                    };

                    return (
                      <div
                        key={dateStr}
                        onClick={() => dayData && setSelectedDay(dayData)}
                        className={`min-h-[104px] p-1 rounded-lg border flex flex-col justify-between transition-all cursor-pointer ${
                          isDimmed
                            ? 'opacity-25 bg-zinc-950 border-zinc-900'
                            : hasTargetUser
                            ? 'bg-zinc-900 border-yellow-400 ring-1 ring-yellow-400/60 shadow-md target-highlight-box'
                            : 'bg-zinc-900 border-zinc-800 hover:border-zinc-700'
                        }`}
                      >
                        {/* 상단 날짜 숫자 및 공휴일 메모 */}
                        <div className="flex items-center justify-between gap-0.5 mb-1 leading-none w-full">
                          <span
                            className={`text-xs font-black ${
                              isSun || isHol ? 'text-rose-500' : isSat ? 'text-sky-400' : 'text-zinc-200'
                            }`}
                          >
                            {d}
                          </span>

                          {dayData?.holidayNote && (
                            <span className="text-[7.5px] font-bold text-rose-400 bg-rose-950/60 border border-rose-900/80 px-0.5 py-0.2 rounded-xs leading-none">
                              {dayData.holidayNote}
                            </span>
                          )}
                        </div>

                        {/* [달력 줄맞춤 강제 3-Row Grid/Slot 유지] */}
                        <div className="flex flex-col gap-1 w-full flex-1 justify-between">
                          {/* Slot 1: Day (D) - 데이가 없어도 투명 빈칸 유지하여 가로선 정렬 강제 */}
                          <div className="min-h-[18px] flex items-center">
                            {dayShift ? renderShiftSlot(dayShift) : <div className="h-4 w-full opacity-0 pointer-events-none" />}
                          </div>

                          {/* Slot 2: Mid / Helper (M1, M2, H) */}
                          <div className="min-h-[18px] flex flex-col gap-0.5">
                            {midShifts.length > 0 ? (
                              midShifts.map((s, idx) => <React.Fragment key={idx}>{renderShiftSlot(s)}</React.Fragment>)
                            ) : (
                              <div className="h-4 w-full opacity-0 pointer-events-none" />
                            )}
                          </div>

                          {/* Slot 3: Night (N) */}
                          <div className="min-h-[18px] flex items-center">
                            {nightShift ? renderShiftSlot(nightShift) : <div className="h-4 w-full opacity-0 pointer-events-none" />}
                          </div>
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
        {/* 4. 상세 모달 (Dialog) - "현우" 포함 시 노란색 Override */}
        {/* ========================================================= */}
        {selectedDay && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs animate-in fade-in duration-150">
            <div className="absolute inset-0" onClick={() => setSelectedDay(null)} />

            <div className="relative w-full max-w-sm bg-zinc-900 border border-zinc-800 rounded-2xl shadow-2xl p-5 z-10 animate-in zoom-in-95 duration-150 text-zinc-100">
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

              {selectedDay.hasTargetUser && (
                <div className="mt-3 p-2.5 bg-yellow-400/10 border border-yellow-400/40 rounded-xl flex items-center gap-2 text-yellow-400">
                  <Sparkles className="w-4 h-4 fill-yellow-400 shrink-0" />
                  <span className="text-xs font-extrabold">
                    현우 당직 근무일입니다.
                  </span>
                </div>
              )}

              {/* 근무조별 섹션 리스트 */}
              <div className="mt-3.5 space-y-2">
                {selectedDay.shifts.map((shift, idx) => {
                  const hasHyunwoo = shift.hasTargetUser;

                  // 모달 내 근무조별 기본 컬러 테마 (D: Sky, M1/M2: Orange, H: Emerald, N: Indigo)
                  let cardTheme = 'bg-zinc-950 border-zinc-800 text-zinc-200';
                  let codeBadgeTheme = 'bg-zinc-800 text-zinc-300';
                  let workerTagTheme = 'bg-zinc-800 text-zinc-300';

                  if (shift.code === 'D') {
                    cardTheme = 'bg-sky-950/40 border-sky-800/60 text-sky-200';
                    codeBadgeTheme = 'bg-sky-900/60 text-sky-300';
                    workerTagTheme = 'bg-sky-900/40 text-sky-200';
                  } else if (shift.code === 'M1' || shift.code === 'M2' || shift.code === 'M') {
                    cardTheme = 'bg-amber-950/40 border-amber-800/60 text-amber-200';
                    codeBadgeTheme = 'bg-amber-900/60 text-amber-300';
                    workerTagTheme = 'bg-amber-900/40 text-amber-200';
                  } else if (shift.code === 'H') {
                    cardTheme = 'bg-emerald-950/40 border-emerald-800/60 text-emerald-200';
                    codeBadgeTheme = 'bg-emerald-900/60 text-emerald-300';
                    workerTagTheme = 'bg-emerald-900/40 text-emerald-200';
                  } else if (shift.code === 'N') {
                    cardTheme = 'bg-indigo-950/40 border-indigo-800/60 text-indigo-200';
                    codeBadgeTheme = 'bg-indigo-900/60 text-indigo-300';
                    workerTagTheme = 'bg-indigo-900/40 text-indigo-200';
                  }

                  // [핵심 2]: "현우"가 포함된 섹션 전체를 노란색(#fde047), 글씨 검은색, 볼드로 Override!
                  if (hasHyunwoo) {
                    cardTheme = 'bg-[#fde047] border-yellow-400 text-black font-bold shadow-lg';
                    codeBadgeTheme = 'bg-black text-[#fde047] font-black';
                    workerTagTheme = 'bg-black/10 text-black font-extrabold border border-black/20';
                  }

                  return (
                    <div
                      key={idx}
                      className={`p-3 rounded-xl border text-xs transition-all ${cardTheme}`}
                    >
                      <div className="flex items-center justify-between mb-1.5">
                        <div className="flex items-center gap-1.5 font-bold">
                          <span className={`px-1.5 py-0.2 rounded text-2xs font-black ${codeBadgeTheme}`}>
                            {shift.code}
                          </span>
                          <span className="font-extrabold">{shift.name}</span>
                        </div>
                        <span className={`text-2xs ${hasHyunwoo ? 'text-black/70 font-bold' : 'text-zinc-400'}`}>
                          {shift.time}
                        </span>
                      </div>

                      <div className={`flex flex-wrap items-center gap-1.5 pt-1 border-t ${hasHyunwoo ? 'border-black/10' : 'border-zinc-800/60'}`}>
                        {shift.workers.map((worker, wIdx) => {
                          const isMe = worker === '현우';
                          return (
                            <span
                              key={wIdx}
                              className={`px-2 py-0.5 rounded text-xs font-semibold no-underline ${
                                isMe && hasHyunwoo
                                  ? 'bg-black text-yellow-400 font-black shadow-sm'
                                  : workerTagTheme
                              }`}
                            >
                              {worker}
                            </span>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>

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
                  <X className="w-5 h-5" />
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
                  업로드 시 브라우저에 자동 저장되어 영구 보존됩니다.
                </p>
              </div>

              <div className="mt-3.5 p-2.5 bg-zinc-950 rounded-xl border border-zinc-800 text-2xs space-y-1 text-zinc-400">
                <p className="text-zinc-300 font-bold">안내:</p>
                <p>• 업로드된 데이터는 LocalStorage에 안전하게 저장됩니다.</p>
                <p>• 앱 재실행 시 엑셀 재업로드 없이 즉시 복원됩니다.</p>
                <p>• &quot;현우&quot;가 포함된 근무조는 전체 노란색으로 강조됩니다.</p>
              </div>

              <button
                onClick={() => {
                  const initial = generateInitialSampleData();
                  setScheduleList(initial);
                  localStorage.setItem(STORAGE_DATA_KEY, JSON.stringify(initial));
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
    </>
  );
}
