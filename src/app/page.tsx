'use client';

import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  ChevronLeft, 
  ChevronRight, 
  ChevronDown,
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
  KeyRound,
  Save
} from 'lucide-react';

const FULL_BLEED_ICON_SVG =
  "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1MTIgNTEyIiB3aWR0aD0iNTEyIiBoZWlnaHQ9IjUxMiI+CiAgPCEtLSDqvYkg7LCsIOqygOydgOyDiSDsgqzqsIHtmJUg67Cw6rK9IChGdWxsIEJsZWVkKSAtLT4KICA8cmVjdCB3aWR0aD0iNTEyIiBoZWlnaHQ9IjUxMiIgZmlsbD0iIzAwMDAwMCIvPgogIDxnIHRyYW5zZm9ybT0idHJhbnNsYXRlKDEyMSwgMTE2KSI+CiAgICA8IS0tIOy6mOumsOuNlCDtnbDsg4kg67O47LK0IC0tPgogICAgPHJlY3QgeD0iMCIgeT0iMTYiIHdpZHRoPSIyNzAiIGhlaWdodD0iMjgwIiByeD0iMzYiIGZpbGw9IiNGRkZGRkYiLz4KICAgIDwhLS0g7LqY66aw642UIOyDgeuLqCDruajqsITsg4kg7Zek642UIC0tPgogICAgPHBhdGggZD0iTSAwIDUyIEMgMCAzMiAxNiAxNiAzNiAxNiBMIDIzNCAxNiBDIDI1NCAxNiAyNzAgMzIgMjcwIDUyIEwgMjcwIDg4IEwgMCA4OCBaIiBmaWxsPSIjRUU0MzQzIi8+CiAgICA8IS0tIOyDgeuLqCAy6rCc7J2YIOqygOydgOyDiSDrsJTsnbjrjZQg66eBIC0tPgogICAgPHJlY3QgeD0iNTIiIHk9IjAiIHdpZHRoPSIyMCIgaGVpZ2h0PSIzNCIgcng9IjEwIiBmaWxsPSIjMDAwMDAwIi8+CiAgICA8cmVjdCB4PSIxOTgiIHk9IjAiIHdpZHRoPSIyMCIgaGVpZ2h0PSIzNCIgcng9IjEwIiBmaWxsPSIjMDAwMDAwIi8+CiAgICA8IS0tIOygleykkeyVmSDruajqsITsg4kg7Iut7J6QIOuniO2BrCAtLT4KICAgIDxyZWN0IHg9Ijc1IiB5PSIxNjAiIHdpZHRoPSIxMjAiIGhlaWdodD0iNDIiIHJ4PSIxNCIgZmlsbD0iI0VFNDM0MyIvPgogICAgPHJlY3QgeD0iMTE0IiB5PSIxMjEiIHdpZHRoPSI0MiIgaGVpZ2h0PSIxMjAiIHJ4PSIxNCIgZmlsbD0iI0VFNDM0MyIvPgogIDwvZz4KPC9zdmc+";

const STORAGE_DATA_KEY = 'er_schedule_data_v2';
const STORAGE_AUTH_KEY = 'er_schedule_auth_token_v2';
const STORAGE_PIN_KEY = 'er_schedule_custom_pin_v2';
const STORAGE_WEBAUTHN_ID = 'er_schedule_webauthn_cred_id';
const STORAGE_WEBAUTHN_REGISTERED = 'er_schedule_webauthn_registered_v3';
const DEFAULT_PASSCODE = process.env.NEXT_PUBLIC_APP_PASSWORD || '1234';

// WebAuthn 버퍼 변환 유틸
function bufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}

function base64ToBuffer(base64: string): ArrayBuffer {
  const binary = window.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

// ==========================================
// 1. 타입 정의
// ==========================================
type ShiftCode = 'D' | 'M1' | 'M2' | 'M' | 'H' | 'N';

interface ShiftItem {
  code: ShiftCode;
  name: string;        // "데이", "미드1", "미드2", "주말 헬퍼" 등
  time: string;        // "08:00 - 15:00"
  workers: string[];   // ["김민준", "현우"]
  hasTargetUser: boolean; // "현우" 포함 여부
}

interface ParsedDay {
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

type ViewMode = '1month' | '1year';

// ==========================================
// [날짜 표준화 헬퍼]: 무조건 로컬 기준의 명확한 YYYY-MM-DD (예: 2026-09-01) 포맷 보장
// ==========================================
function normalizeDateKey(year: number | string, month: number | string, day: number | string): string {
  const y = String(year).trim();
  const m = String(month).trim().padStart(2, '0');
  const d = String(day).trim().padStart(2, '0');
  return `${y}-${m}-${d}`;
}

// ==========================================
// 2. 엑셀 파싱 핵심 알고리즘 (모든 시트 순회 & 매달 1일 누락 방지 & 1Y 완벽 지원)
// ==========================================
function parseExcelData(fileBuffer: ArrayBuffer): ParsedDay[] {
  const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: false });
  if (!workbook.SheetNames.length) return [];

  const parsedData: ParsedDay[] = [];
  const processedDateKeys = new Set<string>();

  // [버그 2 해결]: 엑셀 파일 내의 모든 시트를 순회하여 1년 12개월 데이터를 누락 없이 파싱
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: '',
    });

    let currentYearMonth: string | null = null;
    let fallbackYear = 2026;

    // 시트 이름에서 연/월 유추 (예: "2026.09", "2026-9", "9월")
    const sheetYmMatch = sheetName.match(/(\d{4})[.\-년\s]+(\d{1,2})/);
    if (sheetYmMatch) {
      fallbackYear = parseInt(sheetYmMatch[1], 10);
      currentYearMonth = `${fallbackYear}-${String(parseInt(sheetYmMatch[2], 10)).padStart(2, '0')}`;
    } else {
      const sheetOnlyMonth = sheetName.match(/(\d{1,2})월/);
      if (sheetOnlyMonth) {
        currentYearMonth = `${fallbackYear}-${String(parseInt(sheetOnlyMonth[1], 10)).padStart(2, '0')}`;
      }
    }

    for (let r = 0; r < rawRows.length; r++) {
      const row = rawRows[r] || [];

      // [월 식별]: 셀에서 다양한 형태의 월 표기(2026.9월, 2026년 9월, 2026-09 등) 감지
      for (let c = 0; c < row.length; c++) {
        const cellStr = String(row[c] || '').trim();
        // 1) 연도와 월이 함께 있는 경우 (예: "2026.09월", "2026년 9월", "2026. 9")
        const fullMatch = cellStr.match(/(\d{4})[.\-년\s]+(\d{1,2})월?/);
        if (fullMatch) {
          fallbackYear = parseInt(fullMatch[1], 10);
          const m = String(parseInt(fullMatch[2], 10)).padStart(2, '0');
          currentYearMonth = `${fallbackYear}-${m}`;
          break;
        }
        // 2) 단독 "M월"인 경우
        const monthOnlyMatch = cellStr.match(/^(\d{1,2})월$/);
        if (monthOnlyMatch) {
          const m = String(parseInt(monthOnlyMatch[1], 10)).padStart(2, '0');
          currentYearMonth = `${fallbackYear}-${m}`;
          break;
        }
      }

      if (!currentYearMonth) continue;

      // [요일 식별]: row에 "일", "월" 등이 나타나면 다음 행부터 날짜 및 근무자 행 탐색
      const rowStrArr = row.map(cell => String(cell || '').trim());
      const hasDaysOfWeek = rowStrArr.includes('일') && rowStrArr.includes('월');

      if (hasDaysOfWeek) {
        let nextR = r + 1;
        let hasSeenDayOneInCurrentBlock = false;

        while (nextR < rawRows.length) {
          const dateRow = rawRows[nextR] || [];
          const shiftRow = rawRows[nextR + 1] || [];

          // 새로운 월 헤더가 나오면 현재 월 블록 파싱 종료
          const hasNextMonthHeader = dateRow.some(cell => {
            const cs = String(cell || '').trim();
            return cs.match(/(\d{4})[.\-년\s]+(\d{1,2})월?/) || cs.match(/^(\d{1,2})월$/);
          });
          if (hasNextMonthHeader) break;

          const isAnotherDayOfWeekRow = dateRow.some(c => String(c).trim() === '일') && dateRow.some(c => String(c).trim() === '월');
          if (isAnotherDayOfWeekRow) {
            nextR++;
            continue;
          }

          let foundValidDateInThisRow = false;

          for (let col = 0; col < 7; col++) {
            const dateCellVal = String(dateRow[col] || '').trim();
            if (!dateCellVal) continue;

            // [버그 1 해결]: "9/1", "9.1", "1일", "1(화)", "01", "1" 등 모든 패턴에서 정확한 '일(Day)' 숫자 추출
            let dayNum: number | null = null;

            // 1) "9/1" 또는 "9.1" 형태인 경우 뒤쪽 숫자를 일자로 추출
            const mdMatch = dateCellVal.match(/\b\d{1,2}[\/.](\d{1,2})\b/);
            if (mdMatch) {
              dayNum = parseInt(mdMatch[1], 10);
            } else {
              // 2) "1일", "1(화)" 또는 숫자만 있는 경우
              const dMatch = dateCellVal.match(/^(\d{1,2})/);
              if (dMatch) {
                dayNum = parseInt(dMatch[1], 10);
              } else {
                const anyNum = dateCellVal.match(/\d+/);
                if (anyNum) dayNum = parseInt(anyNum[0], 10);
              }
            }

            if (dayNum === null || isNaN(dayNum) || dayNum < 1 || dayNum > 31) continue;

            // 1주차에서 1일 이전 요일(지난달 말일 25~31)이 1일 앞에 섞여 있는 경우 방어
            if (!hasSeenDayOneInCurrentBlock && dayNum >= 20) {
              continue;
            }
            if (dayNum === 1) {
              hasSeenDayOneInCurrentBlock = true;
            }

            foundValidDateInThisRow = true;

            // 공휴일 메모 (예: "1(신정)", "25(추석)")
            const memoMatch = dateCellVal.match(/\((.*?)\)/);
            const holidayNote = memoMatch ? memoMatch[1].trim() : null;

            // [버그 1 해결]: 무조건 normalizeDateKey를 사용하여 완벽한 YYYY-MM-DD(예: 2026-09-01)로 포맷 통일
            const [yStr, mStr] = currentYearMonth.split('-');
            const yearNum = parseInt(yStr, 10);
            const monthNum = parseInt(mStr, 10);
            const dateStr = normalizeDateKey(yearNum, monthNum, dayNum);

            // 로컬 날짜 객체로 요일 계산
            const dateObj = new Date(yearNum, monthNum - 1, dayNum);
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
  });

  // 날짜순으로 정렬
  parsedData.sort((a, b) => a.date.localeCompare(b.date));

  return parsedData;
}

// ==========================================
// 3. 초기 탑재용 2026년 1~12월 전체 1년 치(365일) 풀 샘플 데이터 생성
// ==========================================
function generateInitialSampleData(): ParsedDay[] {
  const doctors = ['현우', '김민준', '이서연', '정유진', '최준호', '윤도윤', '강예은', '임재현'];
  const list: ParsedDay[] = [];
  const year = 2026;

  // 1월부터 12월까지 전체 월 생성
  for (let m = 1; m <= 12; m++) {
    const daysInMonth = new Date(year, m, 0).getDate();
    const ym = `${year}-${String(m).padStart(2, '0')}`;

    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = normalizeDateKey(year, m, d);
      const dateObj = new Date(year, m - 1, d);
      const dayOfWeek = dateObj.getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

      let holidayNote: string | null = null;
      if (m === 1 && d === 1) holidayNote = '신정';
      else if (m === 2 && (d >= 16 && d <= 18)) holidayNote = '설연휴';
      else if (m === 3 && d === 1) holidayNote = '삼일절';
      else if (m === 5 && d === 5) holidayNote = '어린이날';
      else if (m === 5 && d === 24) holidayNote = '부처님오신날';
      else if (m === 6 && d === 6) holidayNote = '현충일';
      else if (m === 8 && d === 15) holidayNote = '광복절';
      else if (m === 9 && (d === 24 || d === 26)) holidayNote = '추석연휴';
      else if (m === 9 && d === 25) holidayNote = '추석';
      else if (m === 10 && d === 3) holidayNote = '개천절';
      else if (m === 10 && d === 9) holidayNote = '한글날';
      else if (m === 12 && d === 25) holidayNote = '크리스마스';

      const isWeekendOrHol = isWeekend || Boolean(holidayNote);
      const seed = (year * 365 + m * 31 + d) % doctors.length;
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
  }

  return list;
}

// ==========================================
// 4. 메인 뷰 컴포넌트 (초-미니멀 Face ID / PIN 잠금화면 & 캘린더)
// ==========================================
export default function ERSchedulePage() {
  // [요구사항 1]: 잠금화면 강제 - 앱 로드 및 새로고침 시 무조건 false 로 시작
  const [isUnlocked, setIsUnlocked] = useState<boolean>(false);
  const [pinInput, setPinInput] = useState('');
  const [pinError, setPinError] = useState(false);

  // 수동 저장 성공 피드백 상태
  const [saveSuccess, setSaveSuccess] = useState(false);

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

  // [요구사항 1]: 아이폰 네이티브 날짜 선택기(Date Picker) 상태
  const [isDatePickerOpen, setIsDatePickerOpen] = useState(false);
  const [tempPickerYear, setTempPickerYear] = useState<number>(2026);
  const [tempPickerMonth, setTempPickerMonth] = useState<number>(9);
  const monthInputRef = useRef<HTMLInputElement>(null);

  // [요구사항 1]: 기기에 등록된 Face ID 패스키 존재 여부 상태
  const [hasRegisteredPasskey, setHasRegisteredPasskey] = useState<boolean>(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // [핵심 1-1]: LocalStorage에서 스케줄 데이터 안전하게 불러오기
  const loadSavedSchedule = () => {
    if (typeof window === 'undefined') return false;
    try {
      const cachedData = localStorage.getItem(STORAGE_DATA_KEY);
      if (cachedData) {
        const parsed = JSON.parse(cachedData);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setScheduleList(parsed);
          const [y, m] = parsed[0].yearMonth.split('-').map(Number);
          setCurrentYear(y);
          setCurrentMonth(m);
          return true;
        }
      }
    } catch (e) {
      console.error('Failed to load cached schedule data:', e);
    }
    return false;
  };

  // [핵심 1-2 & 요구사항 1 - 단계 B]: 앱 시작 시 잠금화면 강제 & 패스키 기록이 있을 때만 Face ID(Get) 자동 호출
  useEffect(() => {
    setIsUnlocked(false);
    loadSavedSchedule();

    if (typeof window !== 'undefined') {
      const isReg = localStorage.getItem(STORAGE_WEBAUTHN_REGISTERED) === 'true';
      const existingCredId = localStorage.getItem(STORAGE_WEBAUTHN_ID);
      const isActuallyRegistered = Boolean(isReg && existingCredId);
      setHasRegisteredPasskey(isActuallyRegistered);

      // 등록된 패스키 기록이 있을 때만 get() 자동 실행! 기록이 없다면 조용히 PIN 화면 유지 ("일치하는 패스키 없음" 에러 원천 차단)
      if (isActuallyRegistered) {
        const timer = setTimeout(() => {
          handleAuthFaceID(true).catch(err => {
            console.warn('Face ID 자동 실행 에러 (보안 정책 등으로 차단된 경우 수동 터치 가능):', err);
          });
        }, 150);
        return () => clearTimeout(timer);
      }
    }
  }, []);

  // [핵심 1-3]: 수동 저장 함수
  const handleManualSave = () => {
    if (typeof window !== 'undefined') {
      try {
        localStorage.setItem(STORAGE_DATA_KEY, JSON.stringify(scheduleList));
        setSaveSuccess(true);
        setTimeout(() => setSaveSuccess(false), 2000);
      } catch (e) {
        console.error('Failed to save to localStorage:', e);
        alert('저장 중 오류가 발생했습니다.');
      }
    }
  };

  // [요구사항 1 - 단계 A]: Face ID (패스키) 신규 기기 등록 (Create)
  const handleRegisterFaceID = async () => {
    if (typeof window === 'undefined' || !window.PublicKeyCredential || !navigator.credentials) {
      alert('현재 브라우저/기기 환경에서 Face ID(WebAuthn)를 지원하지 않습니다.');
      return;
    }

    try {
      const challenge = new Uint8Array(32);
      window.crypto.getRandomValues(challenge);

      const credential = (await navigator.credentials.create({
        publicKey: {
          challenge,
          rp: {
            name: 'ER Schedule',
            id: window.location.hostname,
          },
          user: {
            id: Uint8Array.from('er-schedule-user-hyunwoo', c => c.charCodeAt(0)),
            name: 'hyunwoo',
            displayName: '박현우 (ER Schedule)',
          },
          pubKeyCredParams: [
            { alg: -7, type: 'public-key' },   // ES256 (Apple Face ID 표준)
            { alg: -257, type: 'public-key' },  // RS256
          ],
          authenticatorSelection: {
            authenticatorAttachment: 'platform', // 아이폰 Face ID / Touch ID 강제
            userVerification: 'required',
            residentKey: 'preferred',
          },
          timeout: 60000,
        },
      })) as PublicKeyCredential | null;

      if (credential) {
        const rawIdBase64 = bufferToBase64(credential.rawId);
        localStorage.setItem(STORAGE_WEBAUTHN_ID, rawIdBase64);
        localStorage.setItem(STORAGE_WEBAUTHN_REGISTERED, 'true');
        setHasRegisteredPasskey(true);
        alert('Face ID(패스키)가 기기에 안전하게 등록되었습니다!\n이제부터 앱 접속 시 Face ID로 자동 잠금 해제됩니다.');
        loadSavedSchedule();
        setIsUnlocked(true);
      }
    } catch (err: any) {
      console.error('Face ID 등록 오류:', err);
      if (err.name === 'NotAllowedError') {
        // 사용자가 취소한 경우 조용히 종료
        return;
      }
      alert('Face ID 등록 중 오류가 발생했습니다. (HTTPS 환경 및 Safari 권한 확인 필요)');
    }
  };

  // [요구사항 1 - 단계 B]: Face ID (패스키) 생체 인증 (Get)
  const handleAuthFaceID = async (isAuto = false) => {
    if (typeof window === 'undefined' || !window.PublicKeyCredential || !navigator.credentials) {
      if (!isAuto) {
        alert('현재 브라우저/기기 환경에서 Face ID(WebAuthn)를 지원하지 않습니다. 암호(PIN)를 입력해 주세요.');
      }
      return;
    }

    const isReg = localStorage.getItem(STORAGE_WEBAUTHN_REGISTERED) === 'true';
    const existingCredId = localStorage.getItem(STORAGE_WEBAUTHN_ID);

    // [핵심]: 패스키가 등록되어 있지 않은 상태 처리
    if (!isReg || !existingCredId) {
      if (isAuto) {
        // 자동 실행일 때는 조용히 PIN 키패드 화면을 유지
        return;
      }
      // 사용자가 버튼을 직접 눌렀을 때는 등록으로 친절히 안내
      const confirmRegister = confirm('현재 기기에 등록된 Face ID 패스키가 없습니다.\n지금 Face ID 기기 등록을 진행하시겠습니까?');
      if (confirmRegister) {
        handleRegisterFaceID();
      }
      return;
    }

    try {
      const challenge = new Uint8Array(32);
      window.crypto.getRandomValues(challenge);

      const assertion = await navigator.credentials.get({
        publicKey: {
          challenge,
          timeout: 60000,
          rpId: window.location.hostname,
          allowCredentials: [
            {
              id: base64ToBuffer(existingCredId),
              type: 'public-key',
              transports: ['internal'],
            },
          ],
          userVerification: 'required',
        },
      });

      if (assertion) {
        loadSavedSchedule();
        setIsUnlocked(true);
        return;
      }
    } catch (err: any) {
      console.warn('Face ID get() 인증 실패 또는 불일치:', err);

      // [핵심 에러 해결]: "일치하는 패스키가 없습니다" (NotFoundError 등) 발생 시 잘못된 로컬 키를 자동 제거하여 다음 실행 시 에러 방지
      if (err.name === 'NotFoundError' || (err.message && err.message.toLowerCase().includes('match'))) {
        localStorage.removeItem(STORAGE_WEBAUTHN_ID);
        localStorage.removeItem(STORAGE_WEBAUTHN_REGISTERED);
        setHasRegisteredPasskey(false);
      }

      if (isAuto) {
        // 자동 실행 실패 시 조용히 넘어가서 PIN 입력 대기
        return;
      }
      if (err.name === 'NotAllowedError') {
        // 사용자가 취소한 경우 리턴
        return;
      }
      alert('Face ID 인증에 실패했습니다. 암호(PIN)로 잠금을 해제하거나 하단에서 [Face ID 기기 등록]을 다시 진행해 주세요.');
    }
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
          loadSavedSchedule();
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

  // [버그 1 완벽 해결]: 날짜 맵 (모든 날짜 키의 정규화 보장 - YYYY-MM-DD 매칭 100% 보장)
  const scheduleMap = useMemo(() => {
    const map = new Map<string, ParsedDay>();
    scheduleList.forEach(day => {
      // 1) 원본 키 매핑
      map.set(day.date, day);

      // 2) YYYY-MM-DD 정규화 키 매핑 (1일이 2026-09-01로 100% 일치)
      const parts = day.date.split('-');
      if (parts.length === 3) {
        const normKey = normalizeDateKey(parts[0], parts[1], parts[2]);
        map.set(normKey, day);

        // 비패딩 키(예: 2026-9-1)도 혹시 모를 상황을 대비해 등록
        const unpaddedKey = `${parseInt(parts[0], 10)}-${parseInt(parts[1], 10)}-${parseInt(parts[2], 10)}`;
        map.set(unpaddedKey, day);
      }
    });
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

  // [요구사항 2]: Now 클릭 시 무조건 현재 실제 날짜가 속한 이번 달로 즉시 이동
  const handleGoToCurrentMonth = () => {
    const today = new Date();
    setCurrentYear(today.getFullYear());
    setCurrentMonth(today.getMonth() + 1);
    setViewMode('1month');
  };

  // 월 / 연도 이동 (1M 모드일 땐 1달씩, 1Y 모드일 땐 1년씩)
  const handlePrevMonth = () => {
    if (viewMode === '1year') {
      setCurrentYear(prev => prev - 1);
    } else {
      const prev = new Date(currentYear, currentMonth - 1 - 1, 1);
      setCurrentYear(prev.getFullYear());
      setCurrentMonth(prev.getMonth() + 1);
    }
  };

  const handleNextMonth = () => {
    if (viewMode === '1year') {
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

  // [요구사항 2]: 렌더링할 월 배열 (1M: 1달, 1Y: 선택된 연도의 1월~12월 전체를 아래로 스크롤하여 조회)
  const monthsToRender = useMemo(() => {
    if (viewMode === '1year') {
      const arr: { year: number; month: number }[] = [];
      for (let m = 1; m <= 12; m++) {
        arr.push({ year: currentYear, month: m });
      }
      return arr;
    }
    // 1month 모드
    return [{ year: currentYear, month: currentMonth }];
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
        {/* 상단 인포 (불필요한 텍스트 제거) */}
        <div className="w-full flex items-center justify-end pt-safe text-zinc-500 text-xs">
          <span className="text-[11px] font-semibold tracking-wider text-zinc-600">ER Schedule</span>
        </div>

        {/* 중앙 Face ID & PIN 헤더 */}
        <div className="flex flex-col items-center my-auto w-full max-w-xs">
          {/* Face ID 인터랙션 버튼 */}
          <button
            type="button"
            onClick={() => handleAuthFaceID(false)}
            className="w-20 h-20 rounded-3xl bg-zinc-900 border border-zinc-800 hover:border-yellow-400/80 active:scale-95 transition-all flex flex-col items-center justify-center text-yellow-400 mb-8 shadow-xl group cursor-pointer"
            title="Face ID"
          >
            <ScanFace className="w-10 h-10 group-hover:scale-110 transition-transform" />
            <span className="text-[9px] font-black text-zinc-400 mt-1 uppercase tracking-tighter">Face ID</span>
          </button>

          {/* PIN 4자리 표시 인디케이터 (텍스트 힌트 없이 깔끔하게 도트만 표시) */}
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
              onClick={() => handleAuthFaceID(false)}
              className="w-16 h-16 rounded-full text-yellow-400 hover:text-yellow-300 font-bold text-xs mx-auto flex items-center justify-center"
            >
              Face ID
            </button>
          </div>
        </div>

        {/* [요구사항 1 - 단계 A]: 하단 Face ID 기기 등록 & 비밀번호 변경 버튼 */}
        <div className="pb-safe flex flex-col items-center gap-2">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleRegisterFaceID}
              className="text-2xs text-yellow-400 hover:text-yellow-300 transition-all font-bold py-2 px-3 bg-zinc-900 hover:bg-zinc-800 rounded-full border border-yellow-400/40 hover:border-yellow-400 flex items-center gap-1.5 active:scale-95 shadow-sm"
              title="이 기기에 Face ID 패스키 등록"
            >
              <ScanFace className="w-3.5 h-3.5 text-yellow-400" />
              <span>{hasRegisteredPasskey ? 'Face ID 재등록' : 'Face ID(패스키) 기기 등록'}</span>
            </button>

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
              className="text-2xs text-zinc-400 hover:text-zinc-200 transition-all font-medium py-2 px-3 bg-zinc-900/80 hover:bg-zinc-800 rounded-full border border-zinc-800 flex items-center gap-1.5 active:scale-95 shadow-xs"
            >
              <KeyRound className="w-3.5 h-3.5 text-zinc-400" />
              <span>비밀번호 변경</span>
            </button>
          </div>

          {hasRegisteredPasskey && (
            <span className="text-[10px] text-zinc-500 font-medium">
              ✓ Face ID 등록 기기 (앱 진입 시 자동 잠금해제)
            </span>
          )}
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
        {/* 1. 상단 심플 헤더: [ER Schedule] 로고 폰트 확대 & 여백 확장 */}
        {/* ========================================================= */}
        <header className="sticky top-0 z-30 pt-safe bg-zinc-950/95 backdrop-blur-md border-b border-zinc-800/80 px-5 py-4 flex items-center justify-between">
          <h1 className="text-2xl font-black tracking-tight text-white flex items-center gap-2">
            ER Schedule
          </h1>

          <div className="flex items-center gap-2">
            {/* [요구사항 1]: 수동 저장 버튼 */}
            <button
              onClick={handleManualSave}
              className={`flex items-center gap-1.5 text-xs font-bold px-3 py-2 rounded-xl border transition-all shadow-sm active:scale-95 ${
                saveSuccess
                  ? 'bg-emerald-950/80 border-emerald-500 text-emerald-300'
                  : 'bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border-zinc-700'
              }`}
              title="현재 스케줄 데이터 브라우저에 영구 저장"
            >
              {saveSuccess ? (
                <>
                  <Check className="w-4 h-4 text-emerald-400" />
                  <span>저장됨</span>
                </>
              ) : (
                <>
                  <Save className="w-4 h-4 text-yellow-400" />
                  <span>저장</span>
                </>
              )}
            </button>

            {/* 엑셀 업로드 버튼 */}
            <button
              onClick={() => setIsUploadOpen(true)}
              className="flex items-center gap-1.5 bg-zinc-900 hover:bg-zinc-800 active:scale-95 text-zinc-100 text-xs font-bold px-3.5 py-2 rounded-xl border border-zinc-700 transition-all shadow-sm"
            >
              <Upload className="w-4 h-4 text-yellow-400" />
              <span>엑셀 업로드</span>
            </button>
          </div>
        </header>

        {/* ========================================================= */}
        {/* 2. 달력 컨트롤: [YYYY년 M월] 확대 및 클릭 시 네이티브 Date Picker 열기, [Now / 1M / 1Y] 탭 */}
        {/* ========================================================= */}
        <div className="px-4 pt-2.5 pb-2 space-y-2 border-b border-zinc-900 bg-zinc-950">
          <div className="flex items-center justify-between">
            {/* 이전/다음 달 이동 및 대형 [YYYY년 M월] 버튼 */}
            <div className="flex items-center gap-1.5">
              <button
                onClick={handlePrevMonth}
                aria-label="이전"
                className="p-2 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border border-zinc-800 transition-colors active:scale-95"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>

              {/* [요구사항 1]: 중앙의 "YYYY년 M월" 텍스트 크기를 키우고 굵게(font-bold, text-xl) + 탭하면 Date Picker 오픈 */}
              <button
                type="button"
                onClick={() => {
                  setTempPickerYear(currentYear);
                  setTempPickerMonth(currentMonth);
                  setIsDatePickerOpen(true);
                }}
                className="text-xl font-bold text-white hover:text-yellow-400 active:scale-95 transition-all px-2 py-1 rounded-xl hover:bg-zinc-900 flex items-center gap-1 tracking-tight group cursor-pointer"
                title="연도 및 월 선택 (Date Picker)"
              >
                <span>{currentYear}년 {currentMonth}월</span>
                <ChevronDown className="w-4 h-4 text-zinc-400 group-hover:text-yellow-400 transition-colors" />
              </button>

              <button
                onClick={handleNextMonth}
                aria-label="다음"
                className="p-2 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border border-zinc-800 transition-colors active:scale-95"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>

            {/* [요구사항 2]: 타임프레임 필터 (Now / 1M / 1Y) */}
            <div className="flex items-center bg-zinc-900 p-1 rounded-xl border border-zinc-800 shadow-inner">
              <button
                type="button"
                onClick={handleGoToCurrentMonth}
                className="text-xs font-bold px-2.5 py-1.5 rounded-lg text-zinc-400 hover:text-yellow-400 active:scale-95 transition-all"
                title="현재 실제 날짜의 이번 달로 즉시 이동"
              >
                Now
              </button>
              <button
                type="button"
                onClick={() => setViewMode('1month')}
                className={`text-xs font-bold px-2.5 py-1.5 rounded-lg transition-all active:scale-95 ${
                  viewMode === '1month'
                    ? 'bg-zinc-800 text-yellow-400 shadow-sm'
                    : 'text-zinc-400 hover:text-zinc-200'
                }`}
                title="1개월씩 달력 보기"
              >
                1M
              </button>
              <button
                type="button"
                onClick={() => setViewMode('1year')}
                className={`text-xs font-bold px-2.5 py-1.5 rounded-lg transition-all active:scale-95 ${
                  viewMode === '1year'
                    ? 'bg-zinc-800 text-yellow-400 shadow-sm'
                    : 'text-zinc-400 hover:text-zinc-200'
                }`}
                title="연간 전체 달력 스크롤 보기"
              >
                1Y
              </button>
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
              className={`flex items-center gap-1 px-2.5 py-1 rounded-lg font-bold text-2xs transition-all border ${
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
        {/* 3. 달력 그리드 영역 (세로 여백 상향 조정으로 화면 상단 쪽에 안정감 있게 배치) */}
        {/* ========================================================= */}
        <main className="flex-1 p-2.5 pt-2 space-y-5 pb-safe">
          {monthsToRender.map(({ year, month }) => {
            const ymStr = `${year}-${String(month).padStart(2, '0')}`;
            const firstDayOfWeek = new Date(year, month - 1, 1).getDay();
            const daysInMonth = new Date(year, month, 0).getDate();

            return (
              <div key={ymStr} className="space-y-1">
                {/* 다중 월 모드일 때 월 헤더 */}
                {viewMode !== '1month' && (
                  <div className="flex items-center justify-between px-1 pt-3 pb-1 border-b border-zinc-900/80 mb-1">
                    <span className="text-sm font-black text-yellow-400 flex items-center gap-1.5">
                      <CalendarIcon className="w-3.5 h-3.5" />
                      <span>{year}년 {month}월</span>
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
                    const dateStr = normalizeDateKey(year, month, d);
                    const dayData = scheduleMap.get(dateStr) || scheduleMap.get(`${ymStr}-${String(d).padStart(2, '0')}`) || scheduleMap.get(`${year}-${month}-${d}`);
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

                    // 근무조 렌더링 헬퍼 컴포넌트 (강제 한 줄 처리 + 현우 가운데 정렬 + 무채색화 + M1/M2는 M으로 표기)
                    const renderShiftSlot = (shift?: ShiftItem) => {
                      if (!shift) return null;
                      const hasHyunwoo = shift.hasTargetUser;

                      // [요구사항 2]: 화면에는 M1, M2 대신 무조건 'M'으로 표기
                      const displayCode = (shift.code === 'M1' || shift.code === 'M2') ? 'M' : shift.code;

                      // [요구사항 2 & 3]: 메인 달력 화면 - 다른 사람들은 무채색 통일, 현우는 노란색 + 가운데 정렬
                      let baseTheme = 'bg-zinc-900/90 border-zinc-800/80 text-zinc-400';
                      let codeTheme = 'text-zinc-500 font-bold';
                      let alignClass = 'justify-start text-left';

                      // "현우"가 포함된 경우 전체 뱃지 컨테이너를 노란색(#fde047)으로 Override + 가운데 정렬!
                      if (hasHyunwoo) {
                        baseTheme = 'bg-[#fde047] border-yellow-400 text-black font-bold shadow-sm';
                        codeTheme = 'text-black font-black';
                        alignClass = 'justify-center text-center';
                      }

                      return (
                        <div
                          className={`flex items-center gap-0.5 px-0.5 py-0.5 rounded border leading-none w-full transition-all ${baseTheme} ${alignClass}`}
                          title={`${displayCode} ${shift.workers.join(', ')}`}
                        >
                          <span className={`shrink-0 text-[8px] sm:text-[9px] font-bold leading-none ${codeTheme}`}>
                            {displayCode}
                          </span>
                          {/* [요구사항 2]: truncate, overflow-hidden, whitespace-nowrap 전부 제거 및 break-all 한 줄 압축 */}
                          <span className="text-[8px] sm:text-[9px] leading-none tracking-tighter whitespace-pre-wrap break-all w-full text-center no-underline">
                            {shift.workers.map((worker, wIdx) => {
                              const isMe = worker === '현우';
                              return (
                                <React.Fragment key={wIdx}>
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

                  // [요구사항 2]: 화면에는 M1, M2 대신 무조건 'M'으로 표기
                  const displayCode = (shift.code === 'M1' || shift.code === 'M2') ? 'M' : shift.code;
                  const displayName = (shift.code === 'M1' || shift.code === 'M2') ? '미드' : shift.name;

                  // [요구사항 3]: 팝업(모달)에서도 다른 근무자들은 색상 없이 다크 모드 무채색으로 통일!
                  let cardTheme = 'bg-zinc-950/80 border-zinc-800 text-zinc-300';
                  let codeBadgeTheme = 'bg-zinc-900 border border-zinc-800 text-zinc-400 font-bold';
                  let workerTagTheme = 'bg-zinc-900/60 border border-zinc-800 text-zinc-400';

                  // [요구사항 3]: "현우"가 포함된 섹션 전체를 노란색(#fde047), 글씨 검은색, 볼드 + 가운데 정렬!
                  if (hasHyunwoo) {
                    cardTheme = 'bg-[#fde047] border-yellow-400 text-black font-bold shadow-lg text-center';
                    codeBadgeTheme = 'bg-black text-[#fde047] font-black';
                    workerTagTheme = 'bg-black/15 text-black font-extrabold border border-black/20';
                  }

                  return (
                    <div
                      key={idx}
                      className={`p-3 rounded-xl border text-xs transition-all ${cardTheme}`}
                    >
                      {/* 상단 근무조 코드 및 시간 - 현우 포함 시 가운데 정렬 */}
                      <div className={`flex items-center mb-1.5 ${hasHyunwoo ? 'justify-center gap-3' : 'justify-between'}`}>
                        <div className="flex items-center gap-1.5 font-bold">
                          <span className={`px-1.5 py-0.2 rounded text-2xs font-black ${codeBadgeTheme}`}>
                            {displayCode}
                          </span>
                          <span className="font-extrabold">{displayName}</span>
                        </div>
                        <span className={`text-2xs ${hasHyunwoo ? 'text-black/80 font-bold' : 'text-zinc-400'}`}>
                          {shift.time}
                        </span>
                      </div>

                      {/* 근무자 태그 목록 - 현우 포함 시 justify-center 가운데 정렬 */}
                      <div className={`flex flex-wrap items-center gap-1.5 pt-1.5 border-t ${hasHyunwoo ? 'border-black/15 justify-center' : 'border-zinc-800/60 justify-start'}`}>
                        {shift.workers.map((worker, wIdx) => {
                          const isMe = worker === '현우';
                          return (
                            <span
                              key={wIdx}
                              className={`px-2.5 py-0.5 rounded text-xs font-semibold no-underline ${
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

        {/* ========================================================= */}
        {/* 6. 아이폰 네이티브 날짜 선택기 (iOS Bottom Sheet Wheel Date Picker) */}
        {/* ========================================================= */}
        {isDatePickerOpen && (
          <div className="fixed inset-0 z-50 flex flex-col justify-end bg-black/75 backdrop-blur-xs animate-in fade-in duration-200">
            {/* 백드롭 클릭 시 닫기 */}
            <div className="absolute inset-0" onClick={() => setIsDatePickerOpen(false)} />

            {/* iOS 스타일 바텀 시트 */}
            <div className="relative w-full max-w-md mx-auto bg-zinc-900 border-t border-zinc-800 rounded-t-3xl shadow-2xl p-5 pb-safe z-10 animate-in slide-in-from-bottom duration-250 text-zinc-100 select-none">
              {/* 핸들 바 */}
              <div className="w-10 h-1 bg-zinc-700/80 rounded-full mx-auto mb-3" />

              {/* 상단 컨트롤 바 */}
              <div className="flex items-center justify-between pb-3 border-b border-zinc-800">
                <button
                  type="button"
                  onClick={() => setIsDatePickerOpen(false)}
                  className="text-sm font-semibold text-zinc-400 hover:text-zinc-200 px-2 py-1 transition-colors active:scale-95"
                >
                  취소
                </button>
                <div className="text-sm font-extrabold text-white tracking-tight flex items-center gap-1.5">
                  <CalendarIcon className="w-4 h-4 text-yellow-400" />
                  <span>날짜 선택</span>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setCurrentYear(tempPickerYear);
                    setCurrentMonth(tempPickerMonth);
                    setIsDatePickerOpen(false);
                  }}
                  className="text-sm font-black text-yellow-400 hover:text-yellow-300 px-2 py-1 transition-colors active:scale-95"
                >
                  완료
                </button>
              </div>

              {/* 아이폰 네이티브 감성 듀얼 휠(Wheel) 피커 */}
              <div className="relative my-4 h-48 flex items-center justify-center overflow-hidden bg-zinc-950/60 rounded-2xl border border-zinc-800/80">
                {/* 중앙 하이라이트 밴드 (iOS 휠 선택 표시선) */}
                <div className="absolute left-2 right-2 h-11 border-y border-yellow-400/30 bg-yellow-400/5 rounded-xl pointer-events-none z-0" />

                {/* 좌측: 연도 휠 컬럼 */}
                <div className="flex-1 h-full overflow-y-auto py-18 text-center scroll-smooth z-10 space-y-1">
                  {[2024, 2025, 2026, 2027, 2028, 2029, 2030].map(y => {
                    const isSelected = tempPickerYear === y;
                    return (
                      <div
                        key={y}
                        onClick={() => setTempPickerYear(y)}
                        className={`h-9 flex items-center justify-center cursor-pointer transition-all duration-150 ${
                          isSelected
                            ? 'text-yellow-400 font-black text-lg scale-110 drop-shadow-sm'
                            : 'text-zinc-500 hover:text-zinc-300 text-sm font-medium'
                        }`}
                      >
                        {y}년
                      </div>
                    );
                  })}
                </div>

                <div className="w-px h-32 bg-zinc-800/80" />

                {/* 우측: 월 휠 컬럼 */}
                <div className="flex-1 h-full overflow-y-auto py-18 text-center scroll-smooth z-10 space-y-1">
                  {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map(m => {
                    const isSelected = tempPickerMonth === m;
                    return (
                      <div
                        key={m}
                        onClick={() => setTempPickerMonth(m)}
                        className={`h-9 flex items-center justify-center cursor-pointer transition-all duration-150 ${
                          isSelected
                            ? 'text-yellow-400 font-black text-lg scale-110 drop-shadow-sm'
                            : 'text-zinc-500 hover:text-zinc-300 text-sm font-medium'
                        }`}
                      >
                        {m}월
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* 하단 퀵 액션 버튼 & 네이티브 인풋 연동 */}
              <div className="flex items-center gap-2 pt-1">
                <button
                  type="button"
                  onClick={() => {
                    const now = new Date();
                    setTempPickerYear(now.getFullYear());
                    setTempPickerMonth(now.getMonth() + 1);
                  }}
                  className="flex-1 py-2.5 rounded-xl bg-zinc-800 hover:bg-zinc-700 active:scale-98 text-xs font-bold text-zinc-300 transition-all flex items-center justify-center gap-1"
                >
                  <Sparkles className="w-3.5 h-3.5 text-yellow-400" />
                  <span>이번 달로 맞추기</span>
                </button>

                {/* 네이티브 <input type="month"> 연동 버튼 */}
                <label className="flex-1 py-2.5 rounded-xl bg-zinc-800 hover:bg-zinc-700 active:scale-98 text-xs font-bold text-zinc-300 transition-all flex items-center justify-center gap-1 cursor-pointer relative overflow-hidden">
                  <CalendarIcon className="w-3.5 h-3.5 text-sky-400" />
                  <span>기기 달력으로 선택</span>
                  <input
                    type="month"
                    ref={monthInputRef}
                    value={`${tempPickerYear}-${String(tempPickerMonth).padStart(2, '0')}`}
                    onChange={(e) => {
                      if (e.target.value) {
                        const [y, m] = e.target.value.split('-').map(Number);
                        if (y && m) {
                          setTempPickerYear(y);
                          setTempPickerMonth(m);
                        }
                      }
                    }}
                    className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                  />
                </label>
              </div>
            </div>
          </div>
        )}
      </div>
    </>
  );
}
