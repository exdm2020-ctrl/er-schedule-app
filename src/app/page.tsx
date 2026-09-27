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
  Save,
  Edit2,
  Plus,
  Trash2
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
// [날짜 표준화 헬퍼]: Date 객체 타임존 오차 없이 순수 정수 기반 YYYY-MM-DD 포맷 보장
// ==========================================
function normalizeDateKey(year: number | string, month: number | string, day: number | string): string {
  const y = parseInt(String(year).trim(), 10);
  const m = parseInt(String(month).trim(), 10);
  const d = parseInt(String(day).trim(), 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) {
    return `${year}-${month}-${day}`;
  }
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

// [날짜 추출 헬퍼]: new Date() 타임존 시차 없이 순수 문자열/정수 기반 추출
interface ExtractedDate {
  year: number;
  month: number;
  day: number;
  holidayNote: string | null;
}

function extractDateInfo(
  rawVal: any,
  defaultYear: number,
  defaultMonth: number
): ExtractedDate | null {
  if (rawVal === null || rawVal === undefined) return null;

  // 1) 자바스크립트 Date 객체 직접 처리 (xlsx 라이브러리/브라우저가 Date 객체로 변환한 경우)
  if (rawVal instanceof Date || Object.prototype.toString.call(rawVal) === '[object Date]') {
    const dObj = rawVal as Date;
    if (!isNaN(dObj.getTime())) {
      // Local 시간과 UTC 시간 둘 다 확인 (타임존 시차 극복)
      const localY = dObj.getFullYear();
      const localM = dObj.getMonth() + 1;
      const localD = dObj.getDate();

      const utcY = dObj.getUTCFullYear();
      const utcM = dObj.getUTCMonth() + 1;
      const utcD = dObj.getUTCDate();

      // 한국 표준시(UTC+9)로 인해 UTC는 8월 31일 15:00, Local은 9월 1일 00:00일 수 있음!
      // defaultMonth와 일치하거나 day === 1인 쪽을 최우선 선택
      if (localD === 1 || localM === defaultMonth) {
        return { year: localY, month: localM, day: localD, holidayNote: null };
      }
      if (utcD === 1 || utcM === defaultMonth) {
        return { year: utcY, month: utcM, day: utcD, holidayNote: null };
      }
      return { year: localY, month: localM, day: localD, holidayNote: null };
    }
  }

  const str = String(rawVal).trim();
  if (!str) return null;

  // 공휴일 메모 또는 괄호 요일 (예: "1(신정)", "2026-09-01(신정)", "25(추석)", "1일 (화)")
  const memoMatch = str.match(/\((.*?)\)/);
  const holidayNote = memoMatch ? memoMatch[1].trim() : null;
  // 괄호 메모 제외한 순수 날짜 문자열
  const pureStr = str.replace(/\(.*?\)/g, '').trim();

  // 2) 전체 날짜 포맷 (YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD, YYYY년 M월 D일, 2026-9-1)
  const fullYmdMatch = pureStr.match(/(\d{4})[.\-\/년\s]+(\d{1,2})[.\-\/월\s]+(\d{1,2})/);
  if (fullYmdMatch) {
    const y = parseInt(fullYmdMatch[1], 10);
    const m = parseInt(fullYmdMatch[2], 10);
    const d = parseInt(fullYmdMatch[3], 10);
    if (y >= 2020 && y <= 2035 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return { year: y, month: m, day: d, holidayNote };
    }
  }

  // 3) 8자리 연속 숫자 (예: '20260901', '20261001')
  const ymd8Match = pureStr.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (ymd8Match) {
    const y = parseInt(ymd8Match[1], 10);
    const m = parseInt(ymd8Match[2], 10);
    const d = parseInt(ymd8Match[3], 10);
    if (y >= 2020 && y <= 2035 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return { year: y, month: m, day: d, holidayNote };
    }
  }

  // 4) 미국식 날짜 (MM/DD/YYYY 또는 MM-DD-YYYY, 예: '09/01/2026')
  const mdyMatch = pureStr.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/);
  if (mdyMatch) {
    const m = parseInt(mdyMatch[1], 10);
    const d = parseInt(mdyMatch[2], 10);
    const y = parseInt(mdyMatch[3], 10);
    if (y >= 2020 && y <= 2035 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return { year: y, month: m, day: d, holidayNote };
    }
  }

  // 5) 엑셀 날짜 일련번호 (40000 ~ 60000, 2009년~2064년)
  const numVal = typeof rawVal === 'number' ? rawVal : parseFloat(pureStr);
  if (!isNaN(numVal) && numVal >= 40000 && numVal <= 60000 && /^\d{5}(\.\d+)?$/.test(pureStr)) {
    const serialDays = Math.floor(numVal);
    const epochUtc = (serialDays - 25569) * 86400 * 1000;
    const dateUtc = new Date(epochUtc);
    const utcY = dateUtc.getUTCFullYear();
    const utcM = dateUtc.getUTCMonth() + 1;
    const utcD = dateUtc.getUTCDate();

    // 혹시 모를 한국시간(+9시간) 보정치
    const dateKst = new Date(epochUtc + 9 * 3600 * 1000);
    const kstY = dateKst.getUTCFullYear();
    const kstM = dateKst.getUTCMonth() + 1;
    const kstD = dateKst.getUTCDate();

    if (kstD === 1 || kstM === defaultMonth) {
      return { year: kstY, month: kstM, day: kstD, holidayNote };
    }
    if (utcD === 1 || utcM === defaultMonth) {
      return { year: utcY, month: utcM, day: utcD, holidayNote };
    }
    return { year: utcY, month: utcM, day: utcD, holidayNote };
  }

  // 6) 'M월 D일' 또는 'M/D' 또는 'M.D' 또는 'M-D' (예: '9월 1일', '9/1', '9.1', '10/1')
  const mdMatch = pureStr.match(/^(\d{1,2})[.\-\/월\s]+(\d{1,2})일?$/);
  if (mdMatch) {
    const m = parseInt(mdMatch[1], 10);
    const d = parseInt(mdMatch[2], 10);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return { year: defaultYear, month: m, day: d, holidayNote };
    }
  }

  // 7) 단순 일자 (예: '1', '01', '1일', '22', '23일', '31')
  const dMatch = pureStr.match(/^(\d{1,2})\s*일?$/);
  if (dMatch) {
    const d = parseInt(dMatch[1], 10);
    if (d >= 1 && d <= 31) {
      return { year: defaultYear, month: defaultMonth, day: d, holidayNote };
    }
  }

  return null;
}

// ==========================================
// 2. 엑셀 파싱 핵심 알고리즘 (모든 시트 순회 & 1일 누락 방지 & 10월 이후 전월 파싱 & M1/M2 분리)
// ==========================================
function parseExcelData(fileBuffer: ArrayBuffer): ParsedDay[] {
  const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: false });
  if (!workbook.SheetNames.length) return [];

  const parsedData: ParsedDay[] = [];
  const processedDateKeys = new Set<string>();

  // 근무자 텍스트 정리 함수 (박현우 -> 현우, 슬래시/쉼표/공백 분리)
  const cleanWorkers = (lineStr: string) => {
    if (!lineStr) return [];
    const stripped = lineStr.replace(/^(D|DAY|M|MID|M1|M2|H|HELPER|N|NIGHT|데이|미드|헬퍼|나이트)\s*[:\-\s]\s*/i, '');
    return stripped
      .split(/[\/,\s+&|]+/)
      .map(w => w.trim())
      .filter(Boolean)
      .map(w => (w === '박현우' ? '현우' : w));
  };

  // 엑셀 파일 내의 모든 시트를 순회하여 1년 12개월 데이터를 누락 없이 파싱
  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: '',
    });

    let currentYear = 2026;
    let currentMonth = 9;

    // 시트 이름에서 연/월 유추 (예: "2026.09", "2026-9", "9월", "10월")
    const sheetYmMatch = sheetName.match(/(\d{4})[.\-년\s/]+(\d{1,2})/);
    if (sheetYmMatch) {
      currentYear = parseInt(sheetYmMatch[1], 10);
      currentMonth = parseInt(sheetYmMatch[2], 10);
    } else {
      const sheetOnlyMonth = sheetName.match(/(\d{1,2})월/);
      if (sheetOnlyMonth) {
        currentMonth = parseInt(sheetOnlyMonth[1], 10);
      }
    }

    for (let r = 0; r < rawRows.length; r++) {
      const row = rawRows[r] || [];

      // [월 식별]: 셀에서 다양한 형태의 월 표기(2026.9월, 2026년 10월, 2026-10, 10월 등)를 모두 유연하게 감지
      for (let c = 0; c < row.length; c++) {
        const cellStr = String(row[c] || '').trim();
        if (!cellStr) continue;

        // 1) 연도와 월이 함께 있는 경우 (예: "2026.10월", "2026년 10월", "2026-10", "2026.10")
        const fullMatch = cellStr.match(/(\d{4})[.\-년\s/]+(\d{1,2})(?:월)?/);
        if (fullMatch) {
          // 단, 2026-09-01 같은 특정 일자 셀은 월 타이틀로 오인하지 않음
          if (!cellStr.match(/\d{4}[.\-\/년\s]+\d{1,2}[.\-\/월\s]+\d{1,2}/)) {
            currentYear = parseInt(fullMatch[1], 10);
            currentMonth = parseInt(fullMatch[2], 10);
            break;
          }
        }
        // 2) 단독 또는 제목 내 "M월"인 경우 (예: "10월", "10월 응급실 근무표")
        const monthOnlyMatch = cellStr.match(/(?:^|[^\d])(\d{1,2})월/);
        if (monthOnlyMatch) {
          currentMonth = parseInt(monthOnlyMatch[1], 10);
          break;
        }
      }

      // [요일 식별]: row에 "일", "월" 등이 나타나면 해당 행을 요일 헤더로 인식
      const rowStrArr = row.map(cell => String(cell || '').trim());
      const hasDaysOfWeek = rowStrArr.includes('일') && rowStrArr.includes('월');

      if (hasDaysOfWeek) {
        // [열 인덱스 동적 감지]: '일', '월', '화', '수', '목', '금', '토'가 위치한 컬럼 번호들 추출
        const dayKeywords = ['일', '월', '화', '수', '목', '금', '토'];
        const detectedCols: number[] = [];
        for (let c = 0; c < row.length; c++) {
          const txt = String(row[c] || '').trim();
          if (dayKeywords.some(dk => txt === dk || txt.startsWith(dk))) {
            detectedCols.push(c);
          }
        }
        const targetCols = detectedCols.length >= 7 ? detectedCols.slice(0, 7) : [0, 1, 2, 3, 4, 5, 6];

        let nextR = r + 1;
        let hasSeenDayOneInCurrentBlock = false;

        // [동적 행(Row) 탐색]: 날짜 행을 찾고, 그 아래 첫 번째 유효 근무자 행을 찾아 1:1 매칭
        // 빈 행이나 병합 셀이 끼어 있어도 절대 인덱스가 어긋나지 않아 22일, 23일 등이 누락되지 않음!
        while (nextR < rawRows.length) {
          const candidateRow = rawRows[nextR] || [];

          // 만약 또 다른 요일 행이 나타나면 새로운 달(10월, 11월...)의 시작이므로 현재 블록 종료(break)!
          const isAnotherDayOfWeek = candidateRow.some(c => String(c).trim() === '일') && candidateRow.some(c => String(c).trim() === '월');
          if (isAnotherDayOfWeek) {
            break;
          }

          // 현재 행(candidateRow)에서 유효한 날짜가 있는지 먼저 검사!
          const validDatesInRow: { col: number; dateInfo: ExtractedDate }[] = [];
          for (const col of targetCols) {
            const cellVal = candidateRow[col];
            const info = extractDateInfo(cellVal, currentYear, currentMonth);
            if (info) {
              validDatesInRow.push({ col, dateInfo: info });
            }
          }

          // [월 헤더 검사]: 유효 날짜가 전혀 없는 행인 경우에만 다음 월 헤더 검사를 수행하여 조기 탈출(break) 방지!
          if (validDatesInRow.length === 0) {
            const isNextMonthTitle = candidateRow.some(cell => {
              const cs = String(cell || '').trim();
              if (!cs) return false;
              if (cs.match(/\d{4}[.\-\/년\s]+\d{1,2}[.\-\/월\s]+\d{1,2}/)) return false;
              return Boolean(cs.match(/(\d{4})[.\-년\s/]+(\d{1,2})/) || cs.match(/\d{1,2}월/));
            });
            if (isNextMonthTitle) {
              break; // 다음 달 헤더 발견 시 탈출하여 외부 루프에서 다음 달 파싱 수행!
            }
          }

          // [날짜 행 발견]: 유효 날짜가 1개 이상 들어있는 행
          if (validDatesInRow.length > 0) {
            // 그 아래 행들 중 근무자 데이터가 적혀 있는 첫 번째 비어있지 않은 행(shiftRow)을 동적으로 탐색!
            let shiftRowIndex = nextR + 1;
            while (shiftRowIndex < rawRows.length) {
              const potentialShiftRow = rawRows[shiftRowIndex] || [];
              // 만약 이 행이 새로운 날짜 행이거나 월 헤더면 근무자 행 탐색 중단
              const hasDates = targetCols.some(c => extractDateInfo(potentialShiftRow[c], currentYear, currentMonth));
              if (hasDates) {
                break;
              }
              const isMonthHeader = potentialShiftRow.some(c => {
                const s = String(c || '').trim();
                return Boolean(s.match(/(\d{4})[.\-년\s/]+(\d{1,2})/) || s.match(/\d{1,2}월/));
              });
              if (isMonthHeader) {
                break;
              }
              // 근무자 내용이 들어있는 행인지 확인
              const hasShiftContent = targetCols.some(c => String(potentialShiftRow[c] || '').trim().length > 0);
              if (hasShiftContent) {
                break; // 찾았다! 이 행이 shiftRow!
              }
              shiftRowIndex++;
            }

            const shiftRow = (shiftRowIndex < rawRows.length) ? (rawRows[shiftRowIndex] || []) : [];

            // 이제 발견된 날짜들과 shiftRow의 해당 열 데이터를 1:1 매칭
            for (const { col, dateInfo } of validDatesInRow) {
              const { year: cellYear, month: cellMonth, day: dayNum, holidayNote } = dateInfo;

              // 1주차 전월 말일 방어: 1일을 아직 보지 못했는데 20일 이상인 날짜(전월 30, 31일 등)는 건너뜀
              if (dayNum === 1) {
                hasSeenDayOneInCurrentBlock = true;
              } else if (!hasSeenDayOneInCurrentBlock && dayNum >= 20 && cellMonth === currentMonth) {
                continue;
              }

              // 무조건 normalizeDateKey를 사용하여 완벽한 YYYY-MM-DD(예: 2026-09-01, 2026-10-01)로 포맷 통일
              const dateStr = normalizeDateKey(cellYear, cellMonth, dayNum);
              const yearMonth = `${cellYear}-${String(cellMonth).padStart(2, '0')}`;

              // 요일 계산
              const dateObj = new Date(cellYear, cellMonth - 1, dayNum);
              const dayOfWeek = dateObj.getDay();
              const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

              // [근무자 줄바꿈(\n) 분리 규칙 적용]
              const shiftCellVal = String(shiftRow[col] || '').trim();
              const lines = shiftCellVal
                .split(/\r?\n/)
                .map(l => l.trim())
                .filter(Boolean)
                .filter(l => l.replace(/[\s\-_=]/g, '').length > 0);

              const shifts: ShiftItem[] = [];

              // [요구사항 1]: M 근무 분리 헬퍼 (M1과 M2가 위아래 2개 줄로 예쁘게 분리되도록 배정)
              const pushMidShifts = (lineText: string) => {
                const workers = cleanWorkers(lineText);
                if (workers.length >= 2) {
                  // 근무자가 2명 이상이면 M1과 M2로 분리하여 2개 줄 렌더링 지원!
                  const m1Workers = [workers[0]];
                  const m2Workers = workers.slice(1);
                  shifts.push({
                    code: 'M1',
                    name: '미드1',
                    time: '14:00 - 24:00',
                    workers: m1Workers,
                    hasTargetUser: m1Workers.some(w => w.includes('현우')),
                  });
                  shifts.push({
                    code: 'M2',
                    name: '미드2',
                    time: '14:00 - 24:00',
                    workers: m2Workers,
                    hasTargetUser: m2Workers.some(w => w.includes('현우')),
                  });
                } else if (workers.length === 1) {
                  shifts.push({
                    code: 'M1',
                    name: '미드',
                    time: '14:00 - 24:00',
                    workers,
                    hasTargetUser: workers.some(w => w.includes('현우')),
                  });
                }
              };

              // [근무조 배정 알고리즘]
              if (lines.length >= 3) {
                // 3줄: [0]=D, [1]=M1/M2, [2]=N
                if (lines[0]) {
                  const workers = cleanWorkers(lines[0]);
                  if (workers.length > 0) {
                    shifts.push({
                      code: 'D',
                      name: '데이',
                      time: '08:00 - 15:00',
                      workers,
                      hasTargetUser: workers.some(w => w.includes('현우')),
                    });
                  }
                }
                if (lines[1]) {
                  pushMidShifts(lines[1]);
                }
                if (lines[2]) {
                  const workers = cleanWorkers(lines[2]);
                  if (workers.length > 0) {
                    shifts.push({
                      code: 'N',
                      name: '나이트',
                      time: '00:00 - 익일 08:00',
                      workers,
                      hasTargetUser: workers.some(w => w.includes('현우')),
                    });
                  }
                }
              } else if (lines.length === 2) {
                // 2줄: D는 없음(투명 빈칸)! [0]=M1/M2, [1]=N
                if (lines[0]) {
                  pushMidShifts(lines[0]);
                }
                if (lines[1]) {
                  const workers = cleanWorkers(lines[1]);
                  if (workers.length > 0) {
                    shifts.push({
                      code: 'N',
                      name: '나이트',
                      time: '00:00 - 익일 08:00',
                      workers,
                      hasTargetUser: workers.some(w => w.includes('현우')),
                    });
                  }
                }
              } else if (lines.length === 1) {
                // 1줄: M 위치에 배정
                pushMidShifts(lines[0]);
              }

              if (!processedDateKeys.has(dateStr)) {
                processedDateKeys.add(dateStr);
                parsedData.push({
                  date: dateStr,
                  yearMonth,
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

            // 다음 탐색 행은 근무자 행 다음으로 이동 (빈 행이 있더라도 유연하게 스킵)
            nextR = Math.max(nextR + 1, shiftRowIndex + 1);
          } else {
            // 날짜가 없는 빈 행이나 비고 행은 1행씩 전진
            nextR++;
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

  // [요구사항 1]: 근무 수동 수정 상태
  const [isEditingDay, setIsEditingDay] = useState<boolean>(false);
  const [editHolidayNote, setEditHolidayNote] = useState<string>('');
  const [editShifts, setEditShifts] = useState<{
    code: ShiftCode;
    name: string;
    time: string;
    workersStr: string;
  }[]>([]);
  const [editSaveSuccess, setEditSaveSuccess] = useState<boolean>(false);

  // [요구사항 2]: 오늘 날짜(YYYY-MM-DD) 추출 - 1년치 스크롤에서도 오늘 위치를 즉각 식별
  const todayStr = useMemo(() => {
    const now = new Date();
    return normalizeDateKey(now.getFullYear(), now.getMonth() + 1, now.getDate());
  }, []);

  // [요구사항 1]: 오직 기기 네이티브 달력 선택기(input type="month") 단일화 강제
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
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        const d = parseInt(parts[2], 10);
        if (!isNaN(y) && !isNaN(m) && !isNaN(d)) {
          const normKey = normalizeDateKey(y, m, d);
          map.set(normKey, day);

          // 비패딩 키(예: 2026-9-1)도 등록
          map.set(`${y}-${m}-${d}`, day);

          // 월만 패딩 키(예: 2026-09-1)
          map.set(`${y}-${String(m).padStart(2, '0')}-${d}`, day);

          // 일만 패딩 키(예: 2026-9-01)
          map.set(`${y}-${m}-${String(d).padStart(2, '0')}`, day);
        }
      }

      // 3) yearMonth + dayNum 조합 키 등록
      if (day.yearMonth && day.dayNum) {
        map.set(`${day.yearMonth}-${String(day.dayNum).padStart(2, '0')}`, day);
        map.set(`${day.yearMonth}-${day.dayNum}`, day);
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
  // [요구사항 2]: Now 클릭 시 무조건 현재 실제 날짜가 속한 이번 달로 즉시 이동
  const handleGoToCurrentMonth = () => {
    const today = new Date();
    const curY = today.getFullYear();
    const curM = today.getMonth() + 1;
    setCurrentYear(curY);
    setCurrentMonth(curM);
    if (viewMode === '1year') {
      setTimeout(() => {
        const el = document.getElementById(`month-block-${curY}-${curM}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      }, 50);
    } else {
      setViewMode('1month');
    }
  };

  // [핵심 1]: 1Y 연속 스크롤 모드로 전환할 때 화면 튐 없이 현재 보고 있던 달을 화면에 안정적으로 유지
  const handleSwitchToYearView = () => {
    const targetY = currentYear;
    const targetM = currentMonth;
    setViewMode('1year');

    // 렌더링 후 현재 보고 있던 월 위치로 화면 튐 없이 자연스럽게 안착
    setTimeout(() => {
      const el = document.getElementById(`month-block-${targetY}-${targetM}`);
      if (el) {
        el.scrollIntoView({ behavior: 'auto', block: 'start' });
      }
    }, 20);
  };

  const handleSwitchToMonthView = () => {
    setViewMode('1month');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  // 월 / 연도 이동 (1M 모드일 땐 1달씩 상태 변경, 1Y 연속 스크롤 모드일 땐 이전/다음 달로 부드러운 스크롤 이동)
  const handlePrevMonth = () => {
    if (viewMode === '1year') {
      const prevM = currentMonth > 1 ? currentMonth - 1 : 12;
      const prevY = currentMonth > 1 ? currentYear : currentYear - 1;
      if (prevY !== currentYear) {
        setCurrentYear(prevY);
      }
      setCurrentMonth(prevM);
      const el = document.getElementById(`month-block-${prevY}-${prevM}`);
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    } else {
      const prev = new Date(currentYear, currentMonth - 1 - 1, 1);
      setCurrentYear(prev.getFullYear());
      setCurrentMonth(prev.getMonth() + 1);
    }
  };

  const handleNextMonth = () => {
    if (viewMode === '1year') {
      const nextM = currentMonth < 12 ? currentMonth + 1 : 1;
      const nextY = currentMonth < 12 ? currentYear : currentYear + 1;
      if (nextY !== currentYear) {
        setCurrentYear(nextY);
      }
      setCurrentMonth(nextM);
      const el = document.getElementById(`month-block-${nextY}-${nextM}`);
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    } else {
      const next = new Date(currentYear, currentMonth - 1 + 1, 1);
      setCurrentYear(next.getFullYear());
      setCurrentMonth(next.getMonth() + 1);
    }
  };

  // [연속 스크롤 뷰 지원]: 1Y 모드에서 스크롤을 내릴 때 현재 뷰포트에 보이는 월을 실시간 감지하여 상단 헤더 및 통계 동기화
  useEffect(() => {
    if (viewMode !== '1year') return;

    let timeoutId: NodeJS.Timeout;
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting && entry.intersectionRatio >= 0.35) {
            clearTimeout(timeoutId);
            timeoutId = setTimeout(() => {
              const id = entry.target.id;
              const match = id.match(/month-block-(\d{4})-(\d{1,2})/);
              if (match) {
                const y = parseInt(match[1], 10);
                const m = parseInt(match[2], 10);
                setCurrentYear(y);
                setCurrentMonth(m);
              }
            }, 80);
          }
        });
      },
      { threshold: [0.35, 0.6] }
    );

    const monthElements = document.querySelectorAll('[id^="month-block-"]');
    monthElements.forEach((el) => observer.observe(el));

    return () => {
      clearTimeout(timeoutId);
      observer.disconnect();
    };
  }, [viewMode, currentYear]);

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

  // [수동 수정 헬퍼]: 근무조 기본 메타 정보
  const SHIFT_META_MAP: Record<ShiftCode, { name: string; defaultTime: string }> = {
    D: { name: '데이', defaultTime: '08:00 - 15:00' },
    M1: { name: '미드1', defaultTime: '14:00 - 24:00' },
    M2: { name: '미드2', defaultTime: '14:00 - 24:00' },
    M: { name: '미드', defaultTime: '14:00 - 24:00' },
    H: { name: '주말 헬퍼', defaultTime: '14:00 - 23:00' },
    N: { name: '나이트', defaultTime: '00:00 - 익일 07:30' },
  };

  // [요구사항 1]: 근무 수동 수정 모드 진입
  const handleStartEditDay = () => {
    if (!selectedDay) return;
    setEditHolidayNote(selectedDay.holidayNote || '');
    if (selectedDay.shifts && selectedDay.shifts.length > 0) {
      setEditShifts(
        selectedDay.shifts.map(s => ({
          code: s.code,
          name: s.name,
          time: s.time,
          workersStr: s.workers.join(', '),
        }))
      );
    } else {
      // 근무조가 아예 없던 날인 경우 기본 D, M1, M2, N 템플릿 제공
      setEditShifts([
        { code: 'D', name: '데이', time: '08:00 - 15:00', workersStr: '' },
        { code: 'M1', name: '미드1', time: '14:00 - 24:00', workersStr: '' },
        { code: 'M2', name: '미드2', time: '14:00 - 24:00', workersStr: '' },
        { code: 'N', name: '나이트', time: '00:00 - 익일 07:30', workersStr: '' },
      ]);
    }
    setIsEditingDay(true);
  };

  // 근무조 추가
  const handleAddShiftSlot = () => {
    const defaultCode: ShiftCode = 'D';
    const meta = SHIFT_META_MAP[defaultCode];
    setEditShifts(prev => [
      ...prev,
      {
        code: defaultCode,
        name: meta.name,
        time: meta.defaultTime,
        workersStr: '',
      },
    ]);
  };

  // 근무조 삭제
  const handleRemoveShiftSlot = (idx: number) => {
    setEditShifts(prev => prev.filter((_, i) => i !== idx));
  };

  // 근무조 코드 변경 시 시간 및 이름 자동 동기화
  const handleShiftCodeChange = (idx: number, newCode: ShiftCode) => {
    const meta = SHIFT_META_MAP[newCode] || { name: newCode, defaultTime: '' };
    setEditShifts(prev =>
      prev.map((s, i) =>
        i === idx
          ? {
              ...s,
              code: newCode,
              name: meta.name,
              time: meta.defaultTime || s.time,
            }
          : s
      )
    );
  };

  // 근무자 입력 텍스트 변경
  const handleWorkersStrChange = (idx: number, val: string) => {
    setEditShifts(prev =>
      prev.map((s, i) => (i === idx ? { ...s, workersStr: val } : s))
    );
  };

  // 근무자 입력창에 "현우" 원클릭 토글 (있으면 제거, 없으면 추가)
  const handleToggleHyunwoo = (idx: number) => {
    setEditShifts(prev =>
      prev.map((s, i) => {
        if (i !== idx) return s;
        const currentWorkers = s.workersStr
          .split(/[/,\s+&|]+/)
          .map(w => w.trim())
          .filter(Boolean)
          .map(w => (w === '박현우' ? '현우' : w));

        const hasHyunwoo = currentWorkers.includes('현우');
        let newWorkers: string[];
        if (hasHyunwoo) {
          newWorkers = currentWorkers.filter(w => w !== '현우');
        } else {
          newWorkers = [...currentWorkers, '현우'];
        }
        return {
          ...s,
          workersStr: newWorkers.join(', '),
        };
      })
    );
  };

  // [핵심 1]: 근무 수정 저장 및 LocalStorage 자동 영구 보존
  const handleSaveDayEdit = () => {
    if (!selectedDay) return;

    // 근무자 파싱 및 정규화
    const updatedShifts: ShiftItem[] = editShifts
      .filter(s => s.workersStr.trim().length > 0 || s.code)
      .map(s => {
        const workers = s.workersStr
          .split(/[/,\s+&|]+/)
          .map(w => w.trim())
          .filter(Boolean)
          .map(w => (w === '박현우' ? '현우' : w));

        const hasTargetUser = workers.some(w => w.includes('현우'));
        const meta = SHIFT_META_MAP[s.code] || { name: s.name, defaultTime: s.time };

        return {
          code: s.code,
          name: meta.name,
          time: s.time || meta.defaultTime,
          workers,
          hasTargetUser,
        };
      });

    const hasTargetUser = updatedShifts.some(s => s.hasTargetUser);
    const holidayNote = editHolidayNote.trim() || null;

    const updatedDay: ParsedDay = {
      ...selectedDay,
      holidayNote,
      shifts: updatedShifts,
      hasTargetUser,
      rawText: updatedShifts.map(s => `${s.code} ${s.workers.join('/')}`).join('\n'),
    };

    // scheduleList 내 해당 날짜 업데이트
    const updatedList = scheduleList.map(item => {
      if (item.date === selectedDay.date) {
        return updatedDay;
      }
      return item;
    });

    if (!scheduleList.some(item => item.date === selectedDay.date)) {
      updatedList.push(updatedDay);
    }

    setScheduleList(updatedList);

    // LocalStorage 영구 저장 (새로고침 시에도 완벽 유지)
    try {
      localStorage.setItem(STORAGE_DATA_KEY, JSON.stringify(updatedList));
    } catch (err) {
      console.error('LocalStorage 저장 오류:', err);
    }

    setSelectedDay(updatedDay);
    setIsEditingDay(false);
    setEditSaveSuccess(true);
    setTimeout(() => setEditSaveSuccess(false), 2000);
  };

  // [요구사항 1]: 연속 스크롤 월간 뷰 - 1Y 모드일 때 해당 연도의 1월~12월을 세로로 차곡차곡 연속 렌더링
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
        {/* 📌 [요구사항 1]: 상단 컨트롤 바 스티키(Sticky) 고정 컨테이너 */}
        {/* 헤더(ER Schedule), 날짜("2026년 9월"), 타임프레임 탭(Now, 1M, 1Y), 업로드/저장 버튼 일체형 고정 */}
        {/* ========================================================= */}
        <div className="sticky top-0 z-50 bg-zinc-950/90 backdrop-blur-md border-b border-zinc-800/80 shadow-md transition-all">
          {/* 1. 상단 심플 헤더: [ER Schedule] 로고 및 버튼 */}
          <header className="pt-safe px-4 pt-3 pb-2 flex items-center justify-between border-b border-zinc-900/80">
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-1.5">
              ER Schedule
            </h1>

            <div className="flex items-center gap-2">
              {/* [요구사항 1]: 수동 저장 버튼 */}
              <button
                onClick={handleManualSave}
                className={`flex items-center gap-1.5 text-xs font-bold px-3 py-1.5 rounded-xl border transition-all shadow-sm active:scale-95 ${
                  saveSuccess
                    ? 'bg-emerald-950/80 border-emerald-500 text-emerald-300'
                    : 'bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border-zinc-700'
                }`}
                title="현재 스케줄 데이터 브라우저에 영구 저장"
              >
                {saveSuccess ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-400" />
                    <span>저장됨</span>
                  </>
                ) : (
                  <>
                    <Save className="w-3.5 h-3.5 text-yellow-400" />
                    <span>저장</span>
                  </>
                )}
              </button>

              {/* 엑셀 업로드 버튼 */}
              <button
                onClick={() => setIsUploadOpen(true)}
                className="flex items-center gap-1.5 bg-zinc-900 hover:bg-zinc-800 active:scale-95 text-zinc-100 text-xs font-bold px-3 py-1.5 rounded-xl border border-zinc-700 transition-all shadow-sm"
              >
                <Upload className="w-3.5 h-3.5 text-yellow-400" />
                <span>엑셀 업로드</span>
              </button>
            </div>
          </header>

          {/* 2. 달력 컨트롤: [YYYY년 M월] 확대 및 클릭 시 네이티브 Date Picker 열기, [Now / 1M / 1Y] 탭 */}
          <div className="px-4 py-2 space-y-1.5">
            <div className="flex items-center justify-between">
              {/* 이전/다음 달 이동 및 대형 [YYYY년 M월] 기기 네이티브 Date Picker 단일화 */}
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={handlePrevMonth}
                  aria-label="이전"
                  className="p-1.5 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border border-zinc-800 transition-colors active:scale-95"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>

                {/* [요구사항 1]: 오직 기기 네이티브 달력 선택 창만 뜨도록 강제 (웹 커스텀 모달 완전 배제, 투명 input overlay) */}
                <label
                  className="relative text-lg sm:text-xl font-bold text-white hover:text-yellow-400 active:scale-95 transition-all px-2 py-1 rounded-xl hover:bg-zinc-900 flex items-center gap-1 tracking-tight group cursor-pointer"
                  title="기기 네이티브 달력으로 날짜 선택"
                >
                  <span>{currentYear}년 {currentMonth}월</span>
                  <ChevronDown className="w-4 h-4 text-zinc-400 group-hover:text-yellow-400 transition-colors" />

                  {/* 텍스트 영역 전체를 덮어 터치 시 100% 아이폰/스마트폰 고유 네이티브 스크롤 피커 실행 */}
                  <input
                    type="month"
                    ref={monthInputRef}
                    value={`${currentYear}-${String(currentMonth).padStart(2, '0')}`}
                    onChange={(e) => {
                      if (e.target.value) {
                        const [y, m] = e.target.value.split('-').map(Number);
                        if (y && m) {
                          setCurrentYear(y);
                          setCurrentMonth(m);
                          if (viewMode === '1year') {
                            setTimeout(() => {
                              const el = document.getElementById(`month-block-${y}-${m}`);
                              if (el) {
                                el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                              }
                            }, 50);
                          }
                        }
                      }
                    }}
                    className="absolute inset-0 opacity-0 w-full h-full cursor-pointer z-10"
                  />
                </label>

                <button
                  type="button"
                  onClick={handleNextMonth}
                  aria-label="다음"
                  className="p-1.5 rounded-xl bg-zinc-900 hover:bg-zinc-800 text-zinc-300 border border-zinc-800 transition-colors active:scale-95"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>

              {/* [요구사항 2]: 타임프레임 필터 (Now / 1M / 1Y) */}
              <div className="flex items-center bg-zinc-900 p-1 rounded-xl border border-zinc-800 shadow-inner">
                <button
                  type="button"
                  onClick={handleGoToCurrentMonth}
                  className="text-xs font-bold px-2.5 py-1 rounded-lg text-zinc-400 hover:text-yellow-400 active:scale-95 transition-all"
                  title="현재 실제 날짜의 이번 달로 즉시 이동"
                >
                  Now
                </button>
                <button
                  type="button"
                  onClick={handleSwitchToMonthView}
                  className={`text-xs font-bold px-2.5 py-1 rounded-lg transition-all active:scale-95 ${
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
                  onClick={handleSwitchToYearView}
                  className={`text-xs font-bold px-2.5 py-1 rounded-lg transition-all active:scale-95 ${
                    viewMode === '1year'
                      ? 'bg-zinc-800 text-yellow-400 shadow-sm'
                      : 'text-zinc-400 hover:text-zinc-200'
                  }`}
                  title="연간 연속 스크롤 달력 보기"
                >
                  1Y
                </button>
              </div>
            </div>

            {/* 내 근무 요약 바 및 [현우만 보기] 토글 */}
            <div className="flex items-center justify-between text-2xs pt-0.5 text-zinc-400">
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
        </div>

        {/* ========================================================= */}
        {/* 3. 달력 그리드 영역 (연속 스크롤 월간 뷰 지원) */}
        {/* ========================================================= */}
        <main className="flex-1 p-2.5 pt-2 space-y-6 pb-safe">
          {monthsToRender.map(({ year, month }) => {
            const ymStr = `${year}-${String(month).padStart(2, '0')}`;
            const firstDayOfWeek = new Date(year, month - 1, 1).getDay();
            const daysInMonth = new Date(year, month, 0).getDate();
            const isCurrentViewingMonth = year === currentYear && month === currentMonth;

            return (
              <div 
                key={ymStr} 
                id={`month-block-${year}-${month}`} 
                className="space-y-1 scroll-mt-36 transition-all"
              >
                {/* 1Y 연속 스크롤 모드일 때 세련된 Sticky 월 구분 헤더 */}
                {viewMode === '1year' && (
                  <div className="sticky top-[124px] z-20 bg-zinc-950/95 backdrop-blur-md flex items-center justify-between px-2.5 py-1.5 border-b border-zinc-800/80 mb-1.5 rounded-lg shadow-sm">
                    <span className={`text-xs sm:text-sm font-black flex items-center gap-1.5 ${isCurrentViewingMonth ? 'text-yellow-400' : 'text-zinc-200'}`}>
                      <CalendarIcon className="w-4 h-4 text-yellow-400" />
                      <span>{year}년 {month}월</span>
                    </span>
                    <span className="text-2xs text-zinc-500 font-semibold bg-zinc-900 px-2 py-0.5 rounded-full border border-zinc-800">
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
                    const dayData = scheduleMap.get(dateStr) 
                      || scheduleMap.get(`${ymStr}-${String(d).padStart(2, '0')}`) 
                      || scheduleMap.get(`${ymStr}-${d}`) 
                      || scheduleMap.get(`${year}-${month}-${d}`) 
                      || scheduleMap.get(`${year}-${month}-${String(d).padStart(2, '0')}`);
                    const dateObj = new Date(year, month - 1, d);
                    const dayOfWeek = dateObj.getDay();
                    const isSun = dayOfWeek === 0;
                    const isSat = dayOfWeek === 6;
                    const isHol = Boolean(dayData?.holidayNote);

                    // [요구사항 2]: 오늘 날짜 판별
                    const isToday = dateStr === todayStr;

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

                    // [요구사항 2]: 오늘 날짜 쨍하고 선명한 초록색 테두리 하이라이트 적용
                    let cellHighlightClass = 'bg-zinc-900 border-zinc-800 hover:border-zinc-700';

                    if (isToday) {
                      // 오늘 날짜인 경우: 선명하고 쨍한 초록색 테두리 (border-2 border-emerald-500 및 ring-2 ring-emerald-400)
                      // 현우 근무일의 노란색(#fde047) 배경과 겹치지 않고 외곽에 아름다운 네온 초록 테두리 형성
                      cellHighlightClass = 'bg-zinc-900 border-2 border-emerald-500 ring-2 ring-emerald-400 shadow-[0_0_12px_rgba(52,211,153,0.4)] z-10';
                    } else if (hasTargetUser) {
                      cellHighlightClass = 'bg-zinc-900 border-yellow-400 ring-1 ring-yellow-400/60 shadow-md target-highlight-box';
                    }

                    if (isDimmed && !isToday) {
                      cellHighlightClass = 'opacity-25 bg-zinc-950 border-zinc-900';
                    }

                    return (
                      <div
                        key={dateStr}
                        onClick={() => dayData && setSelectedDay(dayData)}
                        className={`min-h-[104px] p-1 rounded-lg border flex flex-col justify-between transition-all cursor-pointer ${cellHighlightClass}`}
                      >
                        {/* 상단 날짜 숫자 및 공휴일 메모 */}
                        <div className="flex items-center justify-between gap-0.5 mb-1 leading-none w-full">
                          <div className="flex items-center gap-1">
                            <span
                              className={`text-xs font-black ${
                                isToday
                                  ? 'text-emerald-400 font-extrabold'
                                  : isSun || isHol
                                  ? 'text-rose-500'
                                  : isSat
                                  ? 'text-sky-400'
                                  : 'text-zinc-200'
                              }`}
                            >
                              {d}
                            </span>
                            {isToday && (
                              <span className="text-[7.5px] font-black text-emerald-950 bg-emerald-400 px-1 py-0.2 rounded-xs leading-none shadow-xs uppercase">
                                오늘
                              </span>
                            )}
                          </div>

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
        {/* 4. 상세 모달 (Dialog) - 수동 근무 수정 & "현우" 포함 시 노란색 Override */}
        {/* ========================================================= */}
        {selectedDay && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs animate-in fade-in duration-150">
            <div 
              className="absolute inset-0" 
              onClick={() => {
                setSelectedDay(null);
                setIsEditingDay(false);
              }} 
            />

            <div className="relative w-full max-w-sm bg-zinc-900 border border-zinc-800 rounded-2xl shadow-2xl p-5 z-10 animate-in zoom-in-95 duration-150 text-zinc-100 max-h-[88vh] flex flex-col">
              {/* 상단 모달 헤더 */}
              <div className="flex items-center justify-between pb-3 border-b border-zinc-800 shrink-0">
                <div className="flex items-center gap-1.5 flex-wrap">
                  <span className="text-base font-black text-white">
                    {selectedDay.date}
                  </span>
                  {selectedDay.date === todayStr && (
                    <span className="text-[10px] font-black text-emerald-950 bg-emerald-400 px-2 py-0.5 rounded-full shadow-xs">
                      오늘
                    </span>
                  )}
                  {!isEditingDay && selectedDay.holidayNote && (
                    <span className="text-xs font-bold text-rose-400 bg-rose-950/80 px-2 py-0.5 rounded-full border border-rose-900">
                      {selectedDay.holidayNote}
                    </span>
                  )}
                  {isEditingDay && (
                    <span className="text-[10px] font-extrabold text-yellow-400 bg-yellow-400/10 border border-yellow-400/40 px-2 py-0.5 rounded-md">
                      수정 모드
                    </span>
                  )}
                </div>

                <div className="flex items-center gap-1.5">
                  {!isEditingDay ? (
                    <button
                      type="button"
                      onClick={handleStartEditDay}
                      className="flex items-center gap-1 text-2xs font-bold px-2.5 py-1 rounded-lg bg-zinc-800 hover:bg-zinc-700 text-yellow-400 border border-zinc-700 transition-all active:scale-95 shadow-xs"
                      title="근무 수동 수정"
                    >
                      <Edit2 className="w-3.5 h-3.5" />
                      <span>수정</span>
                    </button>
                  ) : (
                    <button
                      type="button"
                      onClick={() => setIsEditingDay(false)}
                      className="text-2xs font-bold px-2 py-1 rounded-lg bg-zinc-800 hover:bg-zinc-700 text-zinc-400 transition-colors"
                    >
                      취소
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={() => {
                      setSelectedDay(null);
                      setIsEditingDay(false);
                    }}
                    className="p-1 rounded-lg text-zinc-400 hover:text-white bg-zinc-800 hover:bg-zinc-700 transition-colors"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

              {/* 저장 성공 피드백 배너 */}
              {editSaveSuccess && (
                <div className="mt-2.5 p-2 bg-emerald-950/80 border border-emerald-500/80 rounded-xl flex items-center justify-center gap-1.5 text-emerald-300 text-xs font-extrabold animate-in fade-in shrink-0">
                  <Check className="w-4 h-4 text-emerald-400" />
                  <span>근무 수정 사항이 영구 저장되었습니다!</span>
                </div>
              )}

              {/* [분기 1]: ✏️ 근무 수동 수정(편집) 폼 */}
              {isEditingDay ? (
                <div className="mt-3.5 space-y-3.5 overflow-y-auto pr-1 flex-1">
                  {/* 공휴일 / 메모 입력란 */}
                  <div>
                    <label className="text-2xs font-bold text-zinc-400 block mb-1">
                      공휴일 / 특이사항 메모
                    </label>
                    <input
                      type="text"
                      value={editHolidayNote}
                      onChange={(e) => setEditHolidayNote(e.target.value)}
                      placeholder="예: 추석연휴, 대체공휴일, 당직교체 등"
                      className="w-full bg-zinc-950 border border-zinc-700 rounded-lg px-3 py-1.5 text-xs text-white placeholder-zinc-600 focus:outline-none focus:border-yellow-400"
                    />
                  </div>

                  {/* 각 근무조 편집 리스트 */}
                  <div className="space-y-2.5">
                    <div className="flex items-center justify-between">
                      <span className="text-2xs font-bold text-zinc-400">근무조 편성</span>
                      <button
                        type="button"
                        onClick={handleAddShiftSlot}
                        className="flex items-center gap-1 text-2xs font-bold text-yellow-400 hover:text-yellow-300 bg-yellow-400/10 hover:bg-yellow-400/20 px-2 py-0.5 rounded-lg border border-yellow-400/40 transition-colors"
                      >
                        <Plus className="w-3 h-3" />
                        <span>근무조 추가</span>
                      </button>
                    </div>

                    {editShifts.map((shift, idx) => {
                      const hasHyunwoo = shift.workersStr.includes('현우');

                      return (
                        <div
                          key={idx}
                          className={`p-3 rounded-xl border transition-all ${
                            hasHyunwoo
                              ? 'bg-zinc-950 border-yellow-400/80 shadow-xs'
                              : 'bg-zinc-950/70 border-zinc-800'
                          }`}
                        >
                          <div className="flex items-center justify-between gap-2 mb-2">
                            {/* 근무 코드 select 드롭다운 & 시간 input */}
                            <div className="flex items-center gap-1.5 flex-1">
                              <select
                                value={shift.code}
                                onChange={(e) => handleShiftCodeChange(idx, e.target.value as ShiftCode)}
                                className="bg-zinc-900 border border-zinc-700 text-white text-xs font-bold rounded-lg px-2 py-1 focus:outline-none focus:border-yellow-400 cursor-pointer"
                              >
                                <option value="D">D (데이)</option>
                                <option value="M1">M1 (미드1)</option>
                                <option value="M2">M2 (미드2)</option>
                                <option value="M">M (미드)</option>
                                <option value="H">H (헬퍼)</option>
                                <option value="N">N (나이트)</option>
                              </select>

                              <input
                                type="text"
                                value={shift.time}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setEditShifts(prev => prev.map((s, i) => i === idx ? { ...s, time: val } : s));
                                }}
                                placeholder="08:00 - 15:00"
                                className="bg-zinc-900 border border-zinc-800 rounded-md px-2 py-1 text-2xs text-zinc-400 flex-1 focus:outline-none focus:border-yellow-400"
                              />
                            </div>

                            {/* 근무조 삭제 버튼 */}
                            <button
                              type="button"
                              onClick={() => handleRemoveShiftSlot(idx)}
                              className="p-1 rounded text-zinc-500 hover:text-rose-400 hover:bg-zinc-900 transition-colors shrink-0"
                              title="근무조 삭제"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>

                          {/* 근무자 이름 입력창 */}
                          <div className="space-y-1.5">
                            <input
                              type="text"
                              value={shift.workersStr}
                              onChange={(e) => handleWorkersStrChange(idx, e.target.value)}
                              placeholder="근무자 이름 입력 (예: 현우, 민준)"
                              className="w-full bg-zinc-900 border border-zinc-700 rounded-lg px-3 py-1.5 text-xs text-white placeholder-zinc-600 focus:outline-none focus:border-yellow-400 font-medium"
                            />

                            <div className="flex items-center justify-between gap-1">
                              <span className="text-[10px] text-zinc-500">
                                쉼표(,) 또는 슬래시(/) 구분
                              </span>

                              {/* [대타 지원]: 현우 원클릭 토글 퀵 버튼 */}
                              <button
                                type="button"
                                onClick={() => handleToggleHyunwoo(idx)}
                                className={`flex items-center gap-1 text-[10px] font-extrabold px-2 py-0.5 rounded-full transition-all border shrink-0 ${
                                  hasHyunwoo
                                    ? 'bg-yellow-400 text-black border-yellow-400 shadow-xs'
                                    : 'bg-zinc-900 hover:bg-zinc-800 text-yellow-400 border-zinc-700'
                                }`}
                                title="현우 포함 여부 원클릭 변경"
                              >
                                <Sparkles className="w-2.5 h-2.5" />
                                <span>{hasHyunwoo ? '✓ 현우 포함됨 (제거)' : '+ 현우 대타 추가'}</span>
                              </button>
                            </div>
                          </div>
                        </div>
                      );
                    })}

                    {editShifts.length === 0 && (
                      <div className="p-4 rounded-xl border border-dashed border-zinc-800 text-center text-zinc-500 text-xs">
                        등록된 근무조가 없습니다. [근무조 추가]를 눌러 편성하세요.
                      </div>
                    )}
                  </div>

                  {/* 편집 모드 하단 저장 및 취소 버튼 바 */}
                  <div className="pt-2 flex gap-2 shrink-0">
                    <button
                      type="button"
                      onClick={() => setIsEditingDay(false)}
                      className="px-4 py-2.5 rounded-xl bg-zinc-800 hover:bg-zinc-700 text-zinc-300 text-xs font-bold transition-colors"
                    >
                      취소
                    </button>
                    <button
                      type="button"
                      onClick={handleSaveDayEdit}
                      className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl bg-yellow-400 hover:bg-yellow-300 active:scale-98 text-black font-extrabold text-xs transition-all shadow-md"
                    >
                      <Save className="w-4 h-4" />
                      <span>수정 사항 저장</span>
                    </button>
                  </div>
                </div>
              ) : (
                /* [분기 2]: 👁️ 기존 뷰 모드 */
                <div className="mt-3.5 space-y-2 overflow-y-auto pr-1 flex-1">
                  {selectedDay.hasTargetUser && (
                    <div className="p-2.5 bg-yellow-400/10 border border-yellow-400/40 rounded-xl flex items-center gap-2 text-yellow-400">
                      <Sparkles className="w-4 h-4 fill-yellow-400 shrink-0" />
                      <span className="text-xs font-extrabold">
                        현우 당직 근무일입니다.
                      </span>
                    </div>
                  )}

                  {/* 근무조별 섹션 리스트 */}
                  <div className="space-y-2">
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

                    {(!selectedDay.shifts || selectedDay.shifts.length === 0) && (
                      <div className="p-4 rounded-xl border border-dashed border-zinc-800 text-center text-zinc-500 text-xs">
                        등록된 근무가 없습니다. [수정] 버튼을 눌러 근무를 입력하세요.
                      </div>
                    )}
                  </div>

                  {/* 뷰 모드 하단 액션 버튼 바 */}
                  <div className="mt-4 pt-3 border-t border-zinc-800 flex gap-2 shrink-0">
                    <button
                      type="button"
                      onClick={handleStartEditDay}
                      className="flex items-center justify-center gap-1.5 px-3 py-2.5 rounded-xl bg-zinc-800 hover:bg-zinc-700 text-yellow-400 font-bold text-xs border border-zinc-700 transition-all active:scale-98 shadow-sm"
                      title="근무 수동 수정"
                    >
                      <Edit2 className="w-3.5 h-3.5" />
                      <span>수정</span>
                    </button>
                    <button
                      type="button"
                      onClick={handleCopySchedule}
                      className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl bg-yellow-400 hover:bg-yellow-300 active:scale-98 text-black font-extrabold text-xs transition-all shadow-md"
                    >
                      {copied ? <Check className="w-3.5 h-3.5" /> : <Share2 className="w-3.5 h-3.5" />}
                      <span>{copied ? '복사 완료!' : '일정 복사'}</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setSelectedDay(null);
                        setIsEditingDay(false);
                      }}
                      className="px-4 py-2.5 rounded-xl bg-zinc-800 hover:bg-zinc-700 text-zinc-300 text-xs font-bold transition-colors"
                    >
                      닫기
                    </button>
                  </div>
                </div>
              )}
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
