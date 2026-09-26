import * as XLSX from 'xlsx';
import { DaySchedule, ShiftSlot, ScheduleData, UserStats } from '@/types/schedule';
import { getHolidayName, isWeekendOrHoliday } from './holidays';

export const TARGET_USER_NAME = '박현우';

/**
 * 엑셀 셀 내 텍스트 한 줄에서 근무자 이름을 추출하고 정리
 */
function cleanWorkers(line: string): string[] {
  if (!line) return [];
  // 슬래시(/), 쉼표(,), 공백 등으로 구분 가능하도록 처리
  // 요구사항: "슬래시(/)로 같은 시간대 근무자를 구분"
  return line
    .split(/[\/,]/)
    .map(name => name.trim())
    .filter(name => name.length > 0 && !name.includes('근무') && !name.includes('시간'));
}

/**
 * 시간대별 테마 색상 반환
 */
function getShiftTheme(type: string, isTargetUser: boolean) {
  if (isTargetUser) {
    return {
      bg: 'bg-yellow-50',
      border: 'border-yellow-400',
      text: 'text-slate-950 font-black',
      badgeBg: 'bg-[#FDE047] text-slate-950 ring-1 ring-yellow-400 font-bold',
    };
  }

  switch (type) {
    case 'DAY':
      return {
        bg: 'bg-sky-50',
        border: 'border-sky-200',
        text: 'text-sky-800',
        badgeBg: 'bg-sky-100 text-sky-700',
      };
    case 'MID':
    case 'MID1':
    case 'MID2':
      return {
        bg: 'bg-emerald-50',
        border: 'border-emerald-200',
        text: 'text-emerald-800',
        badgeBg: 'bg-emerald-100 text-emerald-700',
      };
    case 'HELPER':
      return {
        bg: 'bg-teal-50',
        border: 'border-teal-200',
        text: 'text-teal-800',
        badgeBg: 'bg-teal-100 text-teal-700',
      };
    case 'NIGHT':
      return {
        bg: 'bg-indigo-50',
        border: 'border-indigo-200',
        text: 'text-indigo-800',
        badgeBg: 'bg-indigo-100 text-indigo-700',
      };
    default:
      return {
        bg: 'bg-slate-50',
        border: 'border-slate-200',
        text: 'text-slate-700',
        badgeBg: 'bg-slate-100 text-slate-700',
      };
  }
}

/**
 * 단일 날짜의 셀 텍스트를 파싱하여 근무 슬롯(ShiftSlot) 배열 반환
 * 
 * [파싱 규칙 엄격 적용]
 * 평일(월~금) 3줄:
 *  - 1번째 줄 = 평일 데이 (08:00 - 15:00) [7시간]
 *  - 2번째 줄 = 평일 미드1 / 미드2 (14:00 - 24:00) [10시간]
 *  - 3번째 줄 = 평일 나이트 (00:00 - 익일 07:30) [7.5시간]
 * 
 * 평일(월~금) 2줄 (데이가 없는 날):
 *  - 1번째 줄 = 평일 미드1 / 미드2 (14:00 - 24:00) [10시간]
 *  - 2번째 줄 = 평일 나이트 (00:00 - 익일 07:30) [7.5시간]
 * 
 * 주말(토, 일) 및 공휴일:
 *  - 기본적으로 3줄 구성:
 *    1번째 줄 = 주말 데이 (08:00 - 16:00) [8시간]
 *    2번째 줄 = 주말 헬퍼 (14:00 - 23:00) / 주말 미드 (15:00 - 24:00) [9시간]
 *    3번째 줄 = 주말 나이트 (00:00 - 익일 08:00) [8시간]
 */
export function parseCellShifts(
  cellText: string,
  dateStr: string,
  dateObj: Date,
  isHolidayExplicit: boolean = false,
  targetName: string = TARGET_USER_NAME
): ShiftSlot[] {
  if (!cellText || typeof cellText !== 'string') return [];

  const lines = cellText
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l.length > 0);

  if (lines.length === 0) return [];

  const dayOfWeek = dateObj.getDay();
  const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
  const isWeekendOrHol = isWeekend || isHolidayExplicit || Boolean(getHolidayName(dateStr));
  const slots: ShiftSlot[] = [];

  if (isWeekendOrHol) {
    // === 주말 및 공휴일 규칙 ===
    // 1번째 줄: 주말 데이 (08:00 - 16:00)
    if (lines[0]) {
      const workers = cleanWorkers(lines[0]);
      const hasTarget = workers.includes(targetName);
      slots.push({
        id: `${dateStr}-weekend-day`,
        type: 'DAY',
        label: '주말 데이',
        shortLabel: '데이',
        timeRange: '08:00 - 16:00',
        hours: 8,
        workers,
        hasTargetUser: hasTarget,
        colorTheme: getShiftTheme('DAY', hasTarget),
      });
    }

    // 2번째 줄: 주말 헬퍼 (14:00 - 23:00) / 주말 미드 (15:00 - 24:00)
    if (lines[1]) {
      const lineWorkers = cleanWorkers(lines[1]);
      if (lineWorkers.length >= 2) {
        const helperWorker = [lineWorkers[0]];
        const midWorker = lineWorkers.slice(1);
        const hasTargetHelper = helperWorker.includes(targetName);
        const hasTargetMid = midWorker.includes(targetName);

        slots.push({
          id: `${dateStr}-weekend-helper`,
          type: 'HELPER',
          label: '주말 헬퍼',
          shortLabel: '헬퍼',
          timeRange: '14:00 - 23:00',
          hours: 9,
          workers: helperWorker,
          hasTargetUser: hasTargetHelper,
          colorTheme: getShiftTheme('HELPER', hasTargetHelper),
        });

        slots.push({
          id: `${dateStr}-weekend-mid`,
          type: 'MID',
          label: '주말 미드',
          shortLabel: '미드',
          timeRange: '15:00 - 24:00',
          hours: 9,
          workers: midWorker,
          hasTargetUser: hasTargetMid,
          colorTheme: getShiftTheme('MID', hasTargetMid),
        });
      } else {
        const hasTarget = lineWorkers.includes(targetName);
        slots.push({
          id: `${dateStr}-weekend-mid-helper`,
          type: 'MID',
          label: '주말 헬퍼/미드',
          shortLabel: '미드',
          timeRange: '14:00 - 24:00',
          hours: 9,
          workers: lineWorkers,
          hasTargetUser: hasTarget,
          colorTheme: getShiftTheme('MID', hasTarget),
        });
      }
    }

    // 3번째 줄: 주말 나이트 (00:00 - 익일 08:00)
    if (lines[2]) {
      const workers = cleanWorkers(lines[2]);
      const hasTarget = workers.includes(targetName);
      slots.push({
        id: `${dateStr}-weekend-night`,
        type: 'NIGHT',
        label: '주말 나이트',
        shortLabel: '나이트',
        timeRange: '00:00 - 익일 08:00',
        hours: 8,
        workers,
        hasTargetUser: hasTarget,
        colorTheme: getShiftTheme('NIGHT', hasTarget),
      });
    }

    // 혹시 4번째 줄 이상이 있는 경우
    for (let i = 3; i < lines.length; i++) {
      const workers = cleanWorkers(lines[i]);
      const hasTarget = workers.includes(targetName);
      slots.push({
        id: `${dateStr}-extra-${i}`,
        type: 'CUSTOM',
        label: `추가근무 ${i + 1}`,
        shortLabel: '기타',
        timeRange: '협의 시간',
        hours: 8,
        workers,
        hasTargetUser: hasTarget,
        colorTheme: getShiftTheme('CUSTOM', hasTarget),
      });
    }
  } else {
    // === 평일(월~금) 규칙 ===
    if (lines.length >= 3) {
      // 3줄인 경우:
      // 1번째 줄 = 평일 데이 (08:00 - 15:00)
      const dayWorkers = cleanWorkers(lines[0]);
      const hasTargetDay = dayWorkers.includes(targetName);
      slots.push({
        id: `${dateStr}-weekday-day`,
        type: 'DAY',
        label: '평일 데이',
        shortLabel: '데이',
        timeRange: '08:00 - 15:00',
        hours: 7,
        workers: dayWorkers,
        hasTargetUser: hasTargetDay,
        colorTheme: getShiftTheme('DAY', hasTargetDay),
      });

      // 2번째 줄 = 평일 미드1 / 미드2 (14:00 - 24:00)
      const midWorkers = cleanWorkers(lines[1]);
      if (midWorkers.length >= 2) {
        const mid1Worker = [midWorkers[0]];
        const mid2Worker = midWorkers.slice(1);
        const hasTargetMid1 = mid1Worker.includes(targetName);
        const hasTargetMid2 = mid2Worker.includes(targetName);

        slots.push({
          id: `${dateStr}-weekday-mid1`,
          type: 'MID1',
          label: '평일 미드1',
          shortLabel: '미드1',
          timeRange: '14:00 - 24:00',
          hours: 10,
          workers: mid1Worker,
          hasTargetUser: hasTargetMid1,
          colorTheme: getShiftTheme('MID1', hasTargetMid1),
        });

        slots.push({
          id: `${dateStr}-weekday-mid2`,
          type: 'MID2',
          label: '평일 미드2',
          shortLabel: '미드2',
          timeRange: '14:00 - 24:00',
          hours: 10,
          workers: mid2Worker,
          hasTargetUser: hasTargetMid2,
          colorTheme: getShiftTheme('MID2', hasTargetMid2),
        });
      } else {
        const hasTargetMid = midWorkers.includes(targetName);
        slots.push({
          id: `${dateStr}-weekday-mid`,
          type: 'MID',
          label: '평일 미드',
          shortLabel: '미드',
          timeRange: '14:00 - 24:00',
          hours: 10,
          workers: midWorkers,
          hasTargetUser: hasTargetMid,
          colorTheme: getShiftTheme('MID', hasTargetMid),
        });
      }

      // 3번째 줄 = 평일 나이트 (00:00 - 익일 07:30)
      const nightWorkers = cleanWorkers(lines[2]);
      const hasTargetNight = nightWorkers.includes(targetName);
      slots.push({
        id: `${dateStr}-weekday-night`,
        type: 'NIGHT',
        label: '평일 나이트',
        shortLabel: '나이트',
        timeRange: '00:00 - 익일 07:30',
        hours: 7.5,
        workers: nightWorkers,
        hasTargetUser: hasTargetNight,
        colorTheme: getShiftTheme('NIGHT', hasTargetNight),
      });

      // 4번째 줄 이상
      for (let i = 3; i < lines.length; i++) {
        const workers = cleanWorkers(lines[i]);
        const hasTarget = workers.includes(targetName);
        slots.push({
          id: `${dateStr}-extra-${i}`,
          type: 'CUSTOM',
          label: `추가근무 ${i + 1}`,
          shortLabel: '기타',
          timeRange: '협의 시간',
          hours: 8,
          workers,
          hasTargetUser: hasTarget,
          colorTheme: getShiftTheme('CUSTOM', hasTarget),
        });
      }
    } else if (lines.length === 2) {
      // 2줄인 경우 (데이가 없는 날):
      // 1번째 줄 = 평일 미드1 / 미드2 (14:00 - 24:00)
      const midWorkers = cleanWorkers(lines[0]);
      if (midWorkers.length >= 2) {
        const mid1Worker = [midWorkers[0]];
        const mid2Worker = midWorkers.slice(1);
        const hasTargetMid1 = mid1Worker.includes(targetName);
        const hasTargetMid2 = mid2Worker.includes(targetName);

        slots.push({
          id: `${dateStr}-weekday-mid1`,
          type: 'MID1',
          label: '평일 미드1 (데이없음)',
          shortLabel: '미드1',
          timeRange: '14:00 - 24:00',
          hours: 10,
          workers: mid1Worker,
          hasTargetUser: hasTargetMid1,
          colorTheme: getShiftTheme('MID1', hasTargetMid1),
        });

        slots.push({
          id: `${dateStr}-weekday-mid2`,
          type: 'MID2',
          label: '평일 미드2 (데이없음)',
          shortLabel: '미드2',
          timeRange: '14:00 - 24:00',
          hours: 10,
          workers: mid2Worker,
          hasTargetUser: hasTargetMid2,
          colorTheme: getShiftTheme('MID2', hasTargetMid2),
        });
      } else {
        const hasTargetMid = midWorkers.includes(targetName);
        slots.push({
          id: `${dateStr}-weekday-mid`,
          type: 'MID',
          label: '평일 미드 (데이없음)',
          shortLabel: '미드',
          timeRange: '14:00 - 24:00',
          hours: 10,
          workers: midWorkers,
          hasTargetUser: hasTargetMid,
          colorTheme: getShiftTheme('MID', hasTargetMid),
        });
      }

      // 2번째 줄 = 평일 나이트 (00:00 - 익일 07:30)
      const nightWorkers = cleanWorkers(lines[1]);
      const hasTargetNight = nightWorkers.includes(targetName);
      slots.push({
        id: `${dateStr}-weekday-night`,
        type: 'NIGHT',
        label: '평일 나이트',
        shortLabel: '나이트',
        timeRange: '00:00 - 익일 07:30',
        hours: 7.5,
        workers: nightWorkers,
        hasTargetUser: hasTargetNight,
        colorTheme: getShiftTheme('NIGHT', hasTargetNight),
      });
    } else if (lines.length === 1) {
      // 1줄만 있는 경우
      const workers = cleanWorkers(lines[0]);
      const hasTarget = workers.includes(targetName);
      slots.push({
        id: `${dateStr}-weekday-single`,
        type: 'MID',
        label: '평일 단일근무',
        shortLabel: '근무',
        timeRange: '14:00 - 24:00',
        hours: 10,
        workers,
        hasTargetUser: hasTarget,
        colorTheme: getShiftTheme('MID', hasTarget),
      });
    }
  }

  return slots;
}

/**
 * 엑셀 셀 텍스트에서 "YYYY.MM월" (예: "2026.09월", "2026.10월") 패턴 매칭
 */
export function matchMonthHeader(text: string): { year: number; month: number } | null {
  if (!text || typeof text !== 'string') return null;
  const clean = text.trim();
  // "2026.09월", "2026.10월", "2026.9월", "2026-09", "2026년 9월", "2026/09"
  const regex = /(20\d{2})[.\-\/년\s]+0?(\d{1,2})월?/;
  const match = clean.match(regex);
  if (match) {
    const y = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    if (y >= 2000 && y <= 2050 && m >= 1 && m <= 12) {
      return { year: y, month: m };
    }
  }
  return null;
}

/**
 * 날짜 셀 파싱:
 * "24(추석연휴)", "25(추석)", "1(신정)", "31일" 처럼 글자와 숫자가 섞여 있는 경우
 * 반드시 정규식 (\d+)을 사용하여 순수한 "숫자"만 추출
 */
export function extractDayNumberAndMemo(cellValue: any): { dayNum: number; memo: string | null } | null {
  if (cellValue === undefined || cellValue === null) return null;
  const text = String(cellValue).trim();
  if (!text) return null;

  // 1. 숫자 추출 (정규식 \d+)
  const match = text.match(/(\d+)/);
  if (!match) return null;

  const dayNum = parseInt(match[1], 10);
  if (dayNum < 1 || dayNum > 31) return null;

  // 2. 괄호 안의 글자 또는 텍스트 추출 (예: "추석연휴", "추석", "신정", "대체공휴일")
  let memo: string | null = null;
  const memoMatch = text.match(/\((.*?)\)/);
  if (memoMatch && memoMatch[1]) {
    memo = memoMatch[1].trim();
  } else {
    // 괄호 없이 글자가 붙어있는 경우 (예: "24추석")
    const wordMatch = text.replace(/\d+/g, '').replace(/일/g, '').trim();
    if (wordMatch.length > 0) {
      memo = wordMatch;
    }
  }

  return { dayNum, memo };
}

/**
 * 엑셀 시트 전체를 순회하며 다중 월(9월, 10월, 11월...) 데이터를 통합 파싱
 * 
 * [요구사항 엄격 적용]
 * 1. 월(Month) 구분 로직:
 *    - 엑셀 데이터를 순회하다가 셀 값에 "YYYY.MM월" (예: "2026.09월", "2026.10월") 패턴의 텍스트가 나타나면
 *      그 시점부터 해당 월의 캘린더 데이터 수집을 시작하는 기준으로 삼음
 *    - 달이 바뀔 때마다 멈추지 말고 데이터를 계속 읽어 모든 월의 데이터를 상태에 통합 저장 (10월, 11월 공란 방지)
 * 2. 날짜 셀 파싱:
 *    - 요일(일~토) 행 바로 아래 행은 날짜 행.
 *    - 날짜 셀에 "24(추석연휴)", "25(추석)"처럼 글자가 섞여 있어도 정규식(\d+)으로 숫자만 정확히 추출
 * 3. 근무 데이터 행:
 *    - 날짜 행 바로 아래 행은 해당 날짜의 근무자 행
 */
export async function parseExcelSchedule(
  fileBuffer: ArrayBuffer,
  fileName: string,
  targetName: string = TARGET_USER_NAME
): Promise<ScheduleData> {
  const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: false });
  if (workbook.SheetNames.length === 0) {
    throw new Error('엑셀 파일에 시트가 존재하지 않습니다.');
  }

  const days: Record<string, DaySchedule> = {};
  const allWorkersSet = new Set<string>();
  const detectedMonths: { year: number; month: number }[] = [];
  let firstDetectedYearMonth: string | null = null;

  // 모든 시트를 순회하거나 첫 번째 시트 탐색
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    // 시트 전체를 2차원 배열로 변환
    const rawGrid: any[][] = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: '',
    });

    if (!rawGrid || rawGrid.length === 0) continue;

    let currentYear: number | null = null;
    let currentMonth: number | null = null;

    // 시트 이름에서 먼저 연월 확인 시도
    const sheetHeader = matchMonthHeader(sheetName);
    if (sheetHeader) {
      currentYear = sheetHeader.year;
      currentMonth = sheetHeader.month;
      if (!firstDetectedYearMonth) {
        firstDetectedYearMonth = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
      }
      detectedMonths.push(sheetHeader);
    }

    // 파일명에서 확인 시도 (초기 백업값)
    const fileHeader = matchMonthHeader(fileName);
    if (fileHeader && !currentYear) {
      currentYear = fileHeader.year;
      currentMonth = fileHeader.month;
      if (!firstDetectedYearMonth) {
        firstDetectedYearMonth = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
      }
    }

    // 행 단위 순회 시작
    for (let r = 0; r < rawGrid.length; r++) {
      const row = rawGrid[r];
      if (!row || row.length === 0) continue;

      // 1. "YYYY.MM월" (예: "2026.09월", "2026.10월") 패턴 감지
      for (let c = 0; c < row.length; c++) {
        const cellText = String(row[c] || '');
        const monthMatch = matchMonthHeader(cellText);
        if (monthMatch) {
          currentYear = monthMatch.year;
          currentMonth = monthMatch.month;
          const ymStr = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
          if (!firstDetectedYearMonth) {
            firstDetectedYearMonth = ymStr;
          }
          if (!detectedMonths.some(m => m.year === currentYear && m.month === currentMonth)) {
            detectedMonths.push({ year: currentYear, month: currentMonth });
          }
          break; // 해당 행에서 월 감지 완료
        }
      }

      // 연월이 아직 감지되지 않았으면 다음 행으로
      if (!currentYear || !currentMonth) continue;

      // 2. 이 행이 "날짜 행"인지 검사
      // 날짜 행의 특징: 행의 여러 셀에서 숫자 1~31 (또는 "24(추석)", "1일")이 발견됨
      // 그리고 바로 다음 행(r + 1)에 근무자 텍스트(줄바꿈 \n 또는 슬래시 / 포함)가 존재함
      let validDayCountInRow = 0;
      const candidateDaysInRow: { colIndex: number; dayNum: number; memo: string | null }[] = [];

      for (let c = 0; c < row.length; c++) {
        const extracted = extractDayNumberAndMemo(row[c]);
        if (extracted) {
          candidateDaysInRow.push({ colIndex: c, ...extracted });
          validDayCountInRow++;
        }
      }

      // 한 행에 1개 이상의 유효 날짜 숫자가 있고,
      // 다음 행(r + 1)이 존재할 때 -> 날짜 행 & 근무자 행 쌍으로 처리!
      if (validDayCountInRow > 0 && r + 1 < rawGrid.length) {
        const nextRow = rawGrid[r + 1];

        // 날짜별 매핑 진행
        let matchedShiftCount = 0;

        for (const item of candidateDaysInRow) {
          const shiftCellVal = String(nextRow[item.colIndex] || '').trim();
          
          // 근무 데이터가 있거나(줄바꿈 등), 최소한 해당 날짜를 기록
          const dateStr = `${currentYear}-${String(currentMonth).padStart(2, '0')}-${String(item.dayNum).padStart(2, '0')}`;
          const dateObj = new Date(currentYear, currentMonth - 1, item.dayNum);
          const dayOfWeek = dateObj.getDay();
          const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

          // 공휴일 여부: 셀 텍스트 메모(예: 추석연휴) 또는 대한민국 공식 공휴일
          const holidayFromDict = getHolidayName(dateStr);
          const holidayName = item.memo || holidayFromDict;
          const isHoliday = Boolean(holidayName);

          // 근무 파싱
          const shifts = parseCellShifts(shiftCellVal, dateStr, dateObj, isHoliday, targetName);
          const targetUserShifts = shifts.filter(s => s.hasTargetUser);
          const hasTargetUser = targetUserShifts.length > 0;

          // 모든 근무자 이름 수집
          shifts.forEach(s => s.workers.forEach(w => allWorkersSet.add(w)));

          days[dateStr] = {
            date: dateStr,
            dayOfWeek,
            dayNumber: item.dayNum,
            isWeekend,
            isHoliday,
            holidayName,
            rawText: shiftCellVal,
            shifts,
            hasTargetUser,
            targetUserShifts,
          };

          if (shiftCellVal) {
            matchedShiftCount++;
          }
        }

        // 만약 다음 행이 근무자 행이었다면 r을 1 증가시켜 다음 루프에서 건너뜀
        if (matchedShiftCount > 0) {
          r++; // r + 1은 근무자 행이므로 건너뜀
        }
      }
    }
  }

  // 감지된 월이 없었던 경우 기본 현재 월 사용
  if (!firstDetectedYearMonth) {
    const now = new Date();
    firstDetectedYearMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    detectedMonths.push({ year: now.getFullYear(), month: now.getMonth() + 1 });
  }

  // 감지된 각 월에 대해 비어있는 날짜를 기본 객체로 채워 완전한 달력 보장
  for (const { year, month } of detectedMonths) {
    const daysInMonth = new Date(year, month, 0).getDate();
    const ymPrefix = `${year}-${String(month).padStart(2, '0')}`;

    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${ymPrefix}-${String(d).padStart(2, '0')}`;
      if (!days[dateStr]) {
        const dateObj = new Date(year, month - 1, d);
        const dayOfWeek = dateObj.getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const holidayName = getHolidayName(dateStr);

        days[dateStr] = {
          date: dateStr,
          dayOfWeek,
          dayNumber: d,
          isWeekend,
          isHoliday: Boolean(holidayName),
          holidayName,
          rawText: '',
          shifts: [],
          hasTargetUser: false,
          targetUserShifts: [],
        };
      }
    }
  }

  const allDates = Object.keys(days).sort();
  const startDate = allDates[0] || `${firstDetectedYearMonth}-01`;
  const endDate = allDates[allDates.length - 1] || `${firstDetectedYearMonth}-28`;

  // 타이틀 생성
  const monthLabels = detectedMonths.map(m => `${m.month}월`).join(', ');
  const title = detectedMonths.length > 1
    ? `${detectedMonths[0].year}년 ${monthLabels} 응급의학과 근무표`
    : `${firstDetectedYearMonth.split('-')[0]}년 ${parseInt(firstDetectedYearMonth.split('-')[1])}월 응급의학과 근무표`;

  return {
    title,
    fileName,
    yearMonth: firstDetectedYearMonth,
    startDate,
    endDate,
    days,
    allWorkers: Array.from(allWorkersSet).sort(),
    parsedAt: new Date().toISOString(),
  };
}

/**
 * 박현우(Target User)의 근무 통계 계산 유틸
 * - 선택된 월(yearMonth) 또는 전체 스케줄 기준 계산 지원
 */
export function calculateUserStats(
  schedule: ScheduleData,
  targetName: string = TARGET_USER_NAME,
  filterYearMonth?: string // 특정 'YYYY-MM'만 필터링하거나 전체
): UserStats {
  let totalDays = 0;
  let totalHours = 0;
  let dayCount = 0;
  let midCount = 0;
  let helperCount = 0;
  let nightCount = 0;
  let weekendHolidayCount = 0;

  Object.values(schedule.days).forEach(day => {
    // 특정 월 필터가 있는 경우
    if (filterYearMonth && !day.date.startsWith(filterYearMonth)) {
      return;
    }

    const myShifts = day.shifts.filter(s => s.workers.includes(targetName));
    if (myShifts.length > 0) {
      totalDays += 1;
      if (day.isWeekend || day.isHoliday) {
        weekendHolidayCount += 1;
      }
      myShifts.forEach(shift => {
        totalHours += shift.hours;
        if (shift.type === 'DAY') dayCount += 1;
        else if (shift.type === 'MID' || shift.type === 'MID1' || shift.type === 'MID2') midCount += 1;
        else if (shift.type === 'HELPER') helperCount += 1;
        else if (shift.type === 'NIGHT') nightCount += 1;
      });
    }
  });

  return {
    targetName,
    totalDays,
    totalHours,
    dayCount,
    midCount,
    helperCount,
    nightCount,
    weekendHolidayCount,
  };
}
