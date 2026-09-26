import * as XLSX from 'xlsx';
import { ScheduleData } from '@/types/schedule';
import { parseCellShifts, TARGET_USER_NAME } from './excelParser';
import { getHolidayName } from './holidays';

const ER_DOCTORS = [
  '박현우', // Target
  '김민준',
  '이서연',
  '정유진',
  '최준호',
  '윤도윤',
  '강예은',
  '임재현',
  '송지민',
  '한승우',
];

/**
 * 현실적인 응급의학과 샘플 스케줄 텍스트 생성
 */
function generateDayShiftRawText(year: number, month: number, day: number): string {
  const dateObj = new Date(year, month - 1, day);
  const dayOfWeek = dateObj.getDay();
  const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  const isWeekendOrHol = dayOfWeek === 0 || dayOfWeek === 6 || Boolean(getHolidayName(dateStr));

  // 일정한 시드로 현실적인 의사 근무자 로테이션 배정
  const seed = (year * 372 + month * 31 + day) % ER_DOCTORS.length;
  const getDoc = (offset: number) => ER_DOCTORS[(seed + offset) % ER_DOCTORS.length];

  if (isWeekendOrHol) {
    // 주말/공휴일 (3줄 규칙)
    // 1행: 주말 데이 (08:00 - 16:00)
    // 2행: 주말 헬퍼 (14:00 - 23:00) / 주말 미드 (15:00 - 24:00)
    // 3행: 주말 나이트 (00:00 - 익일 08:00)
    const dayDoc = getDoc(0);
    const helperDoc = getDoc(1);
    const midDoc = getDoc(2);
    const nightDoc = getDoc(3);

    return `${dayDoc}\n${helperDoc}/${midDoc}\n${nightDoc}`;
  } else {
    // 평일: 요일별로 3줄 또는 2줄(데이가 없는 날) 배정
    const isNoDayShift = (day % 7 === 3 || day % 7 === 5);

    if (isNoDayShift) {
      // 2줄 규칙 (데이가 없는 날):
      // 1행: 평일 미드1 / 미드2 (14:00 - 24:00)
      // 2행: 평일 나이트 (00:00 - 익일 07:30)
      const mid1Doc = getDoc(1);
      const mid2Doc = getDoc(2);
      const nightDoc = getDoc(4);
      return `${mid1Doc}/${mid2Doc}\n${nightDoc}`;
    } else {
      // 3줄 규칙:
      // 1행: 평일 데이 (08:00 - 15:00)
      // 2행: 평일 미드1 / 미드2 (14:00 - 24:00)
      // 3행: 평일 나이트 (00:00 - 익일 07:30)
      const dayDoc = getDoc(0);
      const mid1Doc = getDoc(1);
      const mid2Doc = getDoc(2);
      const nightDoc = getDoc(3);
      return `${dayDoc}\n${mid1Doc}/${mid2Doc}\n${nightDoc}`;
    }
  }
}

/**
 * 기본 탑재용 ScheduleData 객체 생성 (2026년 9월, 10월, 11월 3개월 통합 데이터)
 */
export function getInitialSampleSchedule(): ScheduleData {
  const months = [
    { year: 2026, month: 9 },
    { year: 2026, month: 10 },
    { year: 2026, month: 11 },
  ];

  const days: ScheduleData['days'] = {};
  const allWorkersSet = new Set<string>();

  for (const { year, month } of months) {
    const daysInMonth = new Date(year, month, 0).getDate();

    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const dateObj = new Date(year, month - 1, d);
      const dayOfWeek = dateObj.getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      let holidayName = getHolidayName(dateStr);

      // 2026년 9월 24~26일 추석 연휴 현실 반영
      if (year === 2026 && month === 9) {
        if (d === 24) holidayName = '추석연휴';
        if (d === 25) holidayName = '추석';
        if (d === 26) holidayName = '추석연휴';
      }

      const rawText = generateDayShiftRawText(year, month, d);
      const shifts = parseCellShifts(rawText, dateStr, dateObj, Boolean(holidayName), TARGET_USER_NAME);
      const targetUserShifts = shifts.filter(s => s.hasTargetUser);
      const hasTargetUser = targetUserShifts.length > 0;

      shifts.forEach(s => s.workers.forEach(w => allWorkersSet.add(w)));

      days[dateStr] = {
        date: dateStr,
        dayOfWeek,
        dayNumber: d,
        isWeekend,
        isHoliday: Boolean(holidayName),
        holidayName,
        rawText,
        shifts,
        hasTargetUser,
        targetUserShifts,
      };
    }
  }

  return {
    title: '2026년 9월, 10월, 11월 응급의학과 근무표 (샘플)',
    fileName: 'ER_Schedule_2026_09_10_11.xlsx',
    yearMonth: '2026-09',
    startDate: '2026-09-01',
    endDate: '2026-11-30',
    days,
    allWorkers: Array.from(allWorkersSet).sort(),
    parsedAt: new Date().toISOString(),
  };
}

/**
 * 엑셀 시트 하나에 9월, 10월, 11월 등 여러 달의 스케줄이 아래로 계속 이어져 있는
 * 실제 병원 근무표 규격의 엑셀 (.xlsx) 파일 생성기
 */
export function generateSampleExcelBlob(): Blob {
  const months = [
    { year: 2026, month: 9 },
    { year: 2026, month: 10 },
    { year: 2026, month: 11 },
  ];

  const rows: (string | number)[][] = [];

  months.forEach(({ year, month }, mIdx) => {
    if (mIdx > 0) {
      // 월 사이에 빈 행 2개 추가
      rows.push([]);
      rows.push([]);
    }

    // 1. 월 구분 헤더: "YYYY.MM월" 패턴 (예: "2026.09월", "2026.10월")
    rows.push([`${year}.${String(month).padStart(2, '0')}월 응급의학과 당직 근무표`]);
    rows.push(['* 셀 내 줄바꿈(\\n) 근무조 구분, 슬래시(/) 복수 근무자 구분']);

    // 2. 요일 행 (일 ~ 토)
    rows.push(['일', '월', '화', '수', '목', '금', '토']);

    // 3. 주 단위로 날짜 행과 근무자 행 번갈아 배치
    const daysInMonth = new Date(year, month, 0).getDate();
    const firstDayOfWeek = new Date(year, month - 1, 1).getDay();

    let currentDay = 1;
    let isFirstWeek = true;

    while (currentDay <= daysInMonth) {
      const dateRow: string[] = [];
      const shiftRow: string[] = [];

      for (let dayOfWeek = 0; dayOfWeek < 7; dayOfWeek++) {
        if (isFirstWeek && dayOfWeek < firstDayOfWeek) {
          dateRow.push('');
          shiftRow.push('');
        } else if (currentDay <= daysInMonth) {
          // 날짜 셀: 추석연휴, 개천절 등 글자가 섞여 있는 경우 재현!
          let dateCellText = `${currentDay}`;
          if (year === 2026 && month === 9) {
            if (currentDay === 24) dateCellText = '24(추석연휴)';
            else if (currentDay === 25) dateCellText = '25(추석)';
            else if (currentDay === 26) dateCellText = '26(추석연휴)';
          } else if (year === 2026 && month === 10) {
            if (currentDay === 3) dateCellText = '3(개천절)';
            else if (currentDay === 9) dateCellText = '9(한글날)';
          }

          dateRow.push(dateCellText);
          shiftRow.push(generateDayShiftRawText(year, month, currentDay));
          currentDay++;
        } else {
          dateRow.push('');
          shiftRow.push('');
        }
      }

      isFirstWeek = false;
      rows.push(dateRow);
      rows.push(shiftRow);
    }
  });

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  worksheet['!cols'] = [
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '당직근무표');

  const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  return new Blob([excelBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
