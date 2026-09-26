'use client';

import React, { useState, useMemo, useEffect } from 'react';
import { ViewMode, ScheduleData, DaySchedule } from '@/types/schedule';
import { TARGET_USER_NAME, calculateUserStats } from '@/lib/excelParser';
import { getInitialSampleSchedule } from '@/lib/sampleData';
import { MobileHeader } from '@/components/MobileHeader';
import { SummaryDashboard } from '@/components/SummaryDashboard';
import { CalendarView } from '@/components/CalendarView';
import { ShiftDetailDrawer } from '@/components/ShiftDetailDrawer';
import { FileUploadModal } from '@/components/FileUploadModal';
import { PasscodeLock } from '@/components/PasscodeLock';
import { 
  UploadCloud, 
  Sparkles, 
  Flame,
  Lock
} from 'lucide-react';

const STORAGE_AUTH_KEY = 'er_schedule_auth_token';

export default function HomePage() {
  // 보안 잠금 상태 (나만 볼 수 있게 설정)
  const [isUnlocked, setIsUnlocked] = useState<boolean | null>(null);

  // 기준 날짜 상태 (기본값: 2026년 9월 1일 - 다중 월 샘플 시작일)
  const [currentDate, setCurrentDate] = useState<Date>(new Date(2026, 8, 1));
  const [viewMode, setViewMode] = useState<ViewMode>('month');
  const [onlyMyShifts, setOnlyMyShifts] = useState(false);
  const [selectedDay, setSelectedDay] = useState<DaySchedule | null>(null);
  const [isUploadOpen, setIsUploadOpen] = useState(false);

  // 스케줄 데이터 상태 (2026년 9, 10, 11월 통합 샘플)
  const [scheduleData, setScheduleData] = useState<ScheduleData>(() =>
    getInitialSampleSchedule()
  );

  // 최초 로드 시 기기 인증 상태 확인
  useEffect(() => {
    const savedAuth = localStorage.getItem(STORAGE_AUTH_KEY);
    if (savedAuth) {
      setIsUnlocked(true);
    } else {
      setIsUnlocked(false);
    }
  }, []);

  // 수동 다시 잠금 핸들러
  const handleLockApp = () => {
    localStorage.removeItem(STORAGE_AUTH_KEY);
    setIsUnlocked(false);
  };

  // 현재 보고 있는 월의 YYYY-MM
  const currentYearMonth = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;

  // 박현우 근무 통계 (현재 월 기준 또는 다중 월 종합)
  const userStats = useMemo(() => {
    if (viewMode === 'month' || viewMode === 'week') {
      return calculateUserStats(scheduleData, TARGET_USER_NAME, currentYearMonth);
    }
    return calculateUserStats(scheduleData, TARGET_USER_NAME);
  }, [scheduleData, currentYearMonth, viewMode]);

  // 이전/다음 기간 이동 핸들러
  const handlePrevDate = () => {
    const next = new Date(currentDate);
    if (viewMode === 'week') {
      next.setDate(next.getDate() - 7);
    } else if (viewMode === 'month') {
      next.setMonth(next.getMonth() - 1);
    } else if (viewMode === '3months') {
      next.setMonth(next.getMonth() - 3);
    } else if (viewMode === '6months') {
      next.setMonth(next.getMonth() - 6);
    } else if (viewMode === '1year') {
      next.setFullYear(next.getFullYear() - 1);
    }
    setCurrentDate(next);
  };

  const handleNextDate = () => {
    const next = new Date(currentDate);
    if (viewMode === 'week') {
      next.setDate(next.getDate() + 7);
    } else if (viewMode === 'month') {
      next.setMonth(next.getMonth() + 1);
    } else if (viewMode === '3months') {
      next.setMonth(next.getMonth() + 3);
    } else if (viewMode === '6months') {
      next.setMonth(next.getMonth() + 6);
    } else if (viewMode === '1year') {
      next.setFullYear(next.getFullYear() + 1);
    }
    setCurrentDate(next);
  };

  const handleToday = () => {
    const [y, m] = scheduleData.yearMonth.split('-').map(Number);
    if (y && m) {
      setCurrentDate(new Date(y, m - 1, 1));
    } else {
      setCurrentDate(new Date());
    }
  };

  const handleScheduleLoaded = (newSchedule: ScheduleData) => {
    setScheduleData(newSchedule);
    const [y, m] = newSchedule.yearMonth.split('-').map(Number);
    if (y && m) {
      setCurrentDate(new Date(y, m - 1, 1));
    }
  };

  const handleJumpToMonth = (date: Date) => {
    setCurrentDate(date);
    setViewMode('month');
  };

  // 인증 상태 로딩 중일 때는 빈 화면
  if (isUnlocked === null) {
    return (
      <div className="flex-1 flex items-center justify-center min-h-screen bg-slate-950 text-white">
        <div className="w-6 h-6 border-2 border-amber-400 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  // 잠금 상태인 경우: 나만 볼 수 있는 PIN 잠금화면 렌더링
  if (!isUnlocked) {
    return (
      <PasscodeLock
        targetUserName={TARGET_USER_NAME}
        onUnlock={() => setIsUnlocked(true)}
      />
    );
  }

  return (
    <div className="flex-1 flex flex-col min-h-screen bg-slate-50 relative pb-16">
      {/* 1. 아이폰 최적화 모바일 헤더 */}
      <MobileHeader
        currentDate={currentDate}
        onPrevDate={handlePrevDate}
        onNextDate={handleNextDate}
        onToday={handleToday}
        viewMode={viewMode}
        onViewModeChange={setViewMode}
        onOpenUpload={() => setIsUploadOpen(true)}
        onlyMyShifts={onlyMyShifts}
        onToggleOnlyMyShifts={() => setOnlyMyShifts(!onlyMyShifts)}
        targetUserName={TARGET_USER_NAME}
        hasFileLoaded={true}
        fileName={scheduleData.fileName}
      />

      {/* 2. '박현우' 전용 근무 통계 요약 카드 */}
      <SummaryDashboard stats={userStats} />

      {/* 3. 메인 캘린더 영역 (Month, Week, 3M, 6M, 1Year) */}
      <main className="flex-1">
        <CalendarView
          currentDate={currentDate}
          scheduleData={scheduleData}
          viewMode={viewMode}
          onlyMyShifts={onlyMyShifts}
          targetUserName={TARGET_USER_NAME}
          onSelectDay={(day) => setSelectedDay(day)}
          onJumpToMonth={handleJumpToMonth}
        />
      </main>

      {/* 4. 하단 모바일 고정 바 (아이폰 네이티브 앱 하단 독 스타일) */}
      <div className="fixed bottom-0 left-0 right-0 z-20 flex justify-center pointer-events-none pb-safe">
        <div className="w-full max-w-md px-4 pb-2 pt-1 pointer-events-auto">
          <div className="bg-slate-900/95 backdrop-blur-md text-white rounded-2xl px-3.5 py-2.5 shadow-xl flex items-center justify-between border border-slate-700/60">
            {/* 좌측: 박현우 당직 현황 안내 */}
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-xl bg-[#FDE047] text-slate-950 flex items-center justify-center font-bold text-xs shadow-2xs border border-yellow-400">
                <Sparkles className="w-3.5 h-3.5 fill-slate-950" />
              </div>
              <div>
                <span className="text-2xs text-slate-300 block font-medium">
                  {viewMode === 'month' ? `${currentDate.getMonth() + 1}월 근무` : '총 근무'}
                </span>
                <span className="text-xs font-black text-[#FDE047]">
                  {userStats.totalDays}일 ({userStats.totalHours}시간)
                </span>
              </div>
            </div>

            {/* 우측: 빠른 액션 버튼들 (필터 / 업로드 / 잠금) */}
            <div className="flex items-center gap-1.5">
              <button
                onClick={() => setOnlyMyShifts(!onlyMyShifts)}
                className={`text-2xs font-bold px-2 py-1.5 rounded-xl transition-all flex items-center gap-1 ${
                  onlyMyShifts
                    ? 'bg-[#FDE047] text-slate-950 shadow-sm font-black'
                    : 'bg-slate-800 text-slate-300 hover:text-white'
                }`}
                title="박현우 근무만 보기"
              >
                <Flame className="w-3 h-3" />
                <span>박현우만</span>
              </button>

              <button
                onClick={() => setIsUploadOpen(true)}
                className="bg-blue-600 hover:bg-blue-500 active:scale-95 text-white text-2xs font-bold px-2.5 py-1.5 rounded-xl flex items-center gap-1 transition-all shadow-sm shadow-blue-500/30"
                title="새 엑셀 파일 업로드"
              >
                <UploadCloud className="w-3.5 h-3.5" />
                <span>업로드</span>
              </button>

              {/* 보안 잠금 버튼 */}
              <button
                onClick={handleLockApp}
                className="p-1.5 rounded-xl bg-slate-800 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 transition-colors border border-slate-700/60"
                title="화면 즉시 잠그기"
              >
                <Lock className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* 5. 날짜 클릭 시 나타나는 하단 바텀시트 / 모달 */}
      <ShiftDetailDrawer
        daySchedule={selectedDay}
        onClose={() => setSelectedDay(null)}
        targetUserName={TARGET_USER_NAME}
      />

      {/* 6. 파일 업로드 모달 */}
      <FileUploadModal
        isOpen={isUploadOpen}
        onClose={() => setIsUploadOpen(false)}
        onScheduleLoaded={handleScheduleLoaded}
        targetUserName={TARGET_USER_NAME}
      />
    </div>
  );
}
