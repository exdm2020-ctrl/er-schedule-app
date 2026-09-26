'use client';

import React, { useState, useRef } from 'react';
import { 
  X, 
  UploadCloud, 
  FileSpreadsheet, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Sparkles,
  Info,
  CalendarDays
} from 'lucide-react';
import { parseExcelSchedule } from '@/lib/excelParser';
import { generateSampleExcelBlob, getInitialSampleSchedule } from '@/lib/sampleData';
import { ScheduleData } from '@/types/schedule';

interface FileUploadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onScheduleLoaded: (data: ScheduleData) => void;
  targetUserName: string;
}

export const FileUploadModal: React.FC<FileUploadModalProps> = ({
  isOpen,
  onClose,
  onScheduleLoaded,
  targetUserName,
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successInfo, setSuccessInfo] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  if (!isOpen) return null;

  // 파일 파싱 처리
  const processFile = async (file: File) => {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      setErrorMessage('.xlsx 또는 .xls 형식의 엑셀 파일만 업로드할 수 있습니다.');
      return;
    }

    setIsLoading(true);
    setErrorMessage(null);
    setSuccessInfo(null);

    try {
      const buffer = await file.arrayBuffer();
      const parsedData = await parseExcelSchedule(buffer, file.name, targetUserName);
      const totalDays = Object.keys(parsedData.days).length;
      setSuccessInfo(`"${file.name}" 파싱 완료! (총 ${totalDays}일치 다중 월 데이터 통합 저장)`);
      
      setTimeout(() => {
        onScheduleLoaded(parsedData);
        onClose();
      }, 700);
    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || '엑셀 파일 파싱 중 오류가 발생했습니다. 규격을 확인해주세요.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  };

  // 다중 월 샘플 엑셀 파일 다운로드
  const handleDownloadSample = () => {
    const blob = generateSampleExcelBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '응급의학과_근무표_2026년9_10_11월_다중월연속_샘플.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // 다중 월 샘플 데이터 즉시 로드
  const handleLoadSampleDirectly = () => {
    setIsLoading(true);
    const sample = getInitialSampleSchedule();
    setTimeout(() => {
      onScheduleLoaded(sample);
      setIsLoading(false);
      onClose();
    }, 400);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="absolute inset-0" onClick={onClose} />

      <div className="relative w-full max-w-md bg-white rounded-3xl shadow-2xl p-5 z-10 animate-in zoom-in-95 duration-200 max-h-[92vh] overflow-y-auto">
        {/* 헤더 */}
        <div className="flex items-center justify-between pb-3 border-b border-slate-100">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-xl bg-blue-100 text-blue-600 flex items-center justify-center shadow-xs">
              <FileSpreadsheet className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-base font-bold text-slate-800">응급의학과 근무표 엑셀 업로드</h3>
              <p className="text-2xs text-slate-400">다중 월(9, 10, 11월...) 연속 파싱 지원</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-full text-slate-400 hover:text-slate-700 bg-slate-100 hover:bg-slate-200 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* 파일 드롭 영역 */}
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className={`mt-4 border-2 border-dashed rounded-2xl p-6 text-center cursor-pointer transition-all ${
            isDragging
              ? 'border-blue-500 bg-blue-50/70 scale-102'
              : 'border-slate-200 hover:border-blue-400 bg-slate-50/50 hover:bg-blue-50/20'
          }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".xlsx,.xls"
            className="hidden"
          />

          <div className="w-12 h-12 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center mx-auto mb-3 shadow-xs">
            <UploadCloud className="w-6 h-6" />
          </div>

          <p className="text-xs font-bold text-slate-800 mb-1">
            클릭하여 파일 선택 또는 드래그 앤 드롭
          </p>
          <p className="text-2xs text-slate-400">
            지원 포맷: .xlsx, .xls (다중 월 연속 시트 완벽 지원)
          </p>

          {isLoading && (
            <div className="mt-3 flex items-center justify-center gap-2 text-xs font-semibold text-blue-600">
              <div className="w-3.5 h-3.5 border-2 border-blue-600 border-t-transparent rounded-full animate-spin" />
              <span>다중 월 스케줄 및 날짜 정규식 파싱 중...</span>
            </div>
          )}

          {successInfo && (
            <div className="mt-3 p-2 bg-emerald-50 text-emerald-700 rounded-xl text-xs font-medium flex items-center justify-center gap-1.5 border border-emerald-200">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>{successInfo}</span>
            </div>
          )}

          {errorMessage && (
            <div className="mt-3 p-2 bg-rose-50 text-rose-700 rounded-xl text-xs font-medium flex items-center justify-center gap-1.5 border border-rose-200">
              <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
              <span className="text-left text-xs">{errorMessage}</span>
            </div>
          )}
        </div>

        {/* 엑셀 파싱 가이드 박스 */}
        <div className="mt-3.5 p-3 bg-slate-50 rounded-2xl border border-slate-200/80 text-2xs space-y-1.5 text-slate-600">
          <div className="flex items-center gap-1 text-slate-800 font-bold">
            <Info className="w-3.5 h-3.5 text-blue-600" />
            <span>엑셀 파싱 핵심 규칙 (무결성 보장)</span>
          </div>
          <ul className="list-disc pl-4 space-y-1 leading-relaxed text-slate-500">
            <li>
              <strong className="text-slate-700">다중 월 순회</strong>: &quot;YYYY.MM월&quot; (예: 2026.09월, 2026.10월) 패턴을 만날 때마다 다음 월로 전환하여 시트 끝까지 모든 월을 통합 저장합니다.
            </li>
            <li>
              <strong className="text-slate-700">날짜 정규식 파싱</strong>: &quot;24(추석연휴)&quot;, &quot;25(추석)&quot; 등 문자가 섞여 있어도 <code className="bg-slate-200 px-1 py-0.2 rounded text-slate-800">\d+</code> 정규식으로 순수한 일(Day) 숫자만 추출하여 에러를 방지합니다.
            </li>
            <li>
              <strong className="text-slate-700">평일 3줄/2줄</strong>: 3줄(데이\n미드1,2\n나이트) 및 2줄(미드1,2\n나이트) 자동 판별
            </li>
            <li>
              <strong className="text-slate-700">주말/공휴일 3줄</strong>: 데이(08-16)\n주말 헬퍼/미드\n나이트(00-08) 매핑
            </li>
          </ul>
        </div>

        {/* 하단 샘플 액션 버튼 */}
        <div className="mt-4 pt-3 border-t border-slate-100 flex flex-col gap-2">
          <button
            onClick={handleLoadSampleDirectly}
            className="w-full flex items-center justify-center gap-1.5 py-2.5 px-3 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold transition-colors"
          >
            <Sparkles className="w-3.5 h-3.5" />
            <span>샘플 데이터 바로 불러오기 (2026년 9, 10, 11월 통합)</span>
          </button>

          <button
            onClick={handleDownloadSample}
            className="w-full flex items-center justify-center gap-1.5 py-2.5 px-3 rounded-xl bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 text-xs font-medium transition-colors"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>다중 월 규격 샘플 엑셀 (.xlsx) 다운로드</span>
          </button>
        </div>
      </div>
    </div>
  );
};
