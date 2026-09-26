'use client';

import React, { useState, useEffect } from 'react';
import { Lock, Unlock, KeyRound, ShieldAlert, Sparkles, CheckCircle2 } from 'lucide-react';

interface PasscodeLockProps {
  onUnlock: () => void;
  targetUserName: string;
}

const DEFAULT_PASSCODE = process.env.NEXT_PUBLIC_APP_PASSWORD || '1234';
const STORAGE_KEY = 'er_schedule_auth_token';

export const PasscodeLock: React.FC<PasscodeLockProps> = ({ onUnlock, targetUserName }) => {
  const [pin, setPin] = useState('');
  const [error, setError] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [rememberMe, setRememberMe] = useState(true);

  // 키패드 입력 핸들러
  const handleNumberClick = (num: string) => {
    if (pin.length < 4) {
      const nextPin = pin + num;
      setPin(nextPin);
      setError(false);
      setErrorMessage('');

      if (nextPin.length === 4) {
        verifyPin(nextPin);
      }
    }
  };

  const handleDelete = () => {
    setPin(prev => prev.slice(0, -1));
    setError(false);
    setErrorMessage('');
  };

  const handleClear = () => {
    setPin('');
    setError(false);
    setErrorMessage('');
  };

  const verifyPin = (inputPin: string) => {
    if (inputPin === DEFAULT_PASSCODE) {
      if (rememberMe) {
        localStorage.setItem(STORAGE_KEY, 'authorized_' + Date.now());
      }
      onUnlock();
    } else {
      setError(true);
      setErrorMessage('비밀번호가 올바르지 않습니다.');
      setTimeout(() => {
        setPin('');
      }, 500);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950 text-white p-4">
      {/* 배경 장식 글로우 */}
      <div className="absolute top-1/4 left-1/2 -translate-x-1/2 w-64 h-64 bg-amber-500/10 rounded-full blur-3xl pointer-events-none" />

      <div className="relative w-full max-w-xs flex flex-col items-center">
        {/* 상단 자물쇠 아이콘 */}
        <div className={`w-16 h-16 rounded-2xl flex items-center justify-center mb-4 transition-all duration-300 ${
          error 
            ? 'bg-rose-500/20 text-rose-400 border border-rose-500/40 animate-shake' 
            : 'bg-amber-400/20 text-amber-400 border border-amber-400/30'
        }`}>
          {error ? <ShieldAlert className="w-8 h-8" /> : <Lock className="w-8 h-8" />}
        </div>

        {/* 타이틀 & 대상자 이름 */}
        <div className="text-center mb-6">
          <div className="flex items-center justify-center gap-1.5 mb-1">
            <span className="text-xs font-bold text-amber-400 bg-amber-400/10 px-2 py-0.5 rounded-full border border-amber-400/20 flex items-center gap-1">
              <Sparkles className="w-3 h-3 fill-amber-400" />
              보안 잠금 모드
            </span>
          </div>
          <h2 className="text-lg font-black tracking-tight text-white">
            {targetUserName} 선생님 전용 스케줄
          </h2>
          <p className="text-xs text-slate-400 mt-1">
            4자리 접근 비밀번호를 입력해주세요
          </p>
        </div>

        {/* PIN 4자리 표시 인디케이터 (아이폰 스타일) */}
        <div className={`flex items-center gap-4 mb-6 transition-transform ${error ? 'animate-bounce' : ''}`}>
          {[0, 1, 2, 3].map(idx => {
            const isFilled = pin.length > idx;
            return (
              <div
                key={idx}
                className={`w-3.5 h-3.5 rounded-full transition-all duration-200 ${
                  isFilled
                    ? 'bg-amber-400 scale-125 shadow-sm shadow-amber-400'
                    : 'bg-slate-700/80 border border-slate-600'
                }`}
              />
            );
          })}
        </div>

        {/* 에러 메시지 또는 안내 문구 */}
        <div className="h-6 mb-4 flex items-center justify-center">
          {errorMessage ? (
            <span className="text-xs font-semibold text-rose-400 animate-in fade-in">
              {errorMessage}
            </span>
          ) : (
            <span className="text-2xs text-slate-500">
              * 기본 비밀번호: <strong className="text-amber-400 font-bold">1234</strong>
            </span>
          )}
        </div>

        {/* 숫자 키패드 (아이폰 스타일 둥근 버튼) */}
        <div className="grid grid-cols-3 gap-3.5 w-full mb-5">
          {['1', '2', '3', '4', '5', '6', '7', '8', '9'].map(num => (
            <button
              key={num}
              onClick={() => handleNumberClick(num)}
              className="w-16 h-16 rounded-full bg-slate-800/80 hover:bg-slate-700 active:bg-amber-400 active:text-slate-950 font-bold text-xl transition-all border border-slate-700/60 mx-auto flex items-center justify-center shadow-xs"
            >
              {num}
            </button>
          ))}
          <button
            onClick={handleClear}
            className="w-16 h-16 rounded-full text-slate-400 hover:text-white font-semibold text-xs mx-auto flex items-center justify-center"
          >
            취소
          </button>
          <button
            onClick={() => handleNumberClick('0')}
            className="w-16 h-16 rounded-full bg-slate-800/80 hover:bg-slate-700 active:bg-amber-400 active:text-slate-950 font-bold text-xl transition-all border border-slate-700/60 mx-auto flex items-center justify-center shadow-xs"
          >
            0
          </button>
          <button
            onClick={handleDelete}
            className="w-16 h-16 rounded-full text-slate-400 hover:text-white font-semibold text-xs mx-auto flex items-center justify-center"
          >
            지우기
          </button>
        </div>

        {/* 이 기기 기억하기 체크박스 */}
        <label className="flex items-center gap-2 cursor-pointer text-xs text-slate-400 hover:text-slate-300">
          <input
            type="checkbox"
            checked={rememberMe}
            onChange={(e) => setRememberMe(e.target.checked)}
            className="rounded border-slate-700 text-amber-400 focus:ring-amber-400"
          />
          <span>이 아이폰에서 로그인 상태 유지</span>
        </label>
      </div>
    </div>
  );
};
