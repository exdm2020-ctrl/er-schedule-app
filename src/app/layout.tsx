import type { Metadata, Viewport } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "응급의학과 근무표 | 박현우 스케줄",
  description: "병원 응급의학과 당직 근무표 캘린더 (박현우 맞춤)",
  appleWebApp: {
    capable: true,
    statusBarStyle: "default",
    title: "ER 근무표",
  },
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 1,
  userScalable: false,
  viewportFit: "cover",
  themeColor: "#ffffff",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko" className="h-full">
      <body className="h-full antialiased bg-slate-100 text-slate-900 flex justify-center">
        {/* 모바일 화면 뷰포트 컨테이너 (최대 480px, 아이폰 프로 맥스 폭 최적화) */}
        <div className="w-full max-w-md min-h-screen bg-slate-50 relative flex flex-col shadow-2xl ring-1 ring-slate-200/80">
          {children}
        </div>
      </body>
    </html>
  );
}
