import type { Metadata, Viewport } from "next";
import "./globals.css";

const MEDICAL_CROSS_ICON =
  "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1MTIgNTEyIj48cmVjdCB3aWR0aD0iNTEyIiBoZWlnaHQ9IjUxMiIgcng9IjExMiIgZmlsbD0iIzA5MDkwYiIvPjxyZWN0IHg9IjIwOCIgeT0iMTA2IiB3aWR0aD0iOTYiIGhlaWdodD0iMzAwIiByeD0iMjAiIGZpbGw9IiNmYWNjMTUiLz48cmVjdCB4PSIxMDYiIHk9IjIwOCIgd2lkdGg9IjMwMCIgaGVpZ2h0PSI5NiIgcng9IjIwIiBmaWxsPSIjZmFjYzE1Ii8+PC9zdmc+";

export const metadata: Metadata = {
  title: "🏥 ER Schedule",
  description: "응급의학과 개인 스케줄 캘린더",
  manifest: "/manifest.json",
  appleWebApp: {
    capable: true,
    statusBarStyle: "black-translucent",
    title: "ER Schedule",
  },
  icons: {
    icon: MEDICAL_CROSS_ICON,
    apple: MEDICAL_CROSS_ICON,
  },
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 1,
  userScalable: false,
  viewportFit: "cover",
  themeColor: "#09090b",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko" className="h-full dark">
      <head>
        <link rel="icon" href={MEDICAL_CROSS_ICON} type="image/svg+xml" />
        <link rel="apple-touch-icon" href={MEDICAL_CROSS_ICON} />
      </head>
      <body className="h-full antialiased bg-black text-zinc-50 flex justify-center selection:bg-yellow-400 selection:text-black">
        {/* 모바일 화면 뷰포트 컨테이너 (최대 448px) */}
        <div className="min-h-screen bg-zinc-950 text-zinc-50 w-full max-w-md mx-auto shadow-2xl relative">
          {children}
        </div>
      </body>
    </html>
  );
}
