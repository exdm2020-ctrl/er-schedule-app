import type { Metadata, Viewport } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "ER Schedule",
  description: "응급의학과 개인 스케줄 캘린더",
  appleWebApp: {
    capable: true,
    statusBarStyle: "black-translucent",
    title: "ER Schedule",
  },
  icons: {
    icon: "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect width='100' height='100' rx='20' fill='%2309090b'/><path d='M35 50h30M50 35v30' stroke='%23facc15' stroke-width='12' stroke-linecap='round'/></svg>",
    apple: "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect width='100' height='100' rx='22' fill='%2309090b'/><path d='M35 50h30M50 35v30' stroke='%23facc15' stroke-width='12' stroke-linecap='round'/></svg>",
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
        <link
          rel="apple-touch-icon"
          href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect width='100' height='100' rx='22' fill='%2309090b'/><path d='M35 50h30M50 35v30' stroke='%23facc15' stroke-width='12' stroke-linecap='round'/></svg>"
        />
      </head>
      <body className="h-full antialiased bg-black text-zinc-50 flex justify-center selection:bg-yellow-400 selection:text-black">
        {/* 모바일 화면 뷰포트 컨테이너 (최대 448px, 아이폰 최적화) */}
        <div className="min-h-screen bg-zinc-950 text-zinc-50 flex flex-col w-full max-w-md mx-auto shadow-2xl relative">
          {children}
        </div>
      </body>
    </html>
  );
}
