import type { Metadata, Viewport } from "next";
import "./globals.css";

const MEDICAL_CROSS_ICON =
  "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1MTIgNTEyIiB3aWR0aD0iNTEyIiBoZWlnaHQ9IjUxMiI+CiAgPHJlY3Qgd2lkdGg9IjUxMiIgaGVpZ2h0PSI1MTIiIHJ4PSIxMTIiIGZpbGw9IiMwRDBEMEQiLz4KICA8ZyB0cmFuc2Zvcm09InRyYW5zbGF0ZSgxMjEsIDExNikiPgogICAgPCEtLSDsupjrprDrjZQg7Z2w7IOJIOuzuOyytCAtLT4KICAgIDxyZWN0IHg9IjAiIHk9IjE2IiB3aWR0aD0iMjcwIiBoZWlnaHQ9IjI4MCIgcng9IjM2IiBmaWxsPSIjRkZGRkZGIi8+CiAgICA8IS0tIOy6mOumsOuNlCDsg4Hri6gg67mo6rCE7IOJIO2XpOuNlCAtLT4KICAgIDxwYXRoIGQ9Ik0gMCA1MiBDIDAgMzIgMTYgMTYgMzYgMTYgTCAyMzQgMTYgQyAyNTQgMTYgMjcwIDMyIDI3MCA1MiBMIDI3MCA4OCBMIDAgODggWiIgZmlsbD0iI0VFNDM0MyIvPgogICAgPCEtLSDsg4Hri6ggMuqwnOydmCDqsoDsnYDsg4kg67CU7J24642UIOungSAtLT4KICAgIDxyZWN0IHg9IjUyIiB5PSIwIiB3aWR0aD0iMjAiIGhlaWdodD0iMzQiIHJ4PSIxMCIgZmlsbD0iIzBEMEQwRCIvPgogICAgPHJlY3QgeD0iMTk4IiB5PSIwIiB3aWR0aD0iMjAiIGhlaWdodD0iMzQiIHJ4PSIxMCIgZmlsbD0iIzBEMEQwRCIvPgogICAgPCEtLSDsoJXspJHslZkg67mo6rCE7IOJIOyLreyekCDrp4jtgawgLS0+CiAgICA8cmVjdCB4PSI3NSIgeT0iMTYwIiB3aWR0aD0iMTIwIiBoZWlnaHQ9IjQyIiByeD0iMTQiIGZpbGw9IiNFRTQzNDMiLz4KICAgIDxyZWN0IHg9IjExNCIgeT0iMTIxIiB3aWR0aD0iNDIiIGhlaWdodD0iMTIwIiByeD0iMTQiIGZpbGw9IiNFRTQzNDMiLz4KICA8L2c+Cjwvc3ZnPg==";

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
    icon: [
      { url: "/icon-192.png", sizes: "192x192", type: "image/png" },
      { url: "/icon-512.png", sizes: "512x512", type: "image/png" },
      { url: "/icon.svg", type: "image/svg+xml" },
    ],
    apple: [
      { url: "/apple-touch-icon.png", sizes: "180x180", type: "image/png" },
    ],
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
        <link rel="icon" href="/icon.png" type="image/png" />
        <link rel="icon" href="/icon.svg" type="image/svg+xml" />
        <link rel="apple-touch-icon" href="/apple-touch-icon.png" />
        <link rel="apple-touch-icon-precomposed" href="/apple-touch-icon.png" />
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
