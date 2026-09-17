import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "VOC 월간 보고서 생성기",
  description: "월간 상담 데이터를 정리된 XLSX 보고서로 변환합니다.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body className="antialiased">{children}</body>
    </html>
  );
}
