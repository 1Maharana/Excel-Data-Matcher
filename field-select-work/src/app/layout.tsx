import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel Data Matcher",
  description: "A local-first spreadsheet comparison and duplicate review tool."
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
