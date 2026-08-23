import type { Metadata, Viewport } from "next";
import Script from "next/script";
import "./globals.css";

export const metadata: Metadata = {
  title: "PX Paper — Paper to editable docs",
  description: "Turn PDFs and paper photos into editable DOCX files directly in your browser. Built by pxcheco.",
  applicationName: "PX Paper",
  manifest: "/manifest.webmanifest",
  icons: {
    icon: "/icon.svg",
    apple: "/icon.svg",
  },
  openGraph: {
    title: "PX Paper",
    description: "Paper to editable docs — privately in your browser.",
    type: "website",
  },
};

export const viewport: Viewport = {
  themeColor: "#070914",
  colorScheme: "dark",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>
        {children}
        <Script src="https://accounts.google.com/gsi/client" strategy="afterInteractive" />
      </body>
    </html>
  );
}
