import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "NPI Contact Enrichment",
  description: "Validate licensed U.S. healthcare contacts against NPPES and prepare a CRM-ready enrichment export.",
  icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body className="antialiased">{children}</body></html>;
}
