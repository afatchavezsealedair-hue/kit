import type { Metadata } from "next";
import "./globals.css";
import "./frequency.css";

export const metadata: Metadata = { title: "Kit · Dashboard", description: "Dashboard configurable desde Excel" };
export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="es"><body>{children}</body></html>;
}
