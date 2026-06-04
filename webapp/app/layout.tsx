import type { Metadata } from "next";
import "./globals.css";
import "@xyflow/react/dist/style.css";

export const metadata: Metadata = {
  title: "SecretaryBench — Corpus Authoring",
  description: "Author benchmark emails with a built-in date builder, a structured answer key, and a DAG view — validated against the real grader.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="antialiased">{children}</body>
    </html>
  );
}
