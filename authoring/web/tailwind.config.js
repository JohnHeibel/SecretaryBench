/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        ink: "#0e1525",
        panel: "#161d2e",
        panel2: "#1d2740",
        edge: "#2a3552",
        muted: "#8a97b8",
        accent: "#6aa6ff",
        anchor: "#c084fc",
        ok: "#34d399",
        warn: "#fbbf24",
        bad: "#fb7185",
      },
      fontFamily: {
        mono: ["ui-monospace", "SFMono-Regular", "Menlo", "monospace"],
      },
    },
  },
  plugins: [],
};
