/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        ink: "#0b0d11",
        panel: "#14171c",
        panel2: "#1b1f26",
        edge: "#262b34",
        muted: "#888f9c",
        accent: "#5b8cff",
        anchor: "#b98cff",
        fact: "#2dd4bf",
        ok: "#34d399",
        warn: "#f5b544",
        bad: "#f2647a",
      },
      fontFamily: {
        mono: ["ui-monospace", "SFMono-Regular", "Menlo", "monospace"],
      },
      borderRadius: {
        none: "0",
      },
    },
  },
  plugins: [],
};
