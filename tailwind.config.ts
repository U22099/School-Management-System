import type { Config } from "tailwindcss";

const config: Config = {
  content: [
    "./src/pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/components/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/app/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {
      backgroundImage: {
        "gradient-radial": "radial-gradient(var(--tw-gradient-stops))",
        "gradient-conic":
          "conic-gradient(from 180deg at 50% 50%, var(--tw-gradient-stops))",
      },
      colors: {
        "primary": "var(--primary)",
        "secondary": "var(--secondary)",
        "accent": "var(--accent)",
        "muted": "var(--muted)",
        "sky": "#ceebfa",
        "sky-light": "#edf9fb",
        "purple": "#cfceff",
        "purple-light": "#f1f0ff",
        "yellow": "#fae27c",
        "yellow-light": "#fefce8",
      },
    },
  },
  plugins: [],
};
export default config;
