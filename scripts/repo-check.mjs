import fs from "node:fs";
import path from "node:path";

const root = process.cwd();
const required = [
  "app/page.tsx",
  "app/layout.tsx",
  "app/globals.css",
  "components/Converter.tsx",
  "lib/converter.ts",
  "lib/color.ts",
  "lib/file-utils.ts",
  "package.json",
  "next.config.ts",
  "vercel.json",
  "public/icon.svg",
  "public/manifest.webmanifest",
  "README.md",
  "DEPLOY.md",
];

for (const file of required) {
  if (!fs.existsSync(path.join(root, file))) {
    console.error(`Missing required file: ${file}`);
    process.exit(1);
  }
}

const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
for (const dep of ["next", "react", "react-dom", "pdfjs-dist", "tesseract.js", "docx", "lucide-react"]) {
  if (!pkg.dependencies?.[dep]) {
    console.error(`Missing dependency: ${dep}`);
    process.exit(1);
  }
}

console.log("PX Paper repo structure OK");
