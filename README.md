# PX Paper — by pxcheco

A private, browser-first PDF/photo → editable `.docx` converter built for **GitHub → Vercel** deployment with no API keys and no backend setup.

## What it does

- PDF, PNG, JPG/JPEG, WEBP and HEIC input
- Digital PDFs: extracts native editable text instead of OCR when possible
- Scans/photos: OCR with English + Spanish support
- Color-aware OCR (estimates the original text color from source pixels)
- Approximate layout/alignment reconstruction
- **Smart Edit**, **Clean Text**, and **Pixel Clone** modes
- Generates standard `.docx` files compatible with Word and Google Docs
- Google Docs button downloads the file and opens Google Drive
- File processing stays inside the browser; the app has no upload endpoint
- Local recent-file history stores names only, never file contents
- Responsive PX-branded interface

## Deploy to Vercel

See **`DEPLOY.md`** for the shortest GitHub → Vercel checklist.

1. Upload this folder to a new GitHub repository.
2. Go to Vercel and choose **Add New → Project**.
3. Import the GitHub repository.
4. Vercel should detect **Next.js** automatically.
5. Click **Deploy**.

**No environment variables are required.**

> Deployment sanity check: make sure Vercel is building the latest commit from the `main` branch rather than re-deploying an older deployment snapshot.

## Run locally

```bash
npm install
npm run dev
```

Then open `http://localhost:3000`.

## Production checks

```bash
npm run check
```

This runs TypeScript checking, ESLint and a production Next.js build.

## How conversion works

### Digital PDF
PX Paper uses PDF.js to read the PDF text layer and reconstruct editable Word paragraphs. It also inspects PDF drawing operations to approximate text colors.

### Scan / photo
PX Paper renders the page in the browser and uses Tesseract.js (WebAssembly) for OCR. The source pixels around each recognized word are sampled to estimate the original text color.

### Pixel Clone mode
For maximum visual fidelity, PX Paper places a rendered copy of the page into Word and adds an editable OCR transcript. This is intentionally different from Smart Edit: an exact visual clone and fully editable layout cannot both be guaranteed for arbitrary documents.

## Google Docs
The generated output is a standard `.docx`. The **Open in Google Docs** button downloads the output and opens your Google Drive. Upload/drop the downloaded file into Drive and open it with Google Docs.

Automatic Google Drive upload is intentionally not included because it would require a Google OAuth client/project and user authorization, which would break the zero-configuration Vercel deployment.

## Browser notes

- First OCR conversion can take longer because the OCR WebAssembly/language data must be cached by the browser.
- OCR accuracy depends on image resolution, lighting, handwriting and document complexity.
- HEIC/HEIF support depends on the browser decoder. If a browser rejects it, export the image as JPG or PNG.
- Recommended: current Safari, Chrome or Edge on desktop/iPad.
- PDF.js worker/decoder assets and Tesseract OCR language assets are fetched from their pinned public distribution sources; your document bytes are still processed locally in the browser.

## Stack

- Next.js 16
- React 19
- TypeScript
- PDF.js
- Tesseract.js
- docx
- Lucide icons

---

**PX Paper — by pxcheco**
