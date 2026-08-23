# Deploy PX Paper to Vercel

PX Paper is intentionally zero-config: there is no backend, database, API key, or environment variable.

## 1) GitHub

Create a new repository and upload **the contents of this folder** to it.

Recommended repository name: `px-paper`

## 2) Vercel

1. Sign in to Vercel.
2. Click **Add New → Project**.
3. Import your `px-paper` GitHub repository.
4. Keep the detected framework as **Next.js**.
5. Do not add environment variables.
6. Click **Deploy**.

Vercel installs the dependencies from `package.json` and runs `npm run build` automatically.

## 3) Optional custom domain

After the first deploy, open **Project → Settings → Domains** and add your domain or subdomain.

Examples:

- `paper.pxcheco.com`
- `pxpaper.com`
- `convert.pxcheco.com`

## Local development

```bash
npm install
npm run dev
```

Open `http://localhost:3000`.

## Full verification

```bash
npm run check
```

That runs TypeScript, ESLint, and the production Next.js build.
