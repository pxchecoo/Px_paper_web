"use client";

import {
  ArrowRight,
  Check,
  ChevronDown,
  Download,
  ExternalLink,
  FileImage,
  FileText,
  Image as ImageIcon,
  Languages,
  LockKeyhole,
  Palette,
  RotateCcw,
  ScanText,
  ShieldCheck,
  Sparkles,
  Upload,
  WandSparkles,
  X,
  Zap,
} from "lucide-react";
import { ChangeEvent, DragEvent, useEffect, useMemo, useRef, useState } from "react";
import { convertToDocx } from "@/lib/converter";
import { downloadBlob, formatBytes, isAcceptedFile, MAX_FILE_SIZE } from "@/lib/file-utils";
import type { ConversionResult, ConversionSettings, ProgressUpdate } from "@/lib/types";

const defaultSettings: ConversionSettings = {
  mode: "smart",
  language: "eng+spa",
  preserveColor: true,
  preserveLayout: true,
};

const modes = [
  {
    id: "smart" as const,
    title: "Smart Edit",
    subtitle: "Best balance",
    description: "Keeps editable text, layout and colors whenever possible.",
    icon: WandSparkles,
  },
  {
    id: "clean" as const,
    title: "Clean Text",
    subtitle: "Maximum editability",
    description: "Prioritizes clean paragraphs over exact placement.",
    icon: ScanText,
  },
  {
    id: "pixel" as const,
    title: "Pixel Clone",
    subtitle: "Maximum visual fidelity",
    description: "Adds a visual clone plus an editable OCR transcript.",
    icon: ImageIcon,
  },
];

export default function Converter() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [dragging, setDragging] = useState(false);
  const [settings, setSettings] = useState(defaultSettings);
  const [progress, setProgress] = useState<ProgressUpdate>({ value: 0, phase: "Ready", detail: "Drop a PDF or photo to begin." });
  const [working, setWorking] = useState(false);
  const [result, setResult] = useState<ConversionResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [history, setHistory] = useState<string[]>([]);

  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem("px-paper-history") ?? "[]");
      if (Array.isArray(saved)) setHistory(saved.slice(0, 4));
    } catch {}
  }, []);

  useEffect(() => {
    if (!file) {
      setPreviewUrl(null);
      return;
    }
    const url = URL.createObjectURL(file);
    setPreviewUrl(url);
    return () => URL.revokeObjectURL(url);
  }, [file]);

  useEffect(() => {
    const handler = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "o") {
        event.preventDefault();
        inputRef.current?.click();
      }
      if ((event.metaKey || event.ctrlKey) && event.key === "Enter" && file && !working) {
        event.preventDefault();
        void startConversion();
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  });

  const selectedMode = useMemo(() => modes.find((mode) => mode.id === settings.mode) ?? modes[0], [settings.mode]);

  function acceptFile(nextFile?: File) {
    if (!nextFile) return;
    setError(null);
    setResult(null);
    if (!isAcceptedFile(nextFile)) {
      setError("Use a PDF, PNG, JPG/JPEG, WEBP or HEIC file.");
      return;
    }
    if (nextFile.size > MAX_FILE_SIZE) {
      setError("That file is over 35 MB. Compress it a little and try again.");
      return;
    }
    setFile(nextFile);
    setProgress({ value: 0, phase: "Ready", detail: "File loaded. Choose a mode and convert." });
  }

  function onInput(event: ChangeEvent<HTMLInputElement>) {
    acceptFile(event.target.files?.[0]);
    event.target.value = "";
  }

  function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragging(false);
    acceptFile(event.dataTransfer.files?.[0]);
  }

  async function startConversion() {
    if (!file || working) return;
    setWorking(true);
    setError(null);
    setResult(null);
    try {
      const converted = await convertToDocx(file, settings, setProgress);
      setResult(converted);
      const nextHistory = [file.name, ...history.filter((item) => item !== file.name)].slice(0, 4);
      setHistory(nextHistory);
      localStorage.setItem("px-paper-history", JSON.stringify(nextHistory));
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : "Conversion failed. Try another file or mode.");
      setProgress({ value: 0, phase: "Conversion stopped", detail: "Nothing was uploaded or stored." });
    } finally {
      setWorking(false);
    }
  }

  function reset() {
    setFile(null);
    setResult(null);
    setError(null);
    setProgress({ value: 0, phase: "Ready", detail: "Drop a PDF or photo to begin." });
  }

  function openGoogleDocs() {
    if (!result) return;
    downloadBlob(result.blob, result.fileName);
    window.open("https://drive.google.com/drive/u/0/my-drive", "_blank", "noopener,noreferrer");
  }

  return (
    <main className="site-shell">
      <div className="aurora aurora-one" />
      <div className="aurora aurora-two" />
      <div className="grid-noise" />

      <nav className="topbar container">
        <a className="brand" href="#top" aria-label="PX Paper home">
          <span className="brand-mark">PX</span>
          <span className="brand-copy"><b>PX Paper</b><small>by pxcheco</small></span>
        </a>
        <div className="nav-actions">
          <span className="privacy-pill"><ShieldCheck size={15} /> 100% local</span>
          <span className="px-signature">PX / by pxcheco</span>
        </div>
      </nav>

      <section id="top" className="hero container">
        <div className="eyebrow"><Sparkles size={15} /> PX intelligence, zero uploads</div>
        <h1>Paper in.<br /><span>Editable doc out.</span></h1>
        <p>Convert PDFs, scans and photos into editable Word files — with color-aware OCR, layout reconstruction and Google Docs-ready output.</p>
        <div className="hero-badges">
          <span><LockKeyhole size={16} /> Files stay on your device</span>
          <span><Palette size={16} /> Color-aware</span>
          <span><Languages size={16} /> English + Español</span>
        </div>
      </section>

      <section className="workspace container">
        <div className="workspace-heading">
          <div><span className="step-tag">01</span><h2>Convert a document</h2></div>
          <span className="shortcut">⌘ O to upload · ⌘ ↵ to convert</span>
        </div>

        <div className="workspace-grid">
          <div className="left-stack">
            {!file ? (
              <div
                className={`dropzone ${dragging ? "dragging" : ""}`}
                onDragEnter={(e) => { e.preventDefault(); setDragging(true); }}
                onDragOver={(e) => e.preventDefault()}
                onDragLeave={() => setDragging(false)}
                onDrop={onDrop}
                onClick={() => inputRef.current?.click()}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") inputRef.current?.click(); }}
              >
                <input ref={inputRef} type="file" hidden accept=".pdf,.png,.jpg,.jpeg,.webp,.heic,.heif" onChange={onInput} />
                <div className="upload-orb"><Upload size={30} /></div>
                <h3>Drop your paper here</h3>
                <p>PDF, PNG, JPG, WEBP or HEIC · up to 35 MB</p>
                <button className="ghost-button" type="button">Choose file <ArrowRight size={16} /></button>
                <span className="local-note"><LockKeyhole size={13} /> It never leaves your browser.</span>
              </div>
            ) : (
              <div className="file-card glass-card">
                <div className="file-toolbar">
                  <div className="file-meta">
                    <span className="file-icon">{file.type === "application/pdf" ? <FileText size={21} /> : <FileImage size={21} />}</span>
                    <span><b>{file.name}</b><small>{formatBytes(file.size)} · {file.type || "document"}</small></span>
                  </div>
                  <button className="icon-button subtle" onClick={reset} aria-label="Remove file"><X size={18} /></button>
                </div>
                <div className="preview-frame">
                  {previewUrl && file.type.startsWith("image/") ? (
                    // eslint-disable-next-line @next/next/no-img-element
                    <img src={previewUrl} alt="Uploaded document preview" />
                  ) : previewUrl ? (
                    <object data={previewUrl} type="application/pdf"><div className="pdf-fallback"><FileText size={42}/><b>PDF loaded</b><span>Preview is hidden by this browser.</span></div></object>
                  ) : null}
                  <span className="preview-label">SOURCE PREVIEW</span>
                </div>
              </div>
            )}

            {history.length > 0 && !file && (
              <div className="history glass-card">
                <span className="tiny-label">RECENT ON THIS DEVICE</span>
                {history.map((item) => <div className="history-row" key={item}><Check size={14}/><span>{item}</span></div>)}
              </div>
            )}
          </div>

          <aside className="control-panel glass-card">
            <div className="panel-title"><span className="step-tag">02</span><div><h3>Conversion style</h3><p>Choose what matters most.</p></div></div>
            <div className="mode-list">
              {modes.map((mode) => {
                const Icon = mode.icon;
                const active = settings.mode === mode.id;
                return (
                  <button key={mode.id} className={`mode-card ${active ? "active" : ""}`} onClick={() => setSettings((s) => ({ ...s, mode: mode.id }))} type="button">
                    <span className="mode-icon"><Icon size={19}/></span>
                    <span className="mode-copy"><b>{mode.title}</b><small>{mode.subtitle}</small><em>{mode.description}</em></span>
                    <span className="radio-dot">{active && <span/>}</span>
                  </button>
                );
              })}
            </div>

            <div className="settings-divider" />
            <label className="select-label"><span><Languages size={16}/> OCR language</span><div className="select-wrap"><select value={settings.language} onChange={(e) => setSettings((s) => ({ ...s, language: e.target.value as ConversionSettings["language"] }))}><option value="eng+spa">English + Español</option><option value="eng">English</option><option value="spa">Español</option></select><ChevronDown size={15}/></div></label>
            <label className="toggle-row"><span><Palette size={16}/><span><b>Preserve color</b><small>Estimate original text colors.</small></span></span><input type="checkbox" checked={settings.preserveColor} onChange={(e) => setSettings((s) => ({ ...s, preserveColor: e.target.checked }))}/><i/></label>
            <label className="toggle-row"><span><ScanText size={16}/><span><b>Preserve layout</b><small>Keep approximate position and alignment.</small></span></span><input type="checkbox" checked={settings.preserveLayout} onChange={(e) => setSettings((s) => ({ ...s, preserveLayout: e.target.checked }))}/><i/></label>

            <button className="convert-button" disabled={!file || working} onClick={() => void startConversion()} type="button">
              {working ? <><span className="spinner"/> {progress.phase}</> : <><Zap size={18}/> Convert with {selectedMode.title}</>}
            </button>
            <p className="button-note">Processing happens locally in your browser.</p>
          </aside>
        </div>

        {(file || working || result || error) && (
          <div className={`status-card glass-card ${result ? "success" : ""} ${error ? "error" : ""}`}>
            <div className="status-top">
              <div className="status-icon">{result ? <Check size={20}/> : error ? <X size={20}/> : <Sparkles size={20}/>}</div>
              <div className="status-copy"><b>{result ? "Document ready" : error ? "Something went wrong" : progress.phase}</b><span>{result ? `${result.pages} page${result.pages === 1 ? "" : "s"} · ${result.usedOcr ? "OCR + DOCX" : "native PDF text + DOCX"}` : error ?? progress.detail}</span></div>
              {!result && !error && <strong>{progress.value}%</strong>}
            </div>
            {!result && !error && <div className="progress-track"><span style={{ width: `${progress.value}%` }}/></div>}
            {result && (
              <div className="result-actions">
                <button className="primary-result" onClick={() => downloadBlob(result.blob, result.fileName)}><Download size={17}/> Download .docx</button>
                <button className="secondary-result" onClick={openGoogleDocs}><ExternalLink size={17}/> Open in Google Docs</button>
                <button className="icon-button subtle" onClick={reset} aria-label="Convert another"><RotateCcw size={17}/></button>
              </div>
            )}
          </div>
        )}
      </section>

      <section className="trust container">
        <div className="trust-card"><LockKeyhole/><b>Private by design</b><p>No file upload endpoint exists. Conversion happens in your browser.</p></div>
        <div className="trust-card"><Palette/><b>Fidelity-first</b><p>Digital PDFs keep native text; scans use color-aware OCR reconstruction.</p></div>
        <div className="trust-card"><ExternalLink/><b>Google Docs-ready</b><p>The result is a standard editable .docx. Download it, then open it in Drive.</p></div>
      </section>

      <footer className="footer container"><span><span className="brand-mark mini">PX</span> PX Paper</span><p>Built with obsession <span>·</span> <b>by pxcheco</b></p><small>Files stay yours.</small></footer>
    </main>
  );
}
