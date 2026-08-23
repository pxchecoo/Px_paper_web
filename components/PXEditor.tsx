"use client";

import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowLeft,
  Bold,
  Check,
  Download,
  Eraser,
  FileText,
  Italic,
  List as ListIcon,
  ListOrdered,
  Palette,
  Printer,
  Redo2,
  RotateCcw,
  Sparkles,
  Underline,
  Undo2,
} from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import { docxBlobToEditorHtml, editedOutputName, editorHtmlToDocx } from "@/lib/docx-editor";
import { downloadBlob } from "@/lib/file-utils";
import type { ConversionResult } from "@/lib/types";
import styles from "./PXEditor.module.css";

type Props = {
  result: ConversionResult;
  onClose: () => void;
};

export default function PXEditor({ result, onClose }: Props) {
  const editorRef = useRef<HTMLDivElement>(null);
  const originalHtmlRef = useRef("");
  const selectionRef = useRef<Range | null>(null);
  const saveTimerRef = useRef<number | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dirty, setDirty] = useState(false);
  const [saved, setSaved] = useState(false);
  const [words, setWords] = useState(0);
  const [characters, setCharacters] = useState(0);
  const draftKey = useMemo(() => `px-editor-draft:${result.fileName}`, [result.fileName]);

  useEffect(() => {
    let active = true;
    setLoading(true);
    setError(null);
    void (async () => {
      try {
        const generatedHtml = await docxBlobToEditorHtml(result.blob);
        if (!active) return;
        originalHtmlRef.current = generatedHtml;
        let initialHtml = generatedHtml;
        try {
          const draft = localStorage.getItem(draftKey);
          if (draft) initialHtml = draft;
        } catch {}
        if (editorRef.current) editorRef.current.innerHTML = initialHtml;
        updateCounts(initialHtml);
        setLoading(false);
        setDirty(initialHtml !== generatedHtml);
        setSaved(Boolean(initialHtml !== generatedHtml));
        try { document.execCommand("styleWithCSS", false, "true"); } catch {}
      } catch (err) {
        if (!active) return;
        console.error(err);
        setError(err instanceof Error ? err.message : "PX Editor could not open this document.");
        setLoading(false);
      }
    })();
    return () => {
      active = false;
      if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    };
  }, [draftKey, result.blob]);

  useEffect(() => {
    const keyHandler = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        void exportDocx();
      }
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", keyHandler);
    return () => window.removeEventListener("keydown", keyHandler);
  });

  function updateCounts(html?: string) {
    const container = document.createElement("div");
    if (html != null) container.innerHTML = html;
    else container.innerHTML = editorRef.current?.innerHTML ?? "";
    const text = (container.innerText || container.textContent || "").replace(/\s+/g, " ").trim();
    setCharacters(text.length);
    setWords(text ? text.split(" ").filter(Boolean).length : 0);
  }

  function saveSelection() {
    const editor = editorRef.current;
    const selection = window.getSelection();
    if (!editor || !selection || selection.rangeCount === 0) return;
    const range = selection.getRangeAt(0);
    if (editor.contains(range.commonAncestorContainer)) selectionRef.current = range.cloneRange();
  }

  function restoreSelection() {
    const range = selectionRef.current;
    if (!range) return;
    const selection = window.getSelection();
    if (!selection) return;
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function persistDraft() {
    const html = editorRef.current?.innerHTML ?? "";
    setDirty(true);
    setSaved(false);
    updateCounts(html);
    if (saveTimerRef.current) window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(() => {
      try {
        localStorage.setItem(draftKey, html);
        setSaved(true);
      } catch {
        setSaved(false);
      }
    }, 500);
  }

  function exec(command: string, value?: string) {
    restoreSelection();
    editorRef.current?.focus();
    document.execCommand(command, false, value);
    saveSelection();
    persistDraft();
  }

  function toolbarMouseDown(event: React.MouseEvent) {
    event.preventDefault();
  }

  async function exportDocx() {
    if (!editorRef.current || saving) return;
    setSaving(true);
    setError(null);
    try {
      const blob = await editorHtmlToDocx(editorRef.current.innerHTML, result.fileName);
      downloadBlob(blob, editedOutputName(result.fileName));
      setDirty(false);
      setSaved(true);
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : "Could not export the edited document.");
    } finally {
      setSaving(false);
    }
  }

  function restoreOriginal() {
    if (!editorRef.current) return;
    editorRef.current.innerHTML = originalHtmlRef.current || "<p><br></p>";
    try { localStorage.removeItem(draftKey); } catch {}
    setDirty(false);
    setSaved(false);
    updateCounts();
    editorRef.current.focus();
  }

  function printDocument() {
    editorRef.current?.focus();
    window.print();
  }

  return (
    <main className={styles.shell}>
      <header className={styles.topbar}>
        <div className={styles.brand}>
          <button className={styles.back} onClick={onClose} type="button" aria-label="Back to converter">
            <ArrowLeft size={16}/><span>Back</span>
          </button>
          <span className={styles.brandMark}>PX</span>
          <span className={styles.brandCopy}><b>PX Editor</b><small>by pxcheco</small></span>
        </div>
        <div className={styles.fileInfo}>
          <FileText size={16}/>
          <span><b>{result.fileName}</b><small>{result.pages} page{result.pages === 1 ? "" : "s"} · editing locally</small></span>
        </div>
        <div className={styles.actions}>
          <span className={`${styles.saveState} ${saved ? styles.saved : ""}`}>{saved ? <Check size={12}/> : <Sparkles size={12}/>} {saved ? "Saved locally" : dirty ? "Unsaved changes" : "Original"}</span>
          <button className={styles.downloadButton} disabled={saving || loading} onClick={() => void exportDocx()} type="button">
            <Download size={16}/><span>{saving ? "Building…" : "Download DOCX"}</span>
          </button>
        </div>
      </header>

      {!loading && !error && (
        <div className={styles.toolbar}>
          <div className={styles.toolGroup}>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("undo")} title="Undo"><Undo2 size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("redo")} title="Redo"><Redo2 size={16}/></button>
          </div>
          <div className={styles.toolGroup}>
            <label className={styles.selectControl} title="Paragraph style">
              <select defaultValue="p" onChange={(e) => exec("formatBlock", e.target.value)} onMouseDown={saveSelection}>
                <option value="p">Normal</option>
                <option value="h1">Title</option>
                <option value="h2">Heading</option>
                <option value="h3">Subheading</option>
              </select>
            </label>
            <label className={styles.selectControl} title="Font">
              <select defaultValue="Aptos" onChange={(e) => exec("fontName", e.target.value)} onMouseDown={saveSelection}>
                <option value="Aptos">Aptos</option>
                <option value="Arial">Arial</option>
                <option value="Helvetica">Helvetica</option>
                <option value="Times New Roman">Times New Roman</option>
                <option value="Georgia">Georgia</option>
                <option value="Courier New">Courier New</option>
              </select>
            </label>
            <label className={styles.selectControl} title="Font size">
              <select defaultValue="3" onChange={(e) => exec("fontSize", e.target.value)} onMouseDown={saveSelection}>
                <option value="1">8</option>
                <option value="2">10</option>
                <option value="3">12</option>
                <option value="4">14</option>
                <option value="5">18</option>
                <option value="6">24</option>
                <option value="7">36</option>
              </select>
            </label>
          </div>
          <div className={styles.toolGroup}>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("bold")} title="Bold"><Bold size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("italic")} title="Italic"><Italic size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("underline")} title="Underline"><Underline size={16}/></button>
            <label className={styles.colorControl} title="Text color" onMouseDown={saveSelection}>
              <Palette size={16}/><input type="color" defaultValue="#111827" onChange={(e) => exec("foreColor", e.target.value)}/>
            </label>
          </div>
          <div className={styles.toolGroup}>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("justifyLeft")} title="Align left"><AlignLeft size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("justifyCenter")} title="Align center"><AlignCenter size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("justifyRight")} title="Align right"><AlignRight size={16}/></button>
          </div>
          <div className={styles.toolGroup}>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("insertUnorderedList")} title="Bullets"><ListIcon size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("insertOrderedList")} title="Numbered list"><ListOrdered size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={() => exec("removeFormat")} title="Clear formatting"><Eraser size={16}/></button>
          </div>
          <div className={styles.toolGroup}>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={restoreOriginal} title="Restore converted version"><RotateCcw size={16}/></button>
            <button className={styles.toolbarButton} onMouseDown={toolbarMouseDown} onClick={printDocument} title="Print / Save as PDF"><Printer size={16}/></button>
          </div>
        </div>
      )}

      {loading ? (
        <section className={styles.loading}>
          <div><div className={styles.loadingOrb}><Sparkles size={25}/></div><b>Opening in PX Editor</b><p>Reading Word formatting, colors, tables and images…</p></div>
        </section>
      ) : error ? (
        <div className={styles.errorBox}>{error}</div>
      ) : (
        <section className={styles.canvas}>
          <div className={styles.pageWrap}>
            <span className={styles.pageBadge}>PX EDITOR · LOCAL DOCUMENT</span>
            <div
              ref={editorRef}
              className={styles.page}
              contentEditable
              suppressContentEditableWarning
              spellCheck
              onInput={persistDraft}
              onMouseUp={saveSelection}
              onKeyUp={saveSelection}
              onFocus={saveSelection}
            />
          </div>
          <div className={styles.stats}><span>{words.toLocaleString()} words · {characters.toLocaleString()} characters</span><span>⌘S downloads DOCX · Esc returns</span></div>
        </section>
      )}
    </main>
  );
}
