import { useEffect, useRef, useState } from "react";
import { isSafeUrl } from "./sanitize";

type RichHtmlEditorProps = {
  value: string;
  placeholders: string[];
  onChange: (value: string) => void;
};

type FormatState = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  unorderedList: boolean;
  orderedList: boolean;
};

const emptyFormat: FormatState = {
  bold: false,
  italic: false,
  underline: false,
  unorderedList: false,
  orderedList: false,
};

function RichHtmlEditor({ value, placeholders, onChange }: RichHtmlEditorProps) {
  const editorRef = useRef<HTMLDivElement | null>(null);
  const savedRangeRef = useRef<Range | null>(null);
  const linkInputRef = useRef<HTMLInputElement | null>(null);
  const [sourceMode, setSourceMode] = useState(false);
  const [linkOpen, setLinkOpen] = useState(false);
  const [linkUrl, setLinkUrl] = useState("");
  const [linkError, setLinkError] = useState<string | null>(null);
  const [format, setFormat] = useState<FormatState>(emptyFormat);

  useEffect(() => {
    const editor = editorRef.current;
    if (!sourceMode && editor && editor.innerHTML !== value) {
      editor.innerHTML = value;
    }
  }, [sourceMode, value]);

  useEffect(() => {
    function refresh() {
      if (document.activeElement !== editorRef.current) return;
      setFormat({
        bold: document.queryCommandState("bold"),
        italic: document.queryCommandState("italic"),
        underline: document.queryCommandState("underline"),
        unorderedList: document.queryCommandState("insertUnorderedList"),
        orderedList: document.queryCommandState("insertOrderedList"),
      });
    }
    document.addEventListener("selectionchange", refresh);
    return () => document.removeEventListener("selectionchange", refresh);
  }, []);

  useEffect(() => {
    if (!linkOpen) return;
    const id = window.requestAnimationFrame(() => linkInputRef.current?.focus());
    return () => window.cancelAnimationFrame(id);
  }, [linkOpen]);

  function emitHtml() {
    onChange(editorRef.current?.innerHTML ?? "");
  }

  function runCommand(command: string, commandValue?: string) {
    editorRef.current?.focus();
    document.execCommand(command, false, commandValue);
    emitHtml();
  }

  function saveSelection() {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      savedRangeRef.current = null;
      return;
    }
    const range = selection.getRangeAt(0);
    const editor = editorRef.current;
    if (editor && editor.contains(range.commonAncestorContainer)) {
      savedRangeRef.current = range.cloneRange();
    } else {
      savedRangeRef.current = null;
    }
  }

  function restoreSelection() {
    const editor = editorRef.current;
    if (!editor) return;
    editor.focus();
    const range = savedRangeRef.current;
    if (!range) return;
    const selection = window.getSelection();
    if (!selection) return;
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function openLinkDialog() {
    saveSelection();
    setLinkUrl("");
    setLinkError(null);
    setLinkOpen(true);
  }

  function closeLinkDialog() {
    setLinkOpen(false);
    setLinkUrl("");
    setLinkError(null);
  }

  function escapeAttribute(value: string) {
    return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;");
  }

  function escapeText(value: string) {
    return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function confirmLink() {
    const trimmed = linkUrl.trim();
    if (!trimmed) {
      setLinkError("L'adresse est requise.");
      return;
    }
    if (!isSafeUrl(trimmed)) {
      setLinkError("URL invalide. Seuls http://, https:// et mailto: sont autorisés.");
      return;
    }
    restoreSelection();
    const range = savedRangeRef.current;
    if (range && !range.collapsed) {
      document.execCommand("createLink", false, trimmed);
    } else {
      const anchor = `<a href="${escapeAttribute(trimmed)}" target="_blank" rel="noopener noreferrer">${escapeText(trimmed)}</a>`;
      document.execCommand("insertHTML", false, anchor);
    }
    emitHtml();
    closeLinkDialog();
  }

  function insertPlaceholder(placeholder: string) {
    runCommand("insertText", placeholder);
  }

  function keepSelection(event: React.MouseEvent<HTMLElement>) {
    event.preventDefault();
  }

  if (sourceMode) {
    return (
      <div className="rich-editor">
        <div className="rich-toolbar">
          <button type="button" className="rich-btn" onClick={() => setSourceMode(false)}>
            Éditeur visuel
          </button>
        </div>
        <textarea
          className="body-editor"
          value={value}
          onChange={(event) => onChange(event.target.value)}
          spellCheck={false}
        />
      </div>
    );
  }

  return (
    <div className="rich-editor">
      <div className="rich-toolbar" role="toolbar" aria-label="Mise en forme">
        <button
          type="button"
          className={`rich-btn rich-btn-bold${format.bold ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand("bold")}
          aria-pressed={format.bold}
          title="Gras (Ctrl+B)"
        >
          G
        </button>
        <button
          type="button"
          className={`rich-btn rich-btn-italic${format.italic ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand("italic")}
          aria-pressed={format.italic}
          title="Italique (Ctrl+I)"
        >
          I
        </button>
        <button
          type="button"
          className={`rich-btn rich-btn-underline${format.underline ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand("underline")}
          aria-pressed={format.underline}
          title="Souligné (Ctrl+U)"
        >
          S
        </button>

        <span className="rich-sep" aria-hidden="true" />

        <button
          type="button"
          className={`rich-btn rich-btn-icon${format.unorderedList ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand("insertUnorderedList")}
          aria-pressed={format.unorderedList}
          title="Liste à puces"
        >
          <svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true">
            <circle cx="3" cy="5" r="1.4" fill="currentColor" />
            <circle cx="3" cy="10" r="1.4" fill="currentColor" />
            <circle cx="3" cy="15" r="1.4" fill="currentColor" />
            <rect x="7" y="4" width="11" height="1.6" rx="0.8" fill="currentColor" />
            <rect x="7" y="9" width="11" height="1.6" rx="0.8" fill="currentColor" />
            <rect x="7" y="14" width="11" height="1.6" rx="0.8" fill="currentColor" />
          </svg>
        </button>
        <button
          type="button"
          className={`rich-btn rich-btn-icon${format.orderedList ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand("insertOrderedList")}
          aria-pressed={format.orderedList}
          title="Liste numérotée"
        >
          <svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true">
            <text x="1" y="7" fontSize="5.5" fontWeight="700" fill="currentColor">1.</text>
            <text x="1" y="13" fontSize="5.5" fontWeight="700" fill="currentColor">2.</text>
            <text x="1" y="19" fontSize="5.5" fontWeight="700" fill="currentColor">3.</text>
            <rect x="7" y="4" width="11" height="1.6" rx="0.8" fill="currentColor" />
            <rect x="7" y="9" width="11" height="1.6" rx="0.8" fill="currentColor" />
            <rect x="7" y="14" width="11" height="1.6" rx="0.8" fill="currentColor" />
          </svg>
        </button>

        <span className="rich-sep" aria-hidden="true" />

        <button
          type="button"
          className="rich-btn"
          onMouseDown={keepSelection}
          onClick={openLinkDialog}
          title="Insérer un lien"
        >
          Lien
        </button>
        <button
          type="button"
          className="rich-btn rich-btn-subtle"
          onMouseDown={keepSelection}
          onClick={() => runCommand("unlink")}
          title="Retirer le lien"
        >
          Retirer
        </button>

        <span className="rich-sep grow" aria-hidden="true" />

        <button type="button" className="rich-btn rich-btn-subtle" onClick={() => setSourceMode(true)} title="Éditer le HTML">
          HTML
        </button>
      </div>

      {linkOpen ? (
        <div className="link-popover">
          <input
            ref={linkInputRef}
            type="url"
            value={linkUrl}
            onChange={(event) => {
              setLinkUrl(event.target.value);
              setLinkError(null);
            }}
            onKeyDown={(event) => {
              if (event.key === "Enter") {
                event.preventDefault();
                confirmLink();
              } else if (event.key === "Escape") {
                event.preventDefault();
                closeLinkDialog();
              }
            }}
            placeholder="https://exemple.fr"
            aria-label="Adresse du lien"
          />
          <button type="button" className="primary compact" onClick={confirmLink}>
            Ajouter
          </button>
          <button type="button" className="ghost compact" onClick={closeLinkDialog}>
            Annuler
          </button>
          {linkError ? <span className="link-error">{linkError}</span> : null}
        </div>
      ) : null}

      {placeholders.length ? (
        <div className="placeholder-toolbar">
          {placeholders.map((placeholder) => (
            <button
              key={placeholder}
              type="button"
              className="placeholder-button"
              onMouseDown={keepSelection}
              onClick={() => insertPlaceholder(placeholder)}
            >
              {placeholder}
            </button>
          ))}
        </div>
      ) : null}

      <div
        ref={editorRef}
        className="rich-surface"
        contentEditable
        role="textbox"
        aria-label="Corps HTML"
        onInput={emitHtml}
        onBlur={emitHtml}
        suppressContentEditableWarning
      />
    </div>
  );
}

export default RichHtmlEditor;
