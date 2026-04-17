import { useEffect, useRef, useState } from "react";
import { EditorContent, useEditor } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import Link from "@tiptap/extension-link";
import Underline from "@tiptap/extension-underline";
import Placeholder from "@tiptap/extension-placeholder";
import { TextStyle } from "@tiptap/extension-text-style";
import Color from "@tiptap/extension-color";
import { Table } from "@tiptap/extension-table";
import TableRow from "@tiptap/extension-table-row";
import TableCell from "@tiptap/extension-table-cell";
import TableHeader from "@tiptap/extension-table-header";
import { isSafeUrl, sanitizeEmailHtml } from "./sanitize";

type RichHtmlEditorProps = {
  value: string;
  placeholders: string[];
  onChange: (value: string) => void;
};

type TiptapEditor = NonNullable<ReturnType<typeof useEditor>>;

type FormatState = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  color: string;
  unorderedList: boolean;
  orderedList: boolean;
  link: boolean;
};

const emptyFormat: FormatState = {
  bold: false,
  italic: false,
  underline: false,
  color: "",
  unorderedList: false,
  orderedList: false,
  link: false,
};

const textColors = [
  { label: "Noir", value: "#171b19" },
  { label: "Vert", value: "#0d7c55" },
  { label: "Bleu", value: "#1d4ed8" },
  { label: "Rouge", value: "#b42318" },
  { label: "Gris", value: "#5d6861" },
];

function readFormat(editor: TiptapEditor): FormatState {
  return {
    bold: editor.isActive("bold"),
    italic: editor.isActive("italic"),
    underline: editor.isActive("underline"),
    color: String(editor.getAttributes("textStyle").color ?? ""),
    unorderedList: editor.isActive("bulletList"),
    orderedList: editor.isActive("orderedList"),
    link: editor.isActive("link"),
  };
}

function sameFormat(left: FormatState, right: FormatState) {
  return Object.keys(emptyFormat).every((key) => left[key as keyof FormatState] === right[key as keyof FormatState]);
}

function normalizeHtml(value: string) {
  return value.trim() || "<p></p>";
}

function escapeAttribute(value: string) {
  return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;");
}

function escapeText(value: string) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function normalizeCandidateUrl(url: string, defaultProtocol = "https") {
  const trimmed = url.trim();
  if (/^[a-z][a-z\d+.-]*:/i.test(trimmed)) {
    return trimmed;
  }
  return `${defaultProtocol}://${trimmed}`;
}

function RichHtmlEditor({ value, placeholders, onChange }: RichHtmlEditorProps) {
  const onChangeRef = useRef(onChange);
  const lastSyncedValueRef = useRef<string | null>(null);
  const linkInputRef = useRef<HTMLInputElement | null>(null);
  const [sourceMode, setSourceMode] = useState(false);
  const [linkOpen, setLinkOpen] = useState(false);
  const [linkUrl, setLinkUrl] = useState("");
  const [linkError, setLinkError] = useState<string | null>(null);
  const [format, setFormat] = useState<FormatState>(emptyFormat);

  useEffect(() => {
    onChangeRef.current = onChange;
  }, [onChange]);

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        heading: {
          levels: [2, 3],
        },
      }),
      Underline,
      TextStyle,
      Color,
      Link.configure({
        autolink: true,
        defaultProtocol: "https",
        enableClickSelection: true,
        linkOnPaste: true,
        openOnClick: false,
        protocols: ["mailto"],
        HTMLAttributes: {
          rel: "noopener noreferrer",
          target: "_blank",
        },
        isAllowedUri: (url, ctx) => {
          return ctx.defaultValidate(url) && isSafeUrl(normalizeCandidateUrl(url, ctx.defaultProtocol));
        },
      }),
      Placeholder.configure({
        placeholder: "Rédigez votre email...",
      }),
      Table.configure({
        resizable: true,
      }),
      TableRow,
      TableHeader,
      TableCell,
    ],
    content: normalizeHtml(value),
    editorProps: {
      attributes: {
        "aria-label": "Corps HTML",
        class: "rich-surface",
        role: "textbox",
      },
      transformPastedHTML: (html) => sanitizeEmailHtml(html),
    },
    onUpdate: ({ editor: updatedEditor }) => {
      const html = updatedEditor.getHTML();
      lastSyncedValueRef.current = html;
      onChangeRef.current(html);
    },
    shouldRerenderOnTransaction: false,
  });

  useEffect(() => {
    if (!editor || sourceMode) {
      return;
    }
    if (lastSyncedValueRef.current === value) {
      return;
    }
    lastSyncedValueRef.current = value;
    const next = normalizeHtml(value);
    if (editor.getHTML() !== next) {
      editor.commands.setContent(next, { emitUpdate: false });
      setFormat(readFormat(editor));
    }
  }, [editor, sourceMode, value]);

  useEffect(() => {
    if (!editor) {
      return;
    }

    function refresh() {
      const next = readFormat(editor);
      setFormat((current) => (sameFormat(current, next) ? current : next));
    }

    refresh();
    editor.on("selectionUpdate", refresh);
    editor.on("transaction", refresh);
    return () => {
      editor.off("selectionUpdate", refresh);
      editor.off("transaction", refresh);
    };
  }, [editor]);

  useEffect(() => {
    if (!linkOpen) return;
    const id = window.requestAnimationFrame(() => linkInputRef.current?.focus());
    return () => window.cancelAnimationFrame(id);
  }, [linkOpen]);

  function runCommand(action: (editor: TiptapEditor) => void) {
    if (!editor) return;
    action(editor);
    setFormat(readFormat(editor));
  }

  function openLinkDialog() {
    if (!editor) return;
    const currentHref = editor.getAttributes("link").href;
    setLinkUrl(typeof currentHref === "string" ? currentHref : "");
    setLinkError(null);
    setLinkOpen(true);
  }

  function closeLinkDialog() {
    setLinkOpen(false);
    setLinkUrl("");
    setLinkError(null);
  }

  function confirmLink() {
    if (!editor) return;
    const trimmed = linkUrl.trim();
    if (!trimmed) {
      setLinkError("L'adresse est requise.");
      return;
    }
    if (!isSafeUrl(trimmed)) {
      setLinkError("URL invalide. Seuls http://, https:// et mailto: sont autorisés.");
      return;
    }

    const href = escapeAttribute(trimmed);
    const text = escapeText(trimmed);
    if (editor.state.selection.empty && !editor.isActive("link")) {
      editor.chain().focus().insertContent(`<a href="${href}" target="_blank" rel="noopener noreferrer">${text}</a>`).run();
    } else {
      editor.chain().focus().extendMarkRange("link").setLink({ href: trimmed }).run();
    }
    closeLinkDialog();
    setFormat(readFormat(editor));
  }

  function insertPlaceholder(placeholder: string) {
    runCommand((currentEditor) => {
      currentEditor.chain().focus().insertContent(placeholder).run();
    });
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
          onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().toggleBold().run())}
          aria-pressed={format.bold}
          disabled={!editor}
          title="Gras (Ctrl+B)"
        >
          G
        </button>
        <button
          type="button"
          className={`rich-btn rich-btn-italic${format.italic ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().toggleItalic().run())}
          aria-pressed={format.italic}
          disabled={!editor}
          title="Italique (Ctrl+I)"
        >
          I
        </button>
        <button
          type="button"
          className={`rich-btn rich-btn-underline${format.underline ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().toggleUnderline().run())}
          aria-pressed={format.underline}
          disabled={!editor}
          title="Souligné (Ctrl+U)"
        >
          S
        </button>

        <span className="rich-sep" aria-hidden="true" />

        <span className="rich-color-group" aria-label="Couleur du texte">
          {textColors.map((color) => (
            <button
              key={color.value}
              type="button"
              className={`rich-color-button${format.color === color.value ? " active" : ""}`}
              style={{ backgroundColor: color.value }}
              onMouseDown={keepSelection}
              onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().setColor(color.value).run())}
              aria-label={`Couleur ${color.label}`}
              aria-pressed={format.color === color.value}
              disabled={!editor}
              title={color.label}
            />
          ))}
          <button
            type="button"
            className="rich-btn rich-btn-subtle"
            onMouseDown={keepSelection}
            onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().unsetColor().run())}
            disabled={!editor || !format.color}
            title="Couleur par défaut"
          >
            Auto
          </button>
        </span>

        <span className="rich-sep" aria-hidden="true" />

        <button
          type="button"
          className={`rich-btn rich-btn-icon${format.unorderedList ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().toggleBulletList().run())}
          aria-pressed={format.unorderedList}
          disabled={!editor}
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
          onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().toggleOrderedList().run())}
          aria-pressed={format.orderedList}
          disabled={!editor}
          title="Liste numérotée"
        >
          <svg viewBox="0 0 20 20" width="16" height="16" aria-hidden="true">
            <text x="1" y="7" fontSize="5.5" fontWeight="700" fill="currentColor">
              1.
            </text>
            <text x="1" y="13" fontSize="5.5" fontWeight="700" fill="currentColor">
              2.
            </text>
            <text x="1" y="19" fontSize="5.5" fontWeight="700" fill="currentColor">
              3.
            </text>
            <rect x="7" y="4" width="11" height="1.6" rx="0.8" fill="currentColor" />
            <rect x="7" y="9" width="11" height="1.6" rx="0.8" fill="currentColor" />
            <rect x="7" y="14" width="11" height="1.6" rx="0.8" fill="currentColor" />
          </svg>
        </button>

        <span className="rich-sep" aria-hidden="true" />

        <button
          type="button"
          className={`rich-btn${format.link ? " active" : ""}`}
          onMouseDown={keepSelection}
          onClick={openLinkDialog}
          aria-pressed={format.link}
          disabled={!editor}
          title="Insérer un lien"
        >
          Lien
        </button>
        <button
          type="button"
          className="rich-btn rich-btn-subtle"
          onMouseDown={keepSelection}
          onClick={() => runCommand((currentEditor) => currentEditor.chain().focus().unsetLink().run())}
          disabled={!editor || !format.link}
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
              disabled={!editor}
            >
              {placeholder}
            </button>
          ))}
        </div>
      ) : null}

      {editor ? <EditorContent editor={editor} /> : <div className="rich-surface rich-loading">Chargement...</div>}
    </div>
  );
}

export default RichHtmlEditor;
