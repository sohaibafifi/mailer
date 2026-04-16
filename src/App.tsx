import { useEffect, useMemo, useState } from "react";
import {
  cancelCampaign,
  deleteTemplate,
  downloadAndInstallUpdate,
  loadSmtpConfig,
  loadTemplates,
  onCampaignProgress,
  pickExcelFile,
  previewTemplateSamples,
  saveSmtpConfig,
  saveTemplate,
  sendCampaign,
  sendSmtpTest,
} from "./api";
import type {
  CampaignProgress,
  MailingTemplate,
  PreviewSample,
  RenderedPreview,
  SendSummary,
  SmtpConfigInput,
  TemplateInput,
  WorkbookPreview,
} from "./types";
import RichHtmlEditor from "./RichHtmlEditor";
import ConfirmDialog from "./ConfirmDialog";
import { sanitizeEmailHtml } from "./sanitize";
import logoUrl from "./assets/logo.svg";

const emptyTemplate: TemplateInput = {
  id: null,
  name: "Nouveau modèle",
  subject: "Bonjour {{prenom}}",
  bodyHtml:
    "<p>Bonjour {{prenom}},</p><p>Nous revenons vers vous au sujet de {{societe}}.</p><p>Cordialement,</p>",
  bodyText: "",
};

const emptySmtp: SmtpConfigInput = {
  host: "smtprouter.univ-artois.fr",
  port: 587,
  security: "startTls",
  username: "prenom.nom@univ-artois.fr",
  fromEmail: "prenom.nom@univ-artois.fr",
  fromName: "",
  replyTo: null,
  password: null,
  clearPassword: false,
};

const defaultRateLimit = {
  maxPerMinute: "30",
  minDelayMs: "1000",
  batchSize: "4",
  batchPauseSeconds: "1",
};

type PanelId = "templates" | "excel" | "configuration" | "send" | "about";

const panels: Array<{ id: PanelId; label: string }> = [
  { id: "templates", label: "Modèles" },
  { id: "excel", label: "Excel" },
  { id: "configuration", label: "Configuration" },
  { id: "send", label: "Envoi" },
  { id: "about", label: "À propos" },
];

type Notice = {
  kind: "idle" | "success" | "error" | "info";
  text: string;
};

function getError(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

function parseOptionalInteger(
  value: string,
  label: string,
  options: { min: number; max: number },
) {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  const parsed = Number(trimmed);
  if (!Number.isInteger(parsed) || parsed < options.min || parsed > options.max) {
    throw new Error(`${label} doit être un entier entre ${options.min} et ${options.max}.`);
  }
  return parsed;
}

function findEmailField(workbook: WorkbookPreview | null) {
  if (!workbook) {
    return "";
  }
  return (
    workbook.fields.find((field) => /^(email|e_mail|mail|courriel|adresse_mail)$/i.test(field.key))?.key ??
    workbook.fields.find((field) => /email|mail|courriel/i.test(field.key))?.key ??
    workbook.fields[0]?.key ??
    ""
  );
}

function App() {
  const [templates, setTemplates] = useState<MailingTemplate[]>([]);
  const [draft, setDraft] = useState<TemplateInput>(emptyTemplate);
  const [workbook, setWorkbook] = useState<WorkbookPreview | null>(null);
  const [recipientField, setRecipientField] = useState("");
  const [smtp, setSmtp] = useState<SmtpConfigInput>(emptySmtp);
  const [smtpPasswordSaved, setSmtpPasswordSaved] = useState(false);
  const [testEmail, setTestEmail] = useState("");
  const [sendLimit, setSendLimit] = useState("");
  const [rateLimit, setRateLimit] = useState(defaultRateLimit);
  const [preview, setPreview] = useState<RenderedPreview | null>(null);
  const [previewSamples, setPreviewSamples] = useState<PreviewSample[]>([]);
  const [previewIndex, setPreviewIndex] = useState(0);
  const [summary, setSummary] = useState<SendSummary | null>(null);
  const [notice, setNotice] = useState<Notice>({ kind: "idle", text: "Prêt." });
  const [activePanel, setActivePanel] = useState<PanelId>("templates");
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState<CampaignProgress | null>(null);
  const [sending, setSending] = useState(false);
  const [confirmSend, setConfirmSend] = useState(false);

  const selectedTemplate = useMemo(
    () => templates.find((template) => template.id === draft.id) ?? null,
    [draft.id, templates],
  );

  const placeholders = workbook?.fields.map((field) => `{{${field.key}}}`) ?? [];
  const currentPreviewSample = previewSamples[previewIndex] ?? null;
  const currentPreview = currentPreviewSample ?? preview;
  const effectiveDelayMs = useMemo(() => {
    const minDelay = Number(rateLimit.minDelayMs.trim() || "0");
    const maxPerMinute = Number(rateLimit.maxPerMinute.trim() || "0");
    const minuteDelay = maxPerMinute > 0 ? Math.ceil(60_000 / maxPerMinute) : 0;
    return Math.max(Number.isFinite(minDelay) ? minDelay : 0, minuteDelay);
  }, [rateLimit.maxPerMinute, rateLimit.minDelayMs]);

  useEffect(() => {
    async function boot() {
      try {
        const [savedTemplates, savedSmtp] = await Promise.all([loadTemplates(), loadSmtpConfig()]);
        setTemplates(savedTemplates);
        if (savedTemplates[0]) {
          setDraft(savedTemplates[0]);
        }
        if (savedSmtp) {
          setSmtp({
            host: savedSmtp.host,
            port: savedSmtp.port,
            security: savedSmtp.security,
            username: savedSmtp.username,
            fromEmail: savedSmtp.fromEmail,
            fromName: savedSmtp.fromName,
            replyTo: savedSmtp.replyTo,
            password: null,
            clearPassword: false,
          });
          setSmtpPasswordSaved(savedSmtp.passwordSaved);
        }
      } catch (error) {
        setNotice({ kind: "error", text: getError(error) });
      }
    }
    void boot();
  }, []);

  useEffect(() => {
    setRecipientField(findEmailField(workbook));
  }, [workbook]);

  useEffect(() => {
    const unlistenPromise = onCampaignProgress((payload) => {
      setProgress(payload);
    });
    return () => {
      unlistenPromise.then((unlisten) => unlisten()).catch(() => undefined);
    };
  }, []);

  const safePreviewHtml = useMemo(
    () => (currentPreview ? sanitizeEmailHtml(currentPreview.bodyHtml) : ""),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [currentPreviewSample, preview],
  );

  async function runTask(task: () => Promise<void>, fallback = "Action impossible.") {
    setBusy(true);
    try {
      await task();
    } catch (error) {
      setNotice({ kind: "error", text: getError(error) || fallback });
    } finally {
      setBusy(false);
    }
  }

  function clearPreview() {
    setPreview(null);
    setPreviewSamples([]);
    setPreviewIndex(0);
  }

  function startNewTemplate() {
    setDraft(emptyTemplate);
    clearPreview();
    setSummary(null);
  }

  function selectTemplate(template: MailingTemplate) {
    setDraft(template);
    clearPreview();
    setSummary(null);
  }

  async function handleSaveTemplate() {
    await runTask(async () => {
      const saved = await saveTemplate(draft);
      setTemplates(saved);
      const next = saved.find((template) => template.name === draft.name) ?? saved[0];
      if (next) {
        setDraft(next);
      }
      setNotice({ kind: "success", text: "Modèle sauvegardé." });
    });
  }

  async function handleDeleteTemplate() {
    if (!draft.id) {
      startNewTemplate();
      return;
    }
    await runTask(async () => {
      const saved = await deleteTemplate(draft.id ?? "");
      setTemplates(saved);
      setDraft(saved[0] ?? emptyTemplate);
      clearPreview();
      setNotice({ kind: "success", text: "Modèle supprimé." });
    });
  }

  async function handlePickExcel() {
    await runTask(async () => {
      const picked = await pickExcelFile();
      if (!picked) {
        setNotice({ kind: "info", text: "Import annulé." });
        return;
      }
      setWorkbook(picked);
      clearPreview();
      setSummary(null);
      setNotice({
        kind: "success",
        text: `${picked.totalRows} lignes chargées depuis ${picked.fileName}.`,
      });
    });
  }

  async function handlePreview() {
    if (!workbook || workbook.totalRows === 0) {
      setNotice({ kind: "error", text: "Importez un fichier Excel avant la prévisualisation." });
      return;
    }
    await runTask(async () => {
      const samples = await previewTemplateSamples(
        draft,
        workbook.path,
        workbook.sheetName,
        recipientField || null,
        10,
      );
      setPreviewSamples(samples);
      setPreviewIndex(0);
      setPreview(samples[0] ?? null);
      setNotice({ kind: "success", text: `${samples.length} prévisualisations générées.` });
    });
  }

  async function handleSaveSmtp() {
    await runTask(async () => {
      const passwordForSession = smtp.password?.trim() ? smtp.password : null;
      const saved = await saveSmtpConfig(smtp);
      setSmtp({
        host: saved.host,
        port: saved.port,
        security: saved.security,
        username: saved.username,
        fromEmail: saved.fromEmail,
        fromName: saved.fromName,
        replyTo: saved.replyTo,
        password: passwordForSession,
        clearPassword: false,
      });
      setSmtpPasswordSaved(saved.passwordSaved);
      setNotice({ kind: "success", text: "Configuration SMTP sauvegardée." });
    });
  }

  async function handleSmtpTest() {
    if (!testEmail.trim()) {
      setNotice({ kind: "error", text: "Renseignez une adresse de test." });
      return;
    }
    if (smtp.username.trim() && !smtp.password?.trim() && !smtpPasswordSaved) {
      setNotice({ kind: "error", text: "Saisissez le mot de passe SMTP avant de tester." });
      return;
    }
    await runTask(async () => {
      const message = await sendSmtpTest(smtp, testEmail.trim());
      setNotice({ kind: "success", text: message });
    });
  }

  function buildSendRequest(dryRun: boolean, email: string | null) {
    if (!draft.id) {
      throw new Error("Sauvegardez le modèle avant l'envoi.");
    }
    if (!workbook) {
      throw new Error("Importez un fichier Excel avant l'envoi.");
    }
    if (!recipientField) {
      throw new Error("Sélectionnez le champ qui contient l'adresse email.");
    }
    const parsedLimit = sendLimit.trim() ? Number(sendLimit) : null;
    if (parsedLimit !== null && (!Number.isFinite(parsedLimit) || parsedLimit < 1)) {
      throw new Error("La limite doit être vide ou supérieure à 0.");
    }
    const parsedRateLimit = {
      maxPerMinute: parseOptionalInteger(rateLimit.maxPerMinute, "Le maximum par minute", {
        min: 1,
        max: 600,
      }),
      minDelayMs: parseOptionalInteger(rateLimit.minDelayMs, "La pause minimale", {
        min: 0,
        max: 600_000,
      }),
      batchSize: parseOptionalInteger(rateLimit.batchSize, "La taille de lot", {
        min: 1,
        max: 100_000,
      }),
      batchPauseMs:
        parseOptionalInteger(rateLimit.batchPauseSeconds, "La pause entre lots", {
          min: 0,
          max: 3600,
        }) === null
          ? null
          : Number(rateLimit.batchPauseSeconds.trim()) * 1000,
    };
    if (parsedRateLimit.batchPauseMs !== null && parsedRateLimit.batchPauseMs > 0 && parsedRateLimit.batchSize === null) {
      throw new Error("Renseignez une taille de lot pour utiliser une pause entre lots.");
    }
    return {
      templateId: draft.id,
      excelPath: workbook.path,
      sheetName: workbook.sheetName,
      recipientField,
      dryRun,
      testEmail: email,
      limit: parsedLimit,
      rateLimit: parsedRateLimit,
    };
  }

  function requestSendCampaign() {
    try {
      buildSendRequest(false, null);
      setConfirmSend(true);
    } catch (error) {
      setNotice({ kind: "error", text: getError(error) });
    }
  }

  async function handleSendCampaign() {
    setConfirmSend(false);
    setSending(true);
    setProgress(null);
    await runTask(async () => {
      try {
        const result = await sendCampaign(buildSendRequest(false, null));
        setSummary(result);
        setPreview(result.preview);
        setPreviewSamples([]);
        setPreviewIndex(0);
        const cancelled = progress?.status === "cancelled";
        setNotice({
          kind: result.failed.length > 0 ? "error" : cancelled ? "info" : "success",
          text: cancelled
            ? `Envoi interrompu: ${result.sent} envoyés, ${result.failed.length} erreurs.`
            : `${result.sent} envoyés, ${result.skipped} ignorés, ${result.failed.length} erreurs.`,
        });
      } finally {
        setSending(false);
      }
    });
  }

  async function handleCancelCampaign() {
    try {
      await cancelCampaign();
      setNotice({ kind: "info", text: "Annulation demandée..." });
    } catch (error) {
      setNotice({ kind: "error", text: getError(error) });
    }
  }

  async function handleUpdate() {
    await runTask(async () => {
      const result = await downloadAndInstallUpdate();
      setNotice({ kind: result.status === "installed" ? "success" : "info", text: result.message });
    });
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark" aria-hidden="true">
            <img src={logoUrl} alt="" />
          </div>
          <div className="brand-text">
            <p className="eyebrow">Mailer</p>
            <h1>Publipostage email</h1>
          </div>
        </div>
        <span className={`notice ${notice.kind}`}>{notice.text}</span>
      </header>

      <nav className="panel-tabs" aria-label="Navigation principale">
        {panels.map((panel) => (
          <button
            key={panel.id}
            type="button"
            className={activePanel === panel.id ? "tab active" : "tab"}
            onClick={() => setActivePanel(panel.id)}
          >
            {panel.label}
          </button>
        ))}
      </nav>

      <section className="workspace">
        {activePanel === "templates" ? (
          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>Modèles</h2>
                <p>Choisissez un modèle ou créez-en un nouveau.</p>
              </div>
              <div className="button-row">
                <button type="button" className="ghost" onClick={startNewTemplate} disabled={busy}>
                  Nouveau
                </button>
                <button type="button" className="ghost" onClick={handleDeleteTemplate} disabled={busy || !draft.id}>
                  Supprimer
                </button>
                <button type="button" className="primary" onClick={handleSaveTemplate} disabled={busy}>
                  Sauvegarder
                </button>
              </div>
            </div>

            <label>
              Modèle sauvegardé
              <select
                value={draft.id ?? ""}
                onChange={(event) => {
                  const template = templates.find((item) => item.id === event.target.value);
                  if (template) {
                    selectTemplate(template);
                  }
                }}
              >
                <option value="" disabled>
                  {templates.length ? "Sélectionner" : "Aucun modèle sauvegardé"}
                </option>
                {templates.map((template) => (
                  <option key={template.id} value={template.id}>
                    {template.name}
                  </option>
                ))}
              </select>
            </label>
            <div className="template-name-row">
              <label>
                Nom du modèle
                <input
                  value={draft.name}
                  onChange={(event) => setDraft({ ...draft, name: event.target.value })}
                  placeholder="Relance salon"
                />
              </label>
            </div>

            <label>
              Sujet
              <input
                value={draft.subject}
                onChange={(event) => setDraft({ ...draft, subject: event.target.value })}
                placeholder="Bonjour {{prenom}}"
              />
            </label>

            <div className="field-group">
              <span className="field-label">Corps HTML</span>
              <RichHtmlEditor
                value={draft.bodyHtml}
                placeholders={placeholders}
                onChange={(bodyHtml) => setDraft({ ...draft, bodyHtml, bodyText: "" })}
              />
            </div>

            {placeholders.length ? (
              <div className="field-strip">
                {placeholders.map((placeholder) => (
                  <code key={placeholder}>{placeholder}</code>
                ))}
              </div>
            ) : (
              <p className="muted">Importez un fichier Excel pour afficher les champs disponibles.</p>
            )}
          </section>
        ) : null}

        {activePanel === "excel" ? (
          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>Excel</h2>
                <p>La première ligne doit contenir les noms de colonnes.</p>
              </div>
              <button type="button" className="primary" onClick={handlePickExcel} disabled={busy}>
                Importer Excel
              </button>
            </div>

            {workbook ? (
              <>
                <div className="status-row">
                  <span>{workbook.fileName}</span>
                  <span>{workbook.sheetName}</span>
                  <span>{workbook.totalRows} lignes</span>
                </div>
                <div className="field-strip">
                  {workbook.fields.map((field) => (
                    <code key={field.key} title={field.header}>
                      {`{{${field.key}}}`}
                    </code>
                  ))}
                </div>
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Ligne</th>
                        {workbook.fields.slice(0, 6).map((field) => (
                          <th key={field.key}>{field.header}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {workbook.rows.slice(0, 8).map((row) => (
                        <tr key={row.rowNumber}>
                          <td>{row.rowNumber}</td>
                          {workbook.fields.slice(0, 6).map((field) => (
                            <td key={field.key}>{row.values[field.key]}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            ) : (
              <p className="muted">Aucun fichier chargé.</p>
            )}
          </section>
        ) : null}

        {activePanel === "configuration" ? (
          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>Configuration</h2>
                <p>Réglez SMTP et la temporisation des campagnes.</p>
              </div>
              <div className="button-row">
                <button type="button" className="ghost" onClick={handleSmtpTest} disabled={busy}>
                  Tester SMTP
                </button>
                <button type="button" className="primary" onClick={handleSaveSmtp} disabled={busy}>
                  Sauvegarder
                </button>
              </div>
            </div>

            <section className="settings-section">
              <h3>SMTP</h3>
              <div className="settings-grid smtp-settings">
                <label className="span-two">
                  Serveur
                  <input
                    value={smtp.host}
                    onChange={(event) => setSmtp({ ...smtp, host: event.target.value })}
                    placeholder="smtp.exemple.fr"
                  />
                </label>
                <label>
                  Port
                  <input
                    type="number"
                    min="1"
                    value={smtp.port}
                    onChange={(event) => setSmtp({ ...smtp, port: Number(event.target.value) })}
                  />
                </label>
                <label>
                  Sécurité
                  <select
                    value={smtp.security}
                    onChange={(event) => setSmtp({ ...smtp, security: event.target.value as SmtpConfigInput["security"] })}
                  >
                    <option value="startTls">STARTTLS</option>
                    <option value="tls">TLS direct</option>
                    <option value="none">Aucune</option>
                  </select>
                </label>
                <label className="span-two">
                  Identifiant
                  <input
                    value={smtp.username}
                    onChange={(event) => setSmtp({ ...smtp, username: event.target.value })}
                    placeholder="compte SMTP"
                  />
                </label>
                <label className="span-two">
                  Mot de passe
                  <input
                    type="password"
                    value={smtp.password ?? ""}
                    onChange={(event) => setSmtp({ ...smtp, password: event.target.value, clearPassword: false })}
                    placeholder={smtpPasswordSaved ? "laisser vide pour conserver" : "optionnel à la sauvegarde"}
                  />
                  <span className={smtpPasswordSaved ? "field-hint success-text" : "field-hint"}>
                    {smtpPasswordSaved
                      ? "Mot de passe enregistré."
                      : "Vous pourrez sauvegarder sans mot de passe, mais il sera requis pour tester ou envoyer."}
                  </span>
                </label>
                <label>
                  Email de test
                  <input
                    value={testEmail}
                    onChange={(event) => setTestEmail(event.target.value)}
                    placeholder="vous@exemple.fr"
                  />
                </label>
                <label>
                  Email expéditeur
                  <input
                    value={smtp.fromEmail}
                    onChange={(event) => setSmtp({ ...smtp, fromEmail: event.target.value })}
                    placeholder="contact@exemple.fr"
                  />
                </label>
                <label>
                  Nom expéditeur
                  <input
                    value={smtp.fromName}
                    onChange={(event) => setSmtp({ ...smtp, fromName: event.target.value })}
                    placeholder="Votre équipe"
                  />
                </label>
                <label>
                  Réponse à
                  <input
                    value={smtp.replyTo ?? ""}
                    onChange={(event) => setSmtp({ ...smtp, replyTo: event.target.value || null })}
                    placeholder="optionnel"
                  />
                </label>
              </div>
            </section>

            <section className="settings-section">
              <h3>Temporisation</h3>
              <div className="settings-grid rate-settings">
                <label>
                  Maximum par minute
                  <input
                    value={rateLimit.maxPerMinute}
                    onChange={(event) => setRateLimit({ ...rateLimit, maxPerMinute: event.target.value })}
                    placeholder="30"
                  />
                </label>
                <label>
                  Pause minimale (ms)
                  <input
                    value={rateLimit.minDelayMs}
                    onChange={(event) => setRateLimit({ ...rateLimit, minDelayMs: event.target.value })}
                    placeholder="1000"
                  />
                </label>
                <label>
                  Taille de lot
                  <input
                    value={rateLimit.batchSize}
                    onChange={(event) => setRateLimit({ ...rateLimit, batchSize: event.target.value })}
                    placeholder="vide"
                  />
                </label>
                <label>
                  Pause entre lots (s)
                  <input
                    value={rateLimit.batchPauseSeconds}
                    onChange={(event) => setRateLimit({ ...rateLimit, batchPauseSeconds: event.target.value })}
                    placeholder="vide"
                  />
                </label>
              </div>
              <div className="status-row">
                <span>{effectiveDelayMs} ms minimum entre deux emails</span>
                <span>
                  {rateLimit.batchSize.trim()
                    ? `Pause de ${rateLimit.batchPauseSeconds.trim() || "0"} s tous les ${rateLimit.batchSize.trim()} emails`
                    : "Pas de pause par lot"}
                </span>
              </div>
            </section>
          </section>
        ) : null}

        {activePanel === "send" ? (
          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>Envoi</h2>
                <p>Simulez, vérifiez, puis lancez la campagne.</p>
              </div>
              <div className="button-row">
                <button type="button" className="ghost" onClick={handlePreview} disabled={busy || !workbook}>
                  Prévisualiser
                </button>
                <button
                  type="button"
                  className="danger"
                  onClick={requestSendCampaign}
                  disabled={busy || sending}
                >
                  Envoyer
                </button>
                {sending ? (
                  <button type="button" className="ghost" onClick={handleCancelCampaign}>
                    Annuler l'envoi
                  </button>
                ) : null}
              </div>
            </div>

            <div className="status-row">
              <span>{selectedTemplate ? selectedTemplate.name : "Modèle non sauvegardé"}</span>
              <span>{workbook ? workbook.fileName : "Aucun Excel"}</span>
              <span>{smtp.host || "SMTP non configuré"}</span>
            </div>

            <div className="form-grid two">
              <label>
                Champ email
                <select
                  value={recipientField}
                  onChange={(event) => {
                    setRecipientField(event.target.value);
                    setPreviewSamples([]);
                    setPreviewIndex(0);
                  }}
                >
                  {(workbook?.fields ?? []).map((field) => (
                    <option key={field.key} value={field.key}>
                      {field.header} ({field.key})
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Limite d'envoi
                <input value={sendLimit} onChange={(event) => setSendLimit(event.target.value)} placeholder="vide" />
              </label>
            </div>

            {sending && progress ? (
              <div className="progress-card">
                <div className="progress-head">
                  <strong>
                    {progress.attempted} / {progress.total}
                  </strong>
                  <span>{progress.sent} envoyés</span>
                  <span>{progress.skipped} ignorés</span>
                  <span>{progress.failed} erreurs</span>
                  {progress.currentRecipient ? (
                    <span className="progress-current">→ {progress.currentRecipient}</span>
                  ) : null}
                </div>
                <div
                  className="progress-bar"
                  role="progressbar"
                  aria-valuemin={0}
                  aria-valuemax={progress.total || 1}
                  aria-valuenow={progress.attempted}
                >
                  <div
                    className="progress-fill"
                    style={{
                      width: `${progress.total ? Math.min(100, (progress.attempted / progress.total) * 100) : 0}%`,
                    }}
                  />
                </div>
                {progress.lastError ? (
                  <p className="progress-error">Dernière erreur : {progress.lastError}</p>
                ) : null}
              </div>
            ) : null}

            {summary ? (
              <div className="summary">
                <strong>{summary.attempted} traités</strong>
                <span>{summary.sent} envoyés</span>
                <span>{summary.skipped} ignorés</span>
                <span>{summary.failed.length} erreurs</span>
              </div>
            ) : null}

            {summary?.failed.length ? (
              <div className="failure-list">
                {summary.failed.slice(0, 8).map((failure) => (
                  <p key={`${failure.rowNumber}-${failure.message}`}>
                    Ligne {failure.rowNumber} {failure.recipient ? `(${failure.recipient})` : ""}: {failure.message}
                  </p>
                ))}
              </div>
            ) : null}

            {currentPreview ? (
              <section className="preview-area">
                <div className="preview-heading">
                  <div>
                    <h3>Prévisualisation</h3>
                    {currentPreviewSample ? (
                      <p className="preview-meta">
                        Ligne {currentPreviewSample.rowNumber}
                        {currentPreviewSample.recipient ? ` - ${currentPreviewSample.recipient}` : ""}
                      </p>
                    ) : null}
                  </div>
                  {previewSamples.length > 1 ? (
                    <div className="preview-nav">
                      <button
                        type="button"
                        className="ghost compact"
                        onClick={() => setPreviewIndex((index) => Math.max(0, index - 1))}
                        disabled={previewIndex === 0}
                      >
                        Précédent
                      </button>
                      <span>
                        {previewIndex + 1} / {previewSamples.length}
                      </span>
                      <button
                        type="button"
                        className="ghost compact"
                        onClick={() => setPreviewIndex((index) => Math.min(previewSamples.length - 1, index + 1))}
                        disabled={previewIndex >= previewSamples.length - 1}
                      >
                        Suivant
                      </button>
                    </div>
                  ) : null}
                </div>
                <p className="preview-subject">{currentPreview.subject}</p>
                <div className="preview-box" dangerouslySetInnerHTML={{ __html: safePreviewHtml }} />
              </section>
            ) : (
              <p className="muted">Aucune prévisualisation.</p>
            )}
          </section>
        ) : null}

        {activePanel === "about" ? (
          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>À propos</h2>
                <p>Application bureau locale pour publipostage email.</p>
              </div>
              <button className="primary" type="button" onClick={handleUpdate} disabled={busy}>
                Vérifier les mises à jour
              </button>
            </div>
            <div className="about-list">
              <p><strong>Version</strong><span>0.1.0</span></p>
            </div>
          </section>
        ) : null}
      </section>

      <ConfirmDialog
        open={confirmSend}
        title="Lancer la campagne ?"
        message={`${workbook?.totalRows ?? 0} ligne(s) seront traitées via ${smtp.host || "le SMTP configuré"}. Cette action enverra de vrais emails.`}
        confirmLabel="Envoyer maintenant"
        cancelLabel="Revenir"
        tone="danger"
        onConfirm={handleSendCampaign}
        onCancel={() => setConfirmSend(false)}
      />
    </main>
  );
}

export default App;
