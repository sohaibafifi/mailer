import { useEffect, useMemo, useState } from "react";
import {
  deleteTemplate,
  downloadAndInstallUpdate,
  loadSmtpConfig,
  loadTemplates,
  pickExcelFile,
  previewTemplate,
  saveSmtpConfig,
  saveTemplate,
  sendCampaign,
  sendSmtpTest,
} from "./api";
import type {
  MailingTemplate,
  RenderedPreview,
  SendSummary,
  SmtpConfigInput,
  TemplateInput,
  WorkbookPreview,
} from "./types";

const emptyTemplate: TemplateInput = {
  id: null,
  name: "Nouveau modèle",
  subject: "Bonjour {{prenom}}",
  bodyHtml:
    "<p>Bonjour {{prenom}},</p><p>Nous revenons vers vous au sujet de {{societe}}.</p><p>Cordialement,</p>",
  bodyText: "Bonjour {{prenom}},\n\nNous revenons vers vous au sujet de {{societe}}.\n\nCordialement,",
};

const emptySmtp: SmtpConfigInput = {
  host: "",
  port: 587,
  security: "startTls",
  username: "",
  fromEmail: "",
  fromName: "",
  replyTo: null,
  password: null,
  clearPassword: false,
};

const defaultRateLimit = {
  maxPerMinute: "30",
  minDelayMs: "1000",
  batchSize: "",
  batchPauseSeconds: "",
};

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
  const [testEmail, setTestEmail] = useState("");
  const [sendLimit, setSendLimit] = useState("");
  const [rateLimit, setRateLimit] = useState(defaultRateLimit);
  const [preview, setPreview] = useState<RenderedPreview | null>(null);
  const [summary, setSummary] = useState<SendSummary | null>(null);
  const [notice, setNotice] = useState<Notice>({ kind: "idle", text: "Prêt." });
  const [busy, setBusy] = useState(false);

  const selectedTemplate = useMemo(
    () => templates.find((template) => template.id === draft.id) ?? null,
    [draft.id, templates],
  );

  const sampleRow = workbook?.rows[0]?.values ?? null;
  const placeholders = workbook?.fields.map((field) => `{{${field.key}}}`) ?? [];
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

  function startNewTemplate() {
    setDraft(emptyTemplate);
    setPreview(null);
    setSummary(null);
  }

  function selectTemplate(template: MailingTemplate) {
    setDraft(template);
    setPreview(null);
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
      setPreview(null);
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
      setSummary(null);
      setNotice({
        kind: "success",
        text: `${picked.totalRows} lignes chargées depuis ${picked.fileName}.`,
      });
    });
  }

  async function handlePreview() {
    if (!sampleRow) {
      setNotice({ kind: "error", text: "Importez un fichier Excel avant la prévisualisation." });
      return;
    }
    await runTask(async () => {
      const rendered = await previewTemplate(draft, sampleRow);
      setPreview(rendered);
      setNotice({ kind: "success", text: "Prévisualisation générée." });
    });
  }

  async function handleSaveSmtp() {
    await runTask(async () => {
      const saved = await saveSmtpConfig(smtp);
      setSmtp({
        host: saved.host,
        port: saved.port,
        security: saved.security,
        username: saved.username,
        fromEmail: saved.fromEmail,
        fromName: saved.fromName,
        replyTo: saved.replyTo,
        password: null,
        clearPassword: false,
      });
      setNotice({ kind: "success", text: "Configuration SMTP sauvegardée." });
    });
  }

  async function handleSmtpTest() {
    if (!testEmail.trim()) {
      setNotice({ kind: "error", text: "Renseignez une adresse de test." });
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

  async function handleDryRun() {
    await runTask(async () => {
      const result = await sendCampaign(buildSendRequest(true, null));
      setSummary(result);
      setPreview(result.preview);
      setNotice({ kind: "success", text: "Simulation terminée sans envoi SMTP." });
    });
  }

  async function handleSendTestCampaign() {
    if (!testEmail.trim()) {
      setNotice({ kind: "error", text: "Renseignez une adresse de test." });
      return;
    }
    await runTask(async () => {
      const result = await sendCampaign(buildSendRequest(false, testEmail.trim()));
      setSummary(result);
      setPreview(result.preview);
      setNotice({ kind: "success", text: "Email de test envoyé avec les données de la première ligne." });
    });
  }

  async function handleSendCampaign() {
    await runTask(async () => {
      const result = await sendCampaign(buildSendRequest(false, null));
      setSummary(result);
      setPreview(result.preview);
      setNotice({
        kind: result.failed.length > 0 ? "error" : "success",
        text: `${result.sent} envoyés, ${result.skipped} ignorés, ${result.failed.length} erreurs.`,
      });
    });
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
        <div>
          <p className="eyebrow">ArtoisMailer</p>
          <h1>Publipostage email</h1>
        </div>
        <div className="topbar-actions">
          <button className="ghost" type="button" onClick={handleUpdate} disabled={busy}>
            Vérifier les mises à jour
          </button>
          <span className={`notice ${notice.kind}`}>{notice.text}</span>
        </div>
      </header>

      <section className="workspace">
        <aside className="sidebar">
          <button type="button" className="primary" onClick={startNewTemplate} disabled={busy}>
            Nouveau modèle
          </button>
          <div className="template-list" aria-label="Modèles sauvegardés">
            {templates.length === 0 ? (
              <p className="muted">Aucun modèle sauvegardé.</p>
            ) : (
              templates.map((template) => (
                <button
                  key={template.id}
                  type="button"
                  className={template.id === selectedTemplate?.id ? "template-item active" : "template-item"}
                  onClick={() => selectTemplate(template)}
                >
                  <strong>{template.name}</strong>
                  <span>{template.subject}</span>
                </button>
              ))
            )}
          </div>
        </aside>

        <section className="main-column">
          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>1. Modèle</h2>
                <p>Rédigez le sujet et le contenu HTML avec les champs Excel entre doubles accolades.</p>
              </div>
              <div className="button-row">
                <button type="button" className="ghost" onClick={handleDeleteTemplate} disabled={busy}>
                  Supprimer
                </button>
                <button type="button" className="primary" onClick={handleSaveTemplate} disabled={busy}>
                  Sauvegarder
                </button>
              </div>
            </div>
            <div className="form-grid">
              <label>
                Nom du modèle
                <input
                  value={draft.name}
                  onChange={(event) => setDraft({ ...draft, name: event.target.value })}
                  placeholder="Relance salon"
                />
              </label>
              <label>
                Sujet
                <input
                  value={draft.subject}
                  onChange={(event) => setDraft({ ...draft, subject: event.target.value })}
                  placeholder="Bonjour {{prenom}}"
                />
              </label>
            </div>
            <label>
              Corps HTML
              <textarea
                className="body-editor"
                value={draft.bodyHtml}
                onChange={(event) => setDraft({ ...draft, bodyHtml: event.target.value })}
              />
            </label>
            <label>
              Version texte optionnelle
              <textarea
                className="plain-editor"
                value={draft.bodyText}
                onChange={(event) => setDraft({ ...draft, bodyText: event.target.value })}
              />
            </label>
          </section>

          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>2. Fichier Excel</h2>
                <p>La première ligne doit contenir les noms des colonnes. Les champs sont normalisés pour les modèles.</p>
              </div>
              <button type="button" className="primary" onClick={handlePickExcel} disabled={busy}>
                Importer Excel
              </button>
            </div>
            {workbook ? (
              <>
                <div className="import-meta">
                  <span>{workbook.fileName}</span>
                  <span>{workbook.sheetName}</span>
                  <span>{workbook.totalRows} lignes</span>
                </div>
                <div className="field-strip">
                  {workbook.fields.map((field) => (
                    <span key={field.key} title={field.header}>
                      {`{{${field.key}}}`}
                    </span>
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
                      {workbook.rows.slice(0, 6).map((row) => (
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

          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>3. SMTP</h2>
                <p>La configuration est locale. Le mot de passe est conservé par le trousseau du système.</p>
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
            <div className="form-grid three">
              <label>
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
              <label>
                Identifiant
                <input
                  value={smtp.username}
                  onChange={(event) => setSmtp({ ...smtp, username: event.target.value })}
                  placeholder="compte SMTP"
                />
              </label>
              <label>
                Mot de passe
                <input
                  type="password"
                  value={smtp.password ?? ""}
                  onChange={(event) => setSmtp({ ...smtp, password: event.target.value, clearPassword: false })}
                  placeholder="laisser vide pour conserver"
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
              <label>
                Email de test
                <input
                  value={testEmail}
                  onChange={(event) => setTestEmail(event.target.value)}
                  placeholder="vous@exemple.fr"
                />
              </label>
            </div>
          </section>

          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>4. Rate limit</h2>
                <p>Temporisez les campagnes pour respecter les quotas SMTP et éviter les rafales.</p>
              </div>
            </div>
            <div className="form-grid four">
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
            <div className="rate-summary">
              <span>{effectiveDelayMs} ms minimum entre deux emails</span>
              <span>
                {rateLimit.batchSize.trim()
                  ? `Pause de ${rateLimit.batchPauseSeconds.trim() || "0"} s tous les ${rateLimit.batchSize.trim()} emails`
                  : "Pas de pause par lot"}
              </span>
            </div>
          </section>

          <section className="panel">
            <div className="panel-heading">
              <div>
                <h2>5. Envoi</h2>
                <p>Simulez d'abord, envoyez un test, puis lancez la campagne.</p>
              </div>
              <div className="button-row">
                <button type="button" className="ghost" onClick={handlePreview} disabled={busy || !sampleRow}>
                  Prévisualiser
                </button>
                <button type="button" className="ghost" onClick={handleDryRun} disabled={busy}>
                  Simuler
                </button>
                <button type="button" className="ghost" onClick={handleSendTestCampaign} disabled={busy}>
                  Test campagne
                </button>
                <button type="button" className="danger" onClick={handleSendCampaign} disabled={busy}>
                  Envoyer
                </button>
              </div>
            </div>
            <div className="form-grid three">
              <label>
                Champ email
                <select value={recipientField} onChange={(event) => setRecipientField(event.target.value)}>
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
          </section>
        </section>

        <aside className="inspector">
          <section className="panel slim">
            <h2>Champs</h2>
            {placeholders.length ? (
              <div className="placeholder-list">
                {placeholders.map((placeholder) => (
                  <code key={placeholder}>{placeholder}</code>
                ))}
              </div>
            ) : (
              <p className="muted">Importez Excel pour voir les champs.</p>
            )}
          </section>
          <section className="panel slim">
            <h2>Prévisualisation</h2>
            {preview ? (
              <>
                <p className="preview-subject">{preview.subject}</p>
                <div className="preview-box" dangerouslySetInnerHTML={{ __html: preview.bodyHtml }} />
              </>
            ) : (
              <p className="muted">Aucune prévisualisation.</p>
            )}
          </section>
        </aside>
      </section>
    </main>
  );
}

export default App;
