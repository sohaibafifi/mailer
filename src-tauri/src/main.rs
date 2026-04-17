#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::{
    collections::{BTreeMap, HashSet},
    fs,
    path::PathBuf,
    sync::atomic::{AtomicBool, Ordering},
    thread,
    time::{Duration, SystemTime, UNIX_EPOCH},
};

use calamine::{open_workbook_auto, Data, Reader};
use chrono::Utc;
use deunicode::deunicode;
use handlebars::{no_escape, Handlebars, RenderError, RenderErrorReason};
use keyring::Entry;
use lettre::{
    message::{header, Mailbox, MultiPart, SinglePart},
    transport::smtp::{authentication::Credentials, SmtpTransportBuilder},
    Message, SmtpTransport, Transport,
};
use serde::{de::DeserializeOwned, Deserialize, Serialize};
use tauri::{AppHandle, Emitter, Manager};
use tauri_plugin_updater::UpdaterExt;
use uuid::Uuid;

static CAMPAIGN_CANCEL: AtomicBool = AtomicBool::new(false);
const PROGRESS_EVENT: &str = "campaign:progress";

const TEMPLATES_FILE: &str = "templates.json";
const SMTP_FILE: &str = "smtp.json";
const KEYRING_SERVICE: &str = "fr.univ-artois.mailer";
const KEYRING_ACCOUNT: &str = "smtp-password";
const MAX_FAILURES: usize = 100;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct FieldDescriptor {
    header: String,
    key: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RecipientRow {
    row_number: usize,
    values: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WorkbookPreview {
    path: String,
    file_name: String,
    sheet_name: String,
    sheets: Vec<String>,
    fields: Vec<FieldDescriptor>,
    rows: Vec<RecipientRow>,
    total_rows: usize,
}

#[derive(Debug, Clone)]
struct WorkbookData {
    fields: Vec<FieldDescriptor>,
    rows: Vec<RecipientRow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MailingTemplate {
    id: String,
    name: String,
    subject: String,
    body_html: String,
    body_text: String,
    created_at: String,
    updated_at: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct TemplateInput {
    id: Option<String>,
    name: String,
    subject: String,
    body_html: String,
    body_text: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
enum SmtpSecurity {
    StartTls,
    Tls,
    None,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredSmtpConfig {
    host: String,
    port: u16,
    security: SmtpSecurity,
    username: String,
    from_email: String,
    from_name: String,
    reply_to: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SmtpConfig {
    host: String,
    port: u16,
    security: SmtpSecurity,
    username: String,
    from_email: String,
    from_name: String,
    reply_to: Option<String>,
    password_saved: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SmtpConfigInput {
    host: String,
    port: u16,
    security: SmtpSecurity,
    username: String,
    from_email: String,
    from_name: String,
    reply_to: Option<String>,
    password: Option<String>,
    clear_password: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SendCampaignRequest {
    template_id: String,
    excel_path: String,
    sheet_name: Option<String>,
    recipient_field: String,
    dry_run: bool,
    test_email: Option<String>,
    limit: Option<usize>,
    rate_limit: RateLimitConfig,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct RateLimitConfig {
    max_per_minute: Option<u64>,
    min_delay_ms: Option<u64>,
    batch_size: Option<usize>,
    batch_pause_ms: Option<u64>,
}

#[derive(Debug, Clone)]
struct RateLimitRuntime {
    per_email_delay: Duration,
    batch_size: Option<usize>,
    batch_pause: Duration,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct RenderedPreview {
    subject: String,
    body_html: String,
    body_text: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct PreviewSample {
    row_number: usize,
    recipient: Option<String>,
    subject: String,
    body_html: String,
    body_text: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct RowFailure {
    row_number: usize,
    recipient: Option<String>,
    message: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct SendSummary {
    attempted: usize,
    sent: usize,
    skipped: usize,
    failed: Vec<RowFailure>,
    preview: Option<RenderedPreview>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct UpdateCheckResult {
    status: String,
    version: Option<String>,
    message: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct CampaignProgress {
    status: &'static str,
    total: usize,
    attempted: usize,
    sent: usize,
    skipped: usize,
    failed: usize,
    current_row: Option<usize>,
    current_recipient: Option<String>,
    last_error: Option<String>,
}

#[tauri::command]
fn load_templates(app: AppHandle) -> Result<Vec<MailingTemplate>, String> {
    load_templates_inner(&app)
}

#[tauri::command]
fn save_template(app: AppHandle, input: TemplateInput) -> Result<Vec<MailingTemplate>, String> {
    validate_template(&input)?;

    let mut templates = load_templates_inner(&app)?;
    let now = Utc::now().to_rfc3339();
    let mut saved_id = input.id.clone().unwrap_or_default();

    if let Some(id) = input.id.as_deref() {
        if let Some(existing) = templates.iter_mut().find(|template| template.id == id) {
            existing.name = input.name.trim().to_string();
            existing.subject = input.subject.trim().to_string();
            existing.body_html = input.body_html.trim().to_string();
            existing.body_text = input.body_text.trim().to_string();
            existing.updated_at = now.clone();
            saved_id = existing.id.clone();
        }
    }

    if saved_id.is_empty() || !templates.iter().any(|template| template.id == saved_id) {
        saved_id = Uuid::new_v4().to_string();
        templates.push(MailingTemplate {
            id: saved_id,
            name: input.name.trim().to_string(),
            subject: input.subject.trim().to_string(),
            body_html: input.body_html.trim().to_string(),
            body_text: input.body_text.trim().to_string(),
            created_at: now.clone(),
            updated_at: now,
        });
    }

    templates.sort_by(|a, b| b.updated_at.cmp(&a.updated_at));
    write_json(&app, TEMPLATES_FILE, &templates)?;
    Ok(templates)
}

#[tauri::command]
fn delete_template(app: AppHandle, id: String) -> Result<Vec<MailingTemplate>, String> {
    let mut templates = load_templates_inner(&app)?;
    templates.retain(|template| template.id != id);
    write_json(&app, TEMPLATES_FILE, &templates)?;
    Ok(templates)
}

#[tauri::command]
fn load_smtp_config(app: AppHandle) -> Result<Option<SmtpConfig>, String> {
    let stored = read_optional_json::<StoredSmtpConfig>(&app, SMTP_FILE)?;
    Ok(stored.map(|config| SmtpConfig {
        host: config.host,
        port: config.port,
        security: config.security,
        username: config.username,
        from_email: config.from_email,
        from_name: config.from_name,
        reply_to: config.reply_to,
        password_saved: smtp_password_saved(),
    }))
}

#[tauri::command]
fn save_smtp_config(app: AppHandle, input: SmtpConfigInput) -> Result<SmtpConfig, String> {
    let stored = normalize_smtp_input(&input)?;

    if input.clear_password {
        clear_smtp_password()?;
    }
    if let Some(password) = input
        .password
        .as_deref()
        .map(str::trim)
        .filter(|value| !value.is_empty())
    {
        keyring_entry()?.set_password(password).map_err(|error| {
            format!("Impossible d'enregistrer le mot de passe SMTP dans le trousseau: {error}")
        })?;
    }

    write_json(&app, SMTP_FILE, &stored)?;
    Ok(SmtpConfig {
        host: stored.host,
        port: stored.port,
        security: stored.security,
        username: stored.username,
        from_email: stored.from_email,
        from_name: stored.from_name,
        reply_to: stored.reply_to,
        password_saved: smtp_password_saved(),
    })
}

#[tauri::command]
async fn send_smtp_test(input: SmtpConfigInput, email: String) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let config = normalize_smtp_input(&input)?;
        let password = resolve_password(input.password)?;
        let mailer = build_transport(&config, password.as_deref())?;
        let html = "<p>La configuration SMTP Mailer est opérationnelle.</p>";
        send_email(
            &mailer,
            &config,
            email.trim(),
            "Test Mailer",
            html,
            "La configuration SMTP Mailer est opérationnelle.",
        )?;
        Ok("Email de test envoyé.".to_string())
    })
    .await
    .map_err(|error| format!("Tâche SMTP interrompue: {error}"))?
}

#[tauri::command]
fn pick_excel_file() -> Result<Option<WorkbookPreview>, String> {
    let Some(path) = rfd::FileDialog::new()
        .add_filter("Excel", &["xlsx", "xls", "xlsm", "xlsb", "ods"])
        .pick_file()
    else {
        return Ok(None);
    };

    read_workbook_preview(path, None).map(Some)
}

#[tauri::command]
fn preview_template(
    input: TemplateInput,
    row: BTreeMap<String, String>,
) -> Result<RenderedPreview, String> {
    validate_template(&input)?;
    render_template_parts(&input.subject, &input.body_html, &row)
}

#[tauri::command]
fn preview_template_samples(
    input: TemplateInput,
    excel_path: String,
    sheet_name: Option<String>,
    recipient_field: Option<String>,
    count: usize,
) -> Result<Vec<PreviewSample>, String> {
    validate_template(&input)?;
    let workbook = read_workbook_data(PathBuf::from(excel_path), sheet_name)?;
    if workbook.rows.is_empty() {
        return Err("Aucune ligne exploitable dans le fichier Excel.".to_string());
    }

    let recipient_key = recipient_field.as_deref().and_then(trimmed_optional);
    if let Some(field) = recipient_key.as_deref() {
        if !workbook
            .fields
            .iter()
            .any(|candidate| candidate.key == field)
        {
            return Err("Le champ email sélectionné n'existe pas dans ce fichier.".to_string());
        }
    }

    random_sample_rows(&workbook.rows, count.clamp(1, 50))
        .into_iter()
        .map(|row| {
            let rendered = render_template_parts(&input.subject, &input.body_html, &row.values)?;
            let recipient = recipient_key
                .as_deref()
                .and_then(|field| row.values.get(field))
                .map(|value| value.trim().to_string())
                .filter(|value| !value.is_empty());

            Ok(PreviewSample {
                row_number: row.row_number,
                recipient,
                subject: rendered.subject,
                body_html: rendered.body_html,
                body_text: rendered.body_text,
            })
        })
        .collect()
}

#[tauri::command]
async fn send_campaign(
    app: AppHandle,
    request: SendCampaignRequest,
) -> Result<SendSummary, String> {
    CAMPAIGN_CANCEL.store(false, Ordering::SeqCst);
    tauri::async_runtime::spawn_blocking(move || send_campaign_inner(app, request))
        .await
        .map_err(|error| format!("Tâche d'envoi interrompue: {error}"))?
}

#[tauri::command]
fn cancel_campaign() {
    CAMPAIGN_CANCEL.store(true, Ordering::SeqCst);
}

#[tauri::command]
async fn download_and_install_update(app: AppHandle) -> Result<UpdateCheckResult, String> {
    let update = app
        .updater()
        .map_err(|error| format!("Updater non configuré: {error}"))?
        .check()
        .await
        .map_err(|error| format!("Impossible de vérifier les mises à jour: {error}"))?;

    if let Some(update) = update {
        update
            .download_and_install(|_, _| {}, || {})
            .await
            .map_err(|error| format!("Mise à jour téléchargée mais non installée: {error}"))?;
        Ok(UpdateCheckResult {
            status: "installed".to_string(),
            version: Some(update.version),
            message: "Mise à jour installée. Redémarrez l'application pour finaliser.".to_string(),
        })
    } else {
        Ok(UpdateCheckResult {
            status: "upToDate".to_string(),
            version: None,
            message: "Aucune mise à jour disponible.".to_string(),
        })
    }
}

fn send_campaign_inner(
    app: AppHandle,
    request: SendCampaignRequest,
) -> Result<SendSummary, String> {
    let templates = load_templates_inner(&app)?;
    let template = templates
        .iter()
        .find(|template| template.id == request.template_id)
        .ok_or_else(|| "Modèle introuvable.".to_string())?;
    let workbook = read_workbook_data(
        PathBuf::from(&request.excel_path),
        request.sheet_name.clone(),
    )?;

    if !workbook
        .fields
        .iter()
        .any(|field| field.key == request.recipient_field)
    {
        return Err("Le champ email sélectionné n'existe pas dans ce fichier.".to_string());
    }

    let rows: Vec<RecipientRow> = workbook
        .rows
        .into_iter()
        .take(request.limit.unwrap_or(usize::MAX))
        .collect();

    if rows.is_empty() {
        return Err("Aucune ligne exploitable dans le fichier Excel.".to_string());
    }

    let rate_limit = normalize_rate_limit(&request.rate_limit)?;

    if request.dry_run {
        return simulate_campaign(template, &rows, &request.recipient_field);
    }

    let config =
        load_smtp_config(app.clone())?.ok_or_else(|| "Configuration SMTP absente.".to_string())?;
    let stored_config = StoredSmtpConfig {
        host: config.host,
        port: config.port,
        security: config.security,
        username: config.username,
        from_email: config.from_email,
        from_name: config.from_name,
        reply_to: config.reply_to,
    };
    let password = resolve_password(None)?;
    let mailer = build_transport(&stored_config, password.as_deref())?;

    if let Some(test_email) = request
        .test_email
        .as_deref()
        .map(str::trim)
        .filter(|email| !email.is_empty())
    {
        let row = rows
            .first()
            .ok_or_else(|| "Aucune ligne disponible pour l'email de test.".to_string())?;
        let rendered = render_template_parts(&template.subject, &template.body_html, &row.values)?;
        send_email(
            &mailer,
            &stored_config,
            test_email,
            &rendered.subject,
            &rendered.body_html,
            &rendered.body_text,
        )?;
        return Ok(SendSummary {
            attempted: 1,
            sent: 1,
            skipped: 0,
            failed: Vec::new(),
            preview: Some(rendered),
        });
    }

    let total = rows.len();
    let mut summary = SendSummary {
        attempted: 0,
        sent: 0,
        skipped: 0,
        failed: Vec::new(),
        preview: None,
    };
    let mut delivery_attempts = 0;

    emit_progress(&app, "running", total, &summary, None, None, None);

    for row in rows {
        if CAMPAIGN_CANCEL.load(Ordering::Relaxed) {
            emit_progress(&app, "cancelled", total, &summary, None, None, None);
            return Ok(summary);
        }

        summary.attempted += 1;
        let recipient = row
            .values
            .get(&request.recipient_field)
            .map(|value| value.trim().to_string())
            .unwrap_or_default();

        if recipient.is_empty() {
            summary.skipped += 1;
            emit_progress(
                &app,
                "running",
                total,
                &summary,
                Some(row.row_number),
                None,
                None,
            );
            continue;
        }
        delivery_attempts += 1;
        wait_before_delivery_interruptible(&rate_limit, delivery_attempts);
        if CAMPAIGN_CANCEL.load(Ordering::Relaxed) {
            emit_progress(&app, "cancelled", total, &summary, None, None, None);
            return Ok(summary);
        }

        let outcome = render_template_parts(&template.subject, &template.body_html, &row.values)
            .and_then(|rendered| {
                if summary.preview.is_none() {
                    summary.preview = Some(rendered.clone());
                }
                send_email(
                    &mailer,
                    &stored_config,
                    &recipient,
                    &rendered.subject,
                    &rendered.body_html,
                    &rendered.body_text,
                )
            });

        let mut last_error: Option<String> = None;
        match outcome {
            Ok(()) => summary.sent += 1,
            Err(message) => {
                last_error = Some(message.clone());
                push_failure(
                    &mut summary.failed,
                    RowFailure {
                        row_number: row.row_number,
                        recipient: Some(recipient.clone()),
                        message,
                    },
                );
            }
        }

        emit_progress(
            &app,
            "running",
            total,
            &summary,
            Some(row.row_number),
            Some(recipient),
            last_error,
        );
    }

    emit_progress(&app, "done", total, &summary, None, None, None);
    Ok(summary)
}

fn emit_progress(
    app: &AppHandle,
    status: &'static str,
    total: usize,
    summary: &SendSummary,
    current_row: Option<usize>,
    current_recipient: Option<String>,
    last_error: Option<String>,
) {
    let payload = CampaignProgress {
        status,
        total,
        attempted: summary.attempted,
        sent: summary.sent,
        skipped: summary.skipped,
        failed: summary.failed.len(),
        current_row,
        current_recipient,
        last_error,
    };
    let _ = app.emit(PROGRESS_EVENT, payload);
}

fn wait_before_delivery_interruptible(rate_limit: &RateLimitRuntime, next_attempt: usize) {
    if next_attempt <= 1 {
        return;
    }
    let previous_attempts = next_attempt - 1;
    let mut pause = rate_limit.per_email_delay;
    if let Some(batch_size) = rate_limit.batch_size {
        if previous_attempts % batch_size == 0 {
            pause = pause.max(rate_limit.batch_pause);
        }
    }
    if pause.is_zero() {
        return;
    }

    let step = Duration::from_millis(100);
    let mut remaining = pause;
    while !remaining.is_zero() {
        if CAMPAIGN_CANCEL.load(Ordering::Relaxed) {
            return;
        }
        let slice = remaining.min(step);
        thread::sleep(slice);
        remaining = remaining.saturating_sub(slice);
    }
}

fn normalize_rate_limit(config: &RateLimitConfig) -> Result<RateLimitRuntime, String> {
    let max_per_minute_delay = match config.max_per_minute {
        Some(0) => return Err("Le maximum par minute doit être supérieur à 0.".to_string()),
        Some(value) if value > 600 => {
            return Err("Le maximum par minute ne peut pas dépasser 600.".to_string())
        }
        Some(value) => 60_000_u64.div_ceil(value),
        None => 0,
    };
    let min_delay_ms = config.min_delay_ms.unwrap_or(0);
    if min_delay_ms > 600_000 {
        return Err("La pause minimale ne peut pas dépasser 600000 ms.".to_string());
    }

    let batch_size = config.batch_size;
    if let Some(0) = batch_size {
        return Err("La taille de lot doit être supérieure à 0.".to_string());
    }
    let batch_pause_ms = config.batch_pause_ms.unwrap_or(0);
    if batch_pause_ms > 3_600_000 {
        return Err("La pause entre lots ne peut pas dépasser 3600 secondes.".to_string());
    }
    if batch_pause_ms > 0 && batch_size.is_none() {
        return Err("La pause entre lots nécessite une taille de lot.".to_string());
    }

    Ok(RateLimitRuntime {
        per_email_delay: Duration::from_millis(min_delay_ms.max(max_per_minute_delay)),
        batch_size,
        batch_pause: Duration::from_millis(batch_pause_ms),
    })
}

fn random_sample_rows(rows: &[RecipientRow], count: usize) -> Vec<RecipientRow> {
    let mut selected = rows.to_vec();
    if selected.len() <= 1 {
        selected.truncate(count.min(selected.len()));
        return selected;
    }

    let mut seed = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_nanos() as u64)
        .unwrap_or(0x9e37_79b9_7f4a_7c15);

    for index in (1..selected.len()).rev() {
        seed ^= seed << 13;
        seed ^= seed >> 7;
        seed ^= seed << 17;
        let swap_index = (seed as usize) % (index + 1);
        selected.swap(index, swap_index);
    }

    selected.truncate(count.min(selected.len()));
    selected
}

fn simulate_campaign(
    template: &MailingTemplate,
    rows: &[RecipientRow],
    recipient_field: &str,
) -> Result<SendSummary, String> {
    let mut summary = SendSummary {
        attempted: 0,
        sent: 0,
        skipped: 0,
        failed: Vec::new(),
        preview: None,
    };

    for row in rows {
        summary.attempted += 1;
        let recipient = row
            .values
            .get(recipient_field)
            .map(|value| value.trim().to_string())
            .unwrap_or_default();
        if recipient.is_empty() {
            summary.skipped += 1;
            continue;
        }

        match render_template_parts(&template.subject, &template.body_html, &row.values) {
            Ok(rendered) => {
                if summary.preview.is_none() {
                    summary.preview = Some(rendered);
                }
            }
            Err(message) => push_failure(
                &mut summary.failed,
                RowFailure {
                    row_number: row.row_number,
                    recipient: Some(recipient),
                    message,
                },
            ),
        }
    }

    Ok(summary)
}

fn push_failure(failures: &mut Vec<RowFailure>, failure: RowFailure) {
    if failures.len() < MAX_FAILURES {
        failures.push(failure);
    }
}

fn load_templates_inner(app: &AppHandle) -> Result<Vec<MailingTemplate>, String> {
    read_json_or_default(app, TEMPLATES_FILE, Vec::<MailingTemplate>::new())
}

fn validate_template(input: &TemplateInput) -> Result<(), String> {
    if input.name.trim().is_empty() {
        return Err("Le modèle doit avoir un nom.".to_string());
    }
    if input.subject.trim().is_empty() {
        return Err("Le sujet ne peut pas être vide.".to_string());
    }
    if input.body_html.trim().is_empty() {
        return Err("Le modèle doit contenir un corps HTML.".to_string());
    }
    Ok(())
}

fn normalize_smtp_input(input: &SmtpConfigInput) -> Result<StoredSmtpConfig, String> {
    let host = input.host.trim().to_string();
    if host.is_empty() {
        return Err("Le serveur SMTP est obligatoire.".to_string());
    }
    if input.port == 0 {
        return Err("Le port SMTP est invalide.".to_string());
    }
    input
        .from_email
        .trim()
        .parse::<lettre::Address>()
        .map_err(|error| format!("Email expéditeur invalide: {error}"))?;

    if let Some(reply_to) = input
        .reply_to
        .as_deref()
        .map(str::trim)
        .filter(|value| !value.is_empty())
    {
        reply_to
            .parse::<lettre::Address>()
            .map_err(|error| format!("Email de réponse invalide: {error}"))?;
    }

    Ok(StoredSmtpConfig {
        host,
        port: input.port,
        security: input.security.clone(),
        username: input.username.trim().to_string(),
        from_email: input.from_email.trim().to_string(),
        from_name: input.from_name.trim().to_string(),
        reply_to: input.reply_to.as_deref().and_then(trimmed_optional),
    })
}

fn build_transport(
    config: &StoredSmtpConfig,
    password: Option<&str>,
) -> Result<SmtpTransport, String> {
    let mut builder: SmtpTransportBuilder = match config.security {
        SmtpSecurity::Tls => SmtpTransport::relay(&config.host)
            .map_err(|error| format!("Configuration TLS SMTP invalide: {error}"))?,
        SmtpSecurity::StartTls => SmtpTransport::starttls_relay(&config.host)
            .map_err(|error| format!("Configuration STARTTLS SMTP invalide: {error}"))?,
        SmtpSecurity::None => SmtpTransport::builder_dangerous(&config.host),
    };

    builder = builder.port(config.port);
    if !config.username.trim().is_empty() {
        let Some(password) = password.filter(|value| !value.trim().is_empty()) else {
            return Err(
                "Mot de passe SMTP absent. Saisissez-le puis sauvegardez la configuration."
                    .to_string(),
            );
        };
        builder = builder.credentials(Credentials::new(
            config.username.clone(),
            password.to_string(),
        ));
    }

    Ok(builder.build())
}

fn send_email(
    mailer: &SmtpTransport,
    config: &StoredSmtpConfig,
    recipient: &str,
    subject: &str,
    body_html: &str,
    body_text: &str,
) -> Result<(), String> {
    let from_address = config
        .from_email
        .parse()
        .map_err(|error| format!("Email expéditeur invalide: {error}"))?;
    let from = Mailbox::new(trimmed_optional(&config.from_name), from_address);
    let to = recipient
        .parse::<Mailbox>()
        .map_err(|error| format!("Destinataire invalide: {error}"))?;

    let mut builder = Message::builder().from(from).to(to).subject(subject);
    if let Some(reply_to) = config.reply_to.as_deref().and_then(trimmed_optional) {
        let reply_to_mailbox = reply_to
            .parse::<Mailbox>()
            .map_err(|error| format!("Email de réponse invalide: {error}"))?;
        builder = builder.reply_to(reply_to_mailbox);
    }

    let plain = if body_text.trim().is_empty() {
        strip_html(body_html)
    } else {
        body_text.to_string()
    };
    let html = if body_html.trim().is_empty() {
        format!("<pre>{}</pre>", escape_html(&plain))
    } else {
        body_html.to_string()
    };

    let email = builder
        .multipart(
            MultiPart::alternative()
                .singlepart(SinglePart::plain(plain))
                .singlepart(
                    SinglePart::builder()
                        .header(header::ContentType::TEXT_HTML)
                        .body(html),
                ),
        )
        .map_err(|error| format!("Email impossible à construire: {error}"))?;

    mailer
        .send(&email)
        .map(|_| ())
        .map_err(|error| format!("Envoi SMTP refusé: {error}"))
}

fn render_template_parts(
    subject: &str,
    body_html: &str,
    row: &BTreeMap<String, String>,
) -> Result<RenderedPreview, String> {
    let rendered_html = render_html_template(body_html, row)?;
    let rendered_subject = sanitize_header_value(&render_text_template(subject, row)?);
    Ok(RenderedPreview {
        subject: rendered_subject,
        body_html: rendered_html.clone(),
        body_text: strip_html(&rendered_html),
    })
}

fn render_html_template(source: &str, row: &BTreeMap<String, String>) -> Result<String, String> {
    let mut handlebars = Handlebars::new();
    handlebars.set_strict_mode(true);
    handlebars
        .render_template(source, row)
        .map_err(format_template_error)
}

fn render_text_template(source: &str, row: &BTreeMap<String, String>) -> Result<String, String> {
    let mut handlebars = Handlebars::new();
    handlebars.set_strict_mode(true);
    handlebars.register_escape_fn(no_escape);
    handlebars
        .render_template(source, row)
        .map_err(format_template_error)
}

fn format_template_error(error: RenderError) -> String {
    match error.reason() {
        RenderErrorReason::MissingVariable(Some(field)) => {
            format!("Champ inconnu: {{{{{field}}}}}. Vérifiez le nom du champ dans le fichier Excel.")
        }
        RenderErrorReason::MissingVariable(None) => {
            "Champ inconnu dans le modèle. Vérifiez les champs entre {{ }}.".to_string()
        }
        RenderErrorReason::TemplateError(_) => {
            "Syntaxe du modèle invalide. Vérifiez les accolades {{ }}.".to_string()
        }
        RenderErrorReason::HelperNotFound(_) => {
            "Expression de modèle invalide. Utilisez uniquement des champs comme {{prenom}}."
                .to_string()
        }
        _ => "Modèle invalide. Vérifiez les champs entre {{ }}.".to_string(),
    }
}

fn sanitize_header_value(value: &str) -> String {
    value
        .chars()
        .filter(|character| *character != '\r' && *character != '\n')
        .collect::<String>()
        .trim()
        .to_string()
}

fn read_workbook_preview(
    path: PathBuf,
    sheet_name: Option<String>,
) -> Result<WorkbookPreview, String> {
    let file_name = path
        .file_name()
        .and_then(|file| file.to_str())
        .unwrap_or("fichier.xlsx")
        .to_string();
    let sheets = workbook_sheet_names(&path)?;
    let selected_sheet = choose_sheet(&sheets, sheet_name)?;
    let data = read_workbook_data(path.clone(), Some(selected_sheet.clone()))?;
    let total_rows = data.rows.len();
    Ok(WorkbookPreview {
        path: path.to_string_lossy().to_string(),
        file_name,
        sheet_name: selected_sheet,
        sheets,
        fields: data.fields,
        rows: data.rows.into_iter().take(20).collect(),
        total_rows,
    })
}

fn read_workbook_data(path: PathBuf, sheet_name: Option<String>) -> Result<WorkbookData, String> {
    let mut workbook = open_workbook_auto(&path)
        .map_err(|error| format!("Impossible d'ouvrir le fichier Excel: {error}"))?;
    let sheets = workbook.sheet_names().to_vec();
    let selected_sheet = choose_sheet(&sheets, sheet_name)?;
    let range = workbook
        .worksheet_range(&selected_sheet)
        .map_err(|error| format!("Feuille Excel illisible: {error}"))?;
    let mut rows_iter = range.rows();
    let header_row = rows_iter
        .next()
        .ok_or_else(|| "Le fichier Excel doit contenir une ligne d'en-tête.".to_string())?;
    let fields = build_fields(header_row)?;
    let mut rows = Vec::new();

    for (index, row) in rows_iter.enumerate() {
        let values: BTreeMap<String, String> = fields
            .iter()
            .enumerate()
            .map(|(column_index, field)| {
                let value = row
                    .get(column_index)
                    .map(cell_to_string)
                    .unwrap_or_default();
                (field.key.clone(), value)
            })
            .collect();

        if values.values().any(|value| !value.trim().is_empty()) {
            rows.push(RecipientRow {
                row_number: index + 2,
                values,
            });
        }
    }

    Ok(WorkbookData { fields, rows })
}

fn workbook_sheet_names(path: &PathBuf) -> Result<Vec<String>, String> {
    let workbook = open_workbook_auto(path)
        .map_err(|error| format!("Impossible d'ouvrir le fichier Excel: {error}"))?;
    Ok(workbook.sheet_names().to_vec())
}

fn choose_sheet(sheets: &[String], requested: Option<String>) -> Result<String, String> {
    if sheets.is_empty() {
        return Err("Le classeur ne contient aucune feuille.".to_string());
    }
    if let Some(sheet) = requested.filter(|sheet| sheets.iter().any(|candidate| candidate == sheet))
    {
        return Ok(sheet);
    }
    Ok(sheets[0].clone())
}

fn build_fields(header_row: &[Data]) -> Result<Vec<FieldDescriptor>, String> {
    let mut used = HashSet::new();
    let fields: Vec<FieldDescriptor> = header_row
        .iter()
        .enumerate()
        .filter_map(|(index, cell)| {
            let header = cell_to_string(cell);
            if header.trim().is_empty() {
                return None;
            }
            let key = slugify_header(&header, index, &mut used);
            Some(FieldDescriptor { header, key })
        })
        .collect();

    if fields.is_empty() {
        return Err("La première ligne Excel doit contenir au moins un en-tête.".to_string());
    }
    Ok(fields)
}

fn slugify_header(header: &str, index: usize, used: &mut HashSet<String>) -> String {
    let ascii = deunicode(header).to_lowercase();
    let cleaned: String = ascii
        .chars()
        .map(|character| {
            if character.is_ascii_alphanumeric() {
                character
            } else {
                '_'
            }
        })
        .collect();
    let mut key = cleaned
        .split('_')
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>()
        .join("_");

    if key.is_empty() {
        key = format!("champ_{}", index + 1);
    }
    if key
        .chars()
        .next()
        .is_some_and(|character| character.is_ascii_digit())
    {
        key = format!("champ_{key}");
    }

    let base = key.clone();
    let mut suffix = 2;
    while used.contains(&key) {
        key = format!("{base}_{suffix}");
        suffix += 1;
    }
    used.insert(key.clone());
    key
}

fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(value) => value.trim().to_string(),
        Data::Float(value) => {
            if value.fract() == 0.0 {
                format!("{value:.0}")
            } else {
                let text = value.to_string();
                text.trim_end_matches('0').trim_end_matches('.').to_string()
            }
        }
        Data::Int(value) => value.to_string(),
        Data::Bool(value) => value.to_string(),
        Data::Error(value) => format!("{value:?}"),
        Data::DateTime(value) => value.to_string(),
        Data::DateTimeIso(value) => value.to_string(),
        Data::DurationIso(value) => value.to_string(),
    }
}

fn app_data_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let dir = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Dossier de données inaccessible: {error}"))?;
    fs::create_dir_all(&dir)
        .map_err(|error| format!("Dossier de données impossible à créer: {error}"))?;
    Ok(dir)
}

fn data_path(app: &AppHandle, file: &str) -> Result<PathBuf, String> {
    Ok(app_data_dir(app)?.join(file))
}

fn read_json_or_default<T>(app: &AppHandle, file: &str, default: T) -> Result<T, String>
where
    T: DeserializeOwned,
{
    let path = data_path(app, file)?;
    if !path.exists() {
        return Ok(default);
    }
    let content = fs::read_to_string(&path)
        .map_err(|error| format!("Lecture impossible de {file}: {error}"))?;
    if content.trim().is_empty() {
        return Ok(default);
    }
    serde_json::from_str(&content).map_err(|error| format!("JSON invalide dans {file}: {error}"))
}

fn read_optional_json<T>(app: &AppHandle, file: &str) -> Result<Option<T>, String>
where
    T: DeserializeOwned,
{
    let path = data_path(app, file)?;
    if !path.exists() {
        return Ok(None);
    }
    let content = fs::read_to_string(&path)
        .map_err(|error| format!("Lecture impossible de {file}: {error}"))?;
    if content.trim().is_empty() {
        return Ok(None);
    }
    serde_json::from_str(&content)
        .map(Some)
        .map_err(|error| format!("JSON invalide dans {file}: {error}"))
}

fn write_json<T>(app: &AppHandle, file: &str, value: &T) -> Result<(), String>
where
    T: Serialize,
{
    let path = data_path(app, file)?;
    let content = serde_json::to_string_pretty(value)
        .map_err(|error| format!("Sérialisation impossible: {error}"))?;
    fs::write(&path, content).map_err(|error| format!("Écriture impossible de {file}: {error}"))
}

fn keyring_entry() -> Result<Entry, String> {
    Entry::new(KEYRING_SERVICE, KEYRING_ACCOUNT)
        .map_err(|error| format!("Trousseau système inaccessible: {error}"))
}

fn smtp_password_saved() -> bool {
    matches!(get_smtp_password(), Ok(Some(_)))
}

fn get_smtp_password() -> Result<Option<String>, String> {
    match keyring_entry()?.get_password() {
        Ok(password) => Ok(Some(password)),
        Err(keyring::Error::NoEntry) => Ok(None),
        Err(error) => Err(format!("Mot de passe SMTP inaccessible: {error}")),
    }
}

fn clear_smtp_password() -> Result<(), String> {
    match keyring_entry()?.delete_credential() {
        Ok(()) | Err(keyring::Error::NoEntry) => Ok(()),
        Err(error) => Err(format!("Mot de passe SMTP impossible à supprimer: {error}")),
    }
}

fn resolve_password(input_password: Option<String>) -> Result<Option<String>, String> {
    if let Some(password) = input_password
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
    {
        return Ok(Some(password));
    }
    get_smtp_password()
}

fn trimmed_optional(value: &str) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn strip_html(source: &str) -> String {
    let mut output = String::with_capacity(source.len());
    let mut in_tag = false;
    for character in source.chars() {
        match character {
            '<' => in_tag = true,
            '>' => {
                in_tag = false;
                output.push(' ');
            }
            _ if !in_tag => output.push(character),
            _ => {}
        }
    }
    output
        .replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

fn escape_html(source: &str) -> String {
    source
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_updater::Builder::new().build())
        .invoke_handler(tauri::generate_handler![
            load_templates,
            save_template,
            delete_template,
            load_smtp_config,
            save_smtp_config,
            send_smtp_test,
            pick_excel_file,
            preview_template,
            preview_template_samples,
            send_campaign,
            cancel_campaign,
            download_and_install_update
        ])
        .run(tauri::generate_context!())
        .expect("erreur pendant l'exécution de l'application");
}
