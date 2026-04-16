import { invoke } from "@tauri-apps/api/core";
import type {
  MailingTemplate,
  RenderedPreview,
  SendCampaignRequest,
  SendSummary,
  SmtpConfig,
  SmtpConfigInput,
  TemplateInput,
  UpdateCheckResult,
  WorkbookPreview,
} from "./types";

export function loadTemplates() {
  return invoke<MailingTemplate[]>("load_templates");
}

export function saveTemplate(input: TemplateInput) {
  return invoke<MailingTemplate[]>("save_template", { input });
}

export function deleteTemplate(id: string) {
  return invoke<MailingTemplate[]>("delete_template", { id });
}

export function loadSmtpConfig() {
  return invoke<SmtpConfig | null>("load_smtp_config");
}

export function saveSmtpConfig(input: SmtpConfigInput) {
  return invoke<SmtpConfig>("save_smtp_config", { input });
}

export function sendSmtpTest(input: SmtpConfigInput, email: string) {
  return invoke<string>("send_smtp_test", { input, email });
}

export function pickExcelFile() {
  return invoke<WorkbookPreview | null>("pick_excel_file");
}

export function previewTemplate(input: TemplateInput, row: Record<string, string>) {
  return invoke<RenderedPreview>("preview_template", { input, row });
}

export function sendCampaign(request: SendCampaignRequest) {
  return invoke<SendSummary>("send_campaign", { request });
}

export function downloadAndInstallUpdate() {
  return invoke<UpdateCheckResult>("download_and_install_update");
}

