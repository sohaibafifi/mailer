export type FieldDescriptor = {
  header: string;
  key: string;
};

export type RecipientRow = {
  rowNumber: number;
  values: Record<string, string>;
};

export type WorkbookPreview = {
  path: string;
  fileName: string;
  sheetName: string;
  sheets: string[];
  fields: FieldDescriptor[];
  rows: RecipientRow[];
  totalRows: number;
};

export type MailingTemplate = {
  id: string;
  name: string;
  subject: string;
  bodyHtml: string;
  bodyText: string;
  createdAt: string;
  updatedAt: string;
};

export type TemplateInput = {
  id?: string | null;
  name: string;
  subject: string;
  bodyHtml: string;
  bodyText: string;
};

export type SmtpSecurity = "startTls" | "tls" | "none";

export type SmtpConfig = {
  host: string;
  port: number;
  security: SmtpSecurity;
  username: string;
  fromEmail: string;
  fromName: string;
  replyTo: string | null;
  passwordSaved: boolean;
};

export type SmtpConfigInput = Omit<SmtpConfig, "passwordSaved"> & {
  password: string | null;
  clearPassword: boolean;
};

export type RenderedPreview = {
  subject: string;
  bodyHtml: string;
  bodyText: string;
};

export type PreviewSample = RenderedPreview & {
  rowNumber: number;
  recipient: string | null;
};

export type RateLimitConfig = {
  maxPerMinute: number | null;
  minDelayMs: number | null;
  batchSize: number | null;
  batchPauseMs: number | null;
};

export type SendCampaignRequest = {
  templateId: string;
  excelPath: string;
  sheetName: string | null;
  recipientField: string;
  dryRun: boolean;
  testEmail: string | null;
  limit: number | null;
  rateLimit: RateLimitConfig;
};

export type RowFailure = {
  rowNumber: number;
  recipient: string | null;
  message: string;
};

export type SendSummary = {
  attempted: number;
  sent: number;
  skipped: number;
  failed: RowFailure[];
  preview: RenderedPreview | null;
};

export type UpdateCheckResult = {
  status: "upToDate" | "installed";
  version: string | null;
  message: string;
};

export type CampaignProgress = {
  status: "running" | "done" | "cancelled";
  total: number;
  attempted: number;
  sent: number;
  skipped: number;
  failed: number;
  currentRow: number | null;
  currentRecipient: string | null;
  lastError: string | null;
};
