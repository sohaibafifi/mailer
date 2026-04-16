#!/usr/bin/env node
// Injects TAURI_UPDATER_PUBKEY into src-tauri/tauri.release.conf.json.
// Reads from env first, falls back to src-tauri/.env. The committed file
// keeps an empty pubkey so the secret never lands in git.
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const configPath = join(root, "src-tauri", "tauri.release.conf.json");
const envPath = join(root, "src-tauri", ".env");

function loadDotEnv(path) {
  if (!existsSync(path)) return {};
  const out = {};
  for (const raw of readFileSync(path, "utf8").split("\n")) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    const equals = line.indexOf("=");
    if (equals === -1) continue;
    const key = line.slice(0, equals).trim();
    let value = line.slice(equals + 1).trim();
    if (value.startsWith("\"") && value.endsWith("\"")) value = value.slice(1, -1);
    if (value.startsWith("'") && value.endsWith("'")) value = value.slice(1, -1);
    out[key] = value;
  }
  return out;
}

const dotenv = loadDotEnv(envPath);
const rawPubkey = (process.env.TAURI_UPDATER_PUBKEY ?? dotenv.TAURI_UPDATER_PUBKEY ?? "").trim();

function fail(message) {
  console.error(
    `[inject-release-pubkey] ${message}`,
  );
  process.exit(1);
}

function normalizeLineEndings(value) {
  return value.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

function isEncodedPublicKeyLine(value) {
  const line = value.trim();
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(line)) return false;
  return Buffer.from(line, "base64").length === 42;
}

function parsePublicKeyFile(value) {
  const lines = normalizeLineEndings(value)
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.length < 2) {
    fail(
      "TAURI_UPDATER_PUBKEY must contain a full .pub file, the base64 of that file, " +
        "or the encoded public-key line from the .pub file.",
    );
  }
  if (!isEncodedPublicKeyLine(lines[1])) {
    fail("TAURI_UPDATER_PUBKEY does not contain a valid Tauri updater public key.");
  }
  return `${lines[0]}\n${lines[1]}`;
}

function encodePublicKeyFile(publicKeyFile) {
  return Buffer.from(`${publicKeyFile}\n`, "utf8").toString("base64");
}

function tryDecodeBase64Text(value) {
  try {
    return Buffer.from(value, "base64").toString("utf8");
  } catch {
    return null;
  }
}

function normalizeUpdaterPubkey(value) {
  if (!value) {
    fail(
      "TAURI_UPDATER_PUBKEY is empty. Set it in src-tauri/.env or as a GitHub Actions secret before running the release build.",
    );
  }

  const trimmed = normalizeLineEndings(value);
  if (trimmed.startsWith("untrusted comment:") || trimmed.includes("\n")) {
    return {
      value: encodePublicKeyFile(parsePublicKeyFile(trimmed)),
      format: "raw .pub file",
    };
  }

  const compact = trimmed.replace(/\s+/g, "");
  if (isEncodedPublicKeyLine(compact)) {
    return {
      value: encodePublicKeyFile(`untrusted comment: tauri public key\n${compact}`),
      format: "encoded-key line",
    };
  }

  const decoded = tryDecodeBase64Text(compact);
  if (decoded?.startsWith("untrusted comment:") || decoded?.includes("\n")) {
    return {
      value: encodePublicKeyFile(parsePublicKeyFile(decoded)),
      format: "base64 .pub file",
    };
  }
  if (decoded && isEncodedPublicKeyLine(decoded)) {
    return {
      value: encodePublicKeyFile(`untrusted comment: tauri public key\n${decoded.trim()}`),
      format: "base64 encoded-key line",
    };
  }

  fail(
    "TAURI_UPDATER_PUBKEY is not a valid updater public key. Use the full .pub file, " +
      "base64 of the .pub file, or the second line of the .pub file.",
  );
}

const pubkey = normalizeUpdaterPubkey(rawPubkey);
const config = JSON.parse(readFileSync(configPath, "utf8"));
config.plugins = config.plugins ?? {};
config.plugins.updater = config.plugins.updater ?? {};
config.plugins.updater.pubkey = pubkey.value;
writeFileSync(configPath, JSON.stringify(config, null, 2) + "\n");
console.log(`[inject-release-pubkey] pubkey injected into tauri.release.conf.json (${pubkey.format})`);
