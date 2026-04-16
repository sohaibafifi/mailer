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
const pubkey = (process.env.TAURI_UPDATER_PUBKEY ?? dotenv.TAURI_UPDATER_PUBKEY ?? "").trim();

if (!pubkey) {
  console.error(
    "[inject-release-pubkey] TAURI_UPDATER_PUBKEY is empty. Set it in src-tauri/.env " +
      "(copy .env.example) or as an environment variable before running the release build.",
  );
  process.exit(1);
}

const config = JSON.parse(readFileSync(configPath, "utf8"));
config.plugins = config.plugins ?? {};
config.plugins.updater = config.plugins.updater ?? {};
config.plugins.updater.pubkey = pubkey;
writeFileSync(configPath, JSON.stringify(config, null, 2) + "\n");
console.log("[inject-release-pubkey] pubkey injected into tauri.release.conf.json");
