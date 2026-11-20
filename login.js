#!/usr/bin/env node
import { OutlookSessionManager } from "./outlook-session.js";

const manager = new OutlookSessionManager();

console.log("Abrindo o Google Chrome para login no Outlook Web (outlook.office.com)...");
console.log("Faça o login com sua conta corporativa (e 2FA/Authenticator, se solicitado).");
console.log("Assim que sua Caixa de Entrada carregar, a janela fechará automaticamente.\n");

try {
  const session = await manager.interactiveLogin(240_000);
  console.log("✅ Sessão capturada e salva com sucesso!");
  console.log(`   Origem: ${session.origin}`);
  console.log(`   Expira em: ${new Date(session.expiresAt).toLocaleString()}`);
} catch (err) {
  console.error("❌ Falha ao capturar sessão:", err.message);
  process.exit(1);
}
