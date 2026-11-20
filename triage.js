#!/usr/bin/env node
import { OutlookSessionManager } from "./outlook-session.js";

const manager = new OutlookSessionManager();

// Known marketing/promotional/sales patterns or automated newsletters
function isSpamOrSales(msg) {
  const subject = (msg.subject || "").toLowerCase();
  const fromName = (msg.fromName || "").toLowerCase();
  const fromEmail = (msg.fromEmail || "").toLowerCase();
  const preview = (msg.bodyPreview || "").toLowerCase();

  // Security or account-critical emails MUST NEVER be moved
  const isSecurity =
    subject.includes("código de login") ||
    subject.includes("código de segurança") ||
    subject.includes("código de verificação") ||
    subject.includes("security code") ||
    subject.includes("verification code") ||
    subject.includes("redefinição de senha") ||
    subject.includes("fatura digital copel") || // personal utility bill
    fromEmail.includes("riotgames.com") ||
    fromEmail.includes("copel.com");

  if (isSecurity) return { isSpam: false, reason: "Segurança / Conta / Fatura essencial" };

  // Marketing, sales outreach, cold BPO pitch, streaming promos, general promotions
  if (
    fromName.includes("hbo max") ||
    fromName.includes("ummense") ||
    fromName.includes("freelancer") ||
    fromName.includes("xbox") ||
    fromEmail.includes("marketing") ||
    subject.includes("black friday") ||
    subject.includes("aproveite") ||
    subject.includes("novos projetos em") ||
    subject.includes("playback está pronto") ||
    subject.includes("sua equipe pode aproveitar") ||
    subject.includes("onboarding evita meses de retrabalho") ||
    preview.includes("com o plano pro") ||
    preview.includes("newsletter semanal")
  ) {
    return { isSpam: true, reason: "Prospecção / Vendas / Marketing promocional" };
  }

  return { isSpam: false, reason: "Relevante ou requer revisão manual" };
}

async function runTriage({ limit = 25, dryRun = false } = {}) {
  console.log(`Buscando os últimos ${limit} e-mails não lidos da Caixa de Entrada...`);
  const messages = await manager.listMessages({ folder: "inbox", limit, unreadOnly: true });

  console.log(`Encontrados ${messages.length} e-mails não lidos.\n`);

  const toMove = [];
  const kept = [];

  for (const msg of messages) {
    const analysis = isSpamOrSales(msg);
    if (analysis.isSpam) {
      toMove.push({ ...msg, reason: analysis.reason });
    } else {
      kept.push({ ...msg, reason: analysis.reason });
    }
  }

  console.log(`📋 Classificação:`);
  console.log(`   ➡️  Para mover para 'garbage': ${toMove.length}`);
  console.log(`   ✅  Manter na Caixa de Entrada: ${kept.length}\n`);

  if (toMove.length > 0) {
    console.log("E-mails identificados para 'garbage':");
    toMove.forEach((m, idx) => {
      console.log(`   ${idx + 1}. [${m.fromName}] ${m.subject} (${m.reason})`);
    });

    if (!dryRun) {
      console.log(`\nMovendo ${toMove.length} e-mail(s) para 'garbage'...`);
      const result = await manager.moveMessagesToFolder(
        toMove.map((m) => m.id),
        "garbage"
      );
      console.log(`✅ ${result.movedCount} e-mail(s) movidos com sucesso para a pasta 'garbage'!\n`);
    } else {
      console.log("\n[Dry-run] Nenhum e-mail foi movido.");
    }
  }

  return { toMove, kept };
}

const isDryRun = process.argv.includes("--dry-run");
const limitArg = Number(process.argv.find((a) => !isNaN(Number(a))) || 20);

runTriage({ limit: limitArg, dryRun: isDryRun }).catch((err) => {
  console.error("Erro na triagem:", err);
  process.exit(1);
});
