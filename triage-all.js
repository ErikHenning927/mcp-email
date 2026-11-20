import { OutlookSessionManager } from "./outlook-session.js";

const manager = new OutlookSessionManager();

export function classifyMessage(msg) {
  const subject = (msg.subject || "").toLowerCase();
  const fromName = (msg.fromName || "").toLowerCase();
  const fromEmail = (msg.fromEmail || "").toLowerCase();
  const preview = (msg.bodyPreview || "").toLowerCase();
  const text = `${subject} ${fromName} ${fromEmail} ${preview}`;

  // -------------------------------------------------------------
  // 1. CÓDIGOS DE LOGIN / VALIDAÇÃO / 2FA / PLATAFORMAS -> CODIGOS
  // -------------------------------------------------------------
  const isPix = text.includes("código pix") || text.includes("codigo pix");
  const isParental = text.includes("código dos pais") || text.includes("codigo dos pais");

  const isCode =
    !isPix &&
    !isParental &&
    (/c[oó]digo de login/i.test(text) ||
      /c[oó]digo de valida[cç][aã]o/i.test(text) ||
      /c[oó]digo de seguran[cç]a/i.test(text) ||
      /c[oó]digo de verifica[cç][aã]o/i.test(text) ||
      /c[oó]digo de confirma[cç][aã]o/i.test(text) ||
      /c[oó]digo de acesso/i.test(text) ||
      /c[oó]digo de autoriza[cç][aã]o/i.test(text) ||
      /c[oó]digo de uso [uú]nico/i.test(text) ||
      /seu c[oó]digo/i.test(text) ||
      /insira este c[oó]digo/i.test(text) ||
      /c[oó]digo [eé]:?/i.test(text) ||
      /verify your (email|account)/i.test(text) ||
      /verificação via e-mail/i.test(text) ||
      /novo início de sessão/i.test(text) ||
      /acesso de novo (computador|navegador|dispositivo)/i.test(text) ||
      /security code/i.test(text) ||
      /verification code/i.test(text) ||
      /login code/i.test(text) ||
      /access code/i.test(text) ||
      /one-time pass/i.test(text) ||
      /steam guard/i.test(text) ||
      /c[oó]digo para entrar/i.test(text) ||
      /seu pin/i.test(text) ||
      /use o c[oó]digo/i.test(text) ||
      (fromName.includes("spotify") && /c[oó]digo/i.test(text)) ||
      (fromName.includes("login caixa") && /c[oó]digo/i.test(text)) ||
      (fromName.includes("steam") && /acesso|c[oó]digo|sessão/i.test(text)) ||
      (fromName.includes("nvidia") && /verificação|conta/i.test(text)) ||
      (fromName.includes("acer") && /verify/i.test(text)) ||
      (fromName.includes("riot games") && subject.includes("código")) ||
      (fromName.includes("google") && subject.includes("código")) ||
      (fromName.includes("microsoft") && subject.includes("código")) ||
      (fromName.includes("github") && subject.includes("verification code")));

  if (isCode) {
    return {
      destination: "codigos",
      reason: "Código de login / validação / verificação de plataforma",
    };
  }

  // -------------------------------------------------------------
  // 2. ENTREVISTAS / RECRUTAMENTO / PROCESSOS SELETIVOS -> ENTREVISTA
  // -------------------------------------------------------------
  const isInterview =
    /entrevista/i.test(text) ||
    /interview/i.test(text) ||
    /processo seletivo/i.test(text) ||
    /recrutamento/i.test(text) ||
    /recruiter/i.test(text) ||
    /talent acquisition/i.test(text) ||
    /candidatura/i.test(text) ||
    /desafio t[eé]cnico/i.test(text) ||
    /etapa t[eé]cnica/i.test(text) ||
    /teste t[eé]cnico/i.test(text) ||
    /case t[eé]cnico/i.test(text) ||
    /bate[ -]?papo/i.test(text) ||
    /vaga de/i.test(text) ||
    text.includes("uds") ||
    text.includes("hubii") ||
    text.includes("filipe crespo") ||
    text.includes("hr interview") ||
    (text.includes("convite") &&
      (text.includes("devops") ||
        text.includes("automação") ||
        text.includes("desenvolvedor") ||
        text.includes("engenheiro") ||
        text.includes("python") ||
        text.includes("vaga")));

  if (isInterview) {
    return {
      destination: "entrevista",
      reason: "Entrevista / Processo seletivo / Vaga",
    };
  }

  // -------------------------------------------------------------
  // 3. ESSENCIAIS (FATURAS REAIS, ALERTAS CRÍTICOS) -> MANTER NA ENTRADA
  // -------------------------------------------------------------
  const isEssential =
    subject.includes("comunicado de débito") ||
    subject.includes("reaviso de vencimento") ||
    fromEmail.includes("cobranca.copel.com") ||
    fromEmail.includes("fatura.digital@faturamento.copel.com") ||
    subject.includes("fatura digital copel") ||
    subject.includes("alerta de segurança") ||
    subject.includes("sua conta do google foi recuperada") ||
    subject.includes("só para garantir, foi você mesmo?") ||
    subject.includes("developer program membership") ||
    subject.includes("aws account permanent closure") ||
    subject.includes("aws account") ||
    (fromName.includes("microsoft365") && subject.includes("final notice")) ||
    (fromName.includes("onedrive") && subject.includes("muitos arquivos foram excluídos"));

  if (isEssential) {
    return {
      destination: "inbox",
      reason: "Item essencial (fatura / segurança / conta / aviso crítico)",
    };
  }

  // -------------------------------------------------------------
  // 4. SPAM / VENDAS / PACOTES / PROMOÇÕES / ASSUNTOS ALEATÓRIOS -> GARBAGE
  // -------------------------------------------------------------
  const isGarbage =
    // Venda de pacotes e seguros
    fromEmail.includes("produtos.copel.com") ||
    fromName.includes("produtos e serviços") ||
    subject.includes("seu lar seguro") ||
    subject.includes("limpeza da caixa d’água") ||
    subject.includes("limpeza da caixa d'água") ||
    subject.includes("seguro protege casa") ||
    subject.includes("seguro protege vida") ||
    subject.includes("evite gastos inesperados com um seguro") ||
    subject.includes("a partir de r$") ||
    subject.includes("direto na conta de luz") ||
    // Prospecções BPO / Ummense
    fromName.includes("ummense") ||
    fromEmail.includes("ummense.com") ||
    // E-commerce marketing / Alibaba
    fromName.includes("alibaba") ||
    fromEmail.includes("alibaba.com") ||
    // Vagas / projetos em massa
    fromName.includes("freelancer") ||
    fromEmail.includes("freelancer.com") ||
    // Streaming e entretenimento marketing
    fromName.includes("spotify") ||
    fromEmail.includes("spotify.com") ||
    fromName.includes("hbo max") ||
    fromEmail.includes("hbomax.com") ||
    fromName.includes("max") ||
    fromEmail.includes("max.com") ||
    fromName.includes("globoplay") ||
    fromName.includes("paramount") ||
    fromName.includes("xbox") ||
    fromEmail.includes("xbox.com") ||
    // Marketing de produtos Tech & Newsletters
    fromEmail.includes("aws-marketing") ||
    fromEmail.includes("email.openai.com") ||
    fromName.includes("chatgpt") ||
    fromEmail.includes("clickup.com") ||
    fromName.includes("clickupdates") ||
    fromName.includes("microsoft rewards") ||
    fromEmail.includes("infoemail.microsoft.com") ||
    fromName.includes("twilio") ||
    // Termos de marketing agressivo
    subject.includes("black friday") ||
    subject.includes("black november") ||
    subject.includes("desconto") ||
    subject.includes("aproveite") ||
    subject.includes("assine") ||
    subject.includes("compre agora") ||
    subject.includes("oferta") ||
    subject.includes("audiolivro") ||
    subject.includes("ao vivo: um show") ||
    subject.includes("rock in rio") ||
    subject.includes("recompensas jogando") ||
    subject.includes("nós sentimos sua falta") ||
    subject.includes("what's new in") ||
    subject.includes("chegou o seu guia") ||
    subject.includes("as estreias mais quentes") ||
    preview.includes("com o plano pro") ||
    preview.includes("newsletter semanal") ||
    preview.includes("descadastre-se") ||
    preview.includes("unsubscribe");

  if (isGarbage) {
    return {
      destination: "garbage",
      reason: "Spam / Prospecção comercial / Pacote / Marketing",
    };
  }

  // Default: se não tiver certeza absoluta, mantém na caixa de entrada
  return {
    destination: "inbox",
    reason: "Manter na Entrada",
  };
}

export async function processEntireInbox({ unreadOnly = false } = {}) {
  console.log(`Iniciando varredura na Caixa de Entrada (${unreadOnly ? "Apenas Não Lidos" : "TODOS os e-mails"})...\n`);

  let totalScanned = 0;
  let totalMovedToGarbage = 0;
  let totalMovedToCodigos = 0;
  let totalMovedToEntrevista = 0;
  let totalKept = 0;
  let currentOffset = 0;

  // Ensure folders exist
  await manager.getFolderIdByName("codigos");
  await manager.getFolderIdByName("garbage");
  await manager.getFolderIdByName("entrevista");

  while (true) {
    const batch = await manager.listMessages({
      folder: "inbox",
      limit: 50,
      offset: currentOffset,
      unreadOnly,
    });

    if (batch.length === 0) {
      console.log("Fim da lista de e-mails.");
      break;
    }

    const toGarbage = [];
    const toCodigos = [];
    const toEntrevista = [];
    const toKeep = [];

    for (const msg of batch) {
      const decision = classifyMessage(msg);
      if (decision.destination === "codigos") {
        toCodigos.push({ ...msg, reason: decision.reason });
      } else if (decision.destination === "entrevista") {
        toEntrevista.push({ ...msg, reason: decision.reason });
      } else if (decision.destination === "garbage") {
        toGarbage.push({ ...msg, reason: decision.reason });
      } else {
        toKeep.push({ ...msg, reason: decision.reason });
      }
    }

    // Move to codigos
    if (toCodigos.length > 0) {
      console.log(`➡️  Movendo ${toCodigos.length} código(s) para 'codigos':`);
      toCodigos.forEach((m) => console.log(`   [CÓDIGO] [${m.fromName}] ${m.subject}`));
      await manager.moveMessagesToFolder(
        toCodigos.map((m) => m.id),
        "codigos"
      );
      totalMovedToCodigos += toCodigos.length;
    }

    // Move to entrevista
    if (toEntrevista.length > 0) {
      console.log(`💼 Movendo ${toEntrevista.length} e-mail(s) para 'entrevista':`);
      toEntrevista.forEach((m) => console.log(`   [ENTREVISTA] [${m.fromName}] ${m.subject}`));
      await manager.moveMessagesToFolder(
        toEntrevista.map((m) => m.id),
        "entrevista"
      );
      totalMovedToEntrevista += toEntrevista.length;
    }

    // Move to garbage
    if (toGarbage.length > 0) {
      console.log(`🗑️  Movendo ${toGarbage.length} e-mail(s) para 'garbage':`);
      toGarbage.forEach((m) => console.log(`   [GARBAGE] [${m.fromName}] ${m.subject}`));
      await manager.moveMessagesToFolder(
        toGarbage.map((m) => m.id),
        "garbage"
      );
      totalMovedToGarbage += toGarbage.length;
    }

    totalKept += toKeep.length;
    totalScanned += batch.length;

    currentOffset += toKeep.length;

    console.log(
      `Progresso: ${totalScanned} analisados | ${totalMovedToCodigos} códigos | ${totalMovedToEntrevista} entrevistas | ${totalMovedToGarbage} lixo | ${currentOffset} mantidos\n`
    );

    await new Promise((r) => setTimeout(r, 600));
  }

  console.log("================ RESUMO FINAL DA TRIAGEM ================");
  console.log(`Total analisado: ${totalScanned}`);
  console.log(`🔑 Códigos movidos para 'codigos': ${totalMovedToCodigos}`);
  console.log(`💼 Entrevistas movidas para 'entrevista': ${totalMovedToEntrevista}`);
  console.log(`🗑️  Spam/Pacotes movidos para 'garbage': ${totalMovedToGarbage}`);
  console.log(`✅ Itens legítimos mantidos na Entrada: ${totalKept}`);
  console.log("=========================================================\n");
}

export async function processAllUnread() {
  return processEntireInbox({ unreadOnly: false });
}

if (process.argv[1].endsWith("triage-all.js")) {
  const onlyUnread = process.argv.includes("--unread-only");
  processEntireInbox({ unreadOnly: onlyUnread }).catch((err) => {
    console.error("Erro durante a triagem:", err);
    process.exit(1);
  });
}
