#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { OutlookSessionManager, loadCachedSession } from "./outlook-session.js";

const sessionManager = new OutlookSessionManager();

const server = new McpServer({
  name: "outlook-web-mcp",
  version: "1.0.0",
});

server.tool(
  "outlook_status",
  "Verifica se existe uma sessão autenticada e ativa do Outlook Web sem precisar de Azure App Registration.",
  {},
  async () => {
    const cached = loadCachedSession();
    if (cached) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                authenticated: true,
                origin: cached.origin,
                apiBase: cached.apiBase,
                expiresAt: new Date(cached.expiresAt).toISOString(),
                updatedAt: cached.updatedAt,
              },
              null,
              2
            ),
          },
        ],
      };
    }
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(
            {
              authenticated: false,
              message:
                "Nenhuma sessão ativa no cache. Chame `outlook_login` para abrir o Chrome e fazer login no Outlook Web.",
            },
            null,
            2
          ),
        },
      ],
    };
  }
);

server.tool(
  "outlook_login",
  "Abre uma janela do Google Chrome com perfil persistente em outlook.office.com para o usuário logar (suporta SSO e MFA da empresa) e captura automaticamente o token da sessão.",
  {
    timeoutSeconds: z
      .number()
      .optional()
      .default(240)
      .describe("Tempo máximo em segundos para concluir o login no navegador (padrão: 240s)."),
  },
  async ({ timeoutSeconds }) => {
    try {
      const session = await sessionManager.interactiveLogin(timeoutSeconds * 1000);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                success: true,
                message: "Login no Outlook Web concluído e sessão capturada com sucesso!",
                origin: session.origin,
                expiresAt: new Date(session.expiresAt).toISOString(),
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: `Erro no login: ${err.message}` }],
      };
    }
  }
);

server.tool(
  "outlook_list_folders",
  "Lista todas as pastas de e-mail da conta do Outlook Web (ex: Caixa de Entrada, garbage, Itens Enviados).",
  {},
  async () => {
    try {
      const folders = await sessionManager.listFolders();
      return {
        content: [{ type: "text", text: JSON.stringify(folders, null, 2) }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: `Erro ao listar pastas: ${err.message}` }],
      };
    }
  }
);

server.tool(
  "outlook_list_emails",
  "Lista os e-mails recentes do Outlook Web com remetente, assunto e prévia/texto do corpo para análise de contexto e triagem.",
  {
    folder: z
      .string()
      .optional()
      .default("inbox")
      .describe("ID da pasta ou 'inbox' para a Caixa de Entrada."),
    limit: z
      .number()
      .min(1)
      .max(100)
      .optional()
      .default(25)
      .describe("Quantidade máxima de e-mails para buscar (padrão: 25)."),
    unreadOnly: z
      .boolean()
      .optional()
      .default(false)
      .describe("Se true, retorna apenas e-mails não lidos."),
  },
  async ({ folder, limit, unreadOnly }) => {
    try {
      const messages = await sessionManager.listMessages({ folder, limit, unreadOnly });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                count: messages.length,
                messages,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: `Erro ao listar e-mails: ${err.message}` }],
      };
    }
  }
);

server.tool(
  "outlook_get_email",
  "Lê o conteúdo completo de um e-mail específico pelo seu ID caso o preview não seja suficiente para entender o contexto.",
  {
    messageId: z.string().describe("O ID do e-mail retornado por outlook_list_emails."),
  },
  async ({ messageId }) => {
    try {
      const details = await sessionManager.getMessageDetails(messageId);
      return {
        content: [{ type: "text", text: JSON.stringify(details, null, 2) }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: `Erro ao ler e-mail: ${err.message}` }],
      };
    }
  }
);

server.tool(
  "outlook_move_to_folder",
  "Move um ou mais e-mails (ex: spam, prospecção fria, vendas de pacotes, assuntos aleatórios) para uma pasta específica (padrão: 'garbage'). Cria a pasta automaticamente se ela não existir.",
  {
    messageIds: z
      .array(z.string())
      .min(1)
      .describe("Lista de IDs dos e-mails que devem ser movidos."),
    targetFolderName: z
      .string()
      .optional()
      .default("garbage")
      .describe("Nome da pasta de destino (padrão: 'garbage')."),
  },
  async ({ messageIds, targetFolderName }) => {
    try {
      const result = await sessionManager.moveMessagesToFolder(messageIds, targetFolderName);
      return {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: `Erro ao mover e-mails: ${err.message}` }],
      };
    }
  }
);

server.tool(
  "outlook_auto_triage",
  "Varre os e-mails não lidos da Caixa de Entrada e aplica as regras automaticamente: códigos de login/2FA/validação vão para 'codigos', e-mails de vendas/pacotes/marketing/spam vão para 'garbage', e e-mails essenciais (faturas reais, convites, segurança) são mantidos.",
  {},
  async () => {
    try {
      const { processAllUnread } = await import("./triage-all.js");
      await processAllUnread();
      return {
        content: [{ type: "text", text: "Triagem concluída com sucesso!" }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: `Erro na triagem: ${err.message}` }],
      };
    }
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error("Fatal error in outlook-web-mcp:", err);
  process.exit(1);
});
