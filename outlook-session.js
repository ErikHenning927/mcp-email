import { chromium } from "playwright-core";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";

const CONFIG_DIR = path.join(os.homedir(), ".config", "outlook-web-mcp");
const PROFILE_DIR = path.join(CONFIG_DIR, "chrome-profile");
const CACHE_FILE = path.join(CONFIG_DIR, "session-cache.json");
const CHROME_PATH = process.env.CHROME_PATH || "/usr/bin/google-chrome";

export function loadCachedSession() {
  try {
    if (fs.existsSync(CACHE_FILE)) {
      const data = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8"));
      if (data && data.authorization && data.anchor) {
        return data;
      }
    }
  } catch {
    // ignore
  }
  return null;
}

export function saveCachedSession(data) {
  if (!fs.existsSync(CONFIG_DIR)) {
    fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  }
  fs.writeFileSync(CACHE_FILE, JSON.stringify(data, null, 2), { mode: 0o600 });
}

export class OutlookSessionManager {
  constructor() {
    this.session = loadCachedSession();
    this.folderCache = null;
  }

  async getValidSession({ forceRefresh = false } = {}) {
    if (!forceRefresh && this.session) {
      return this.session;
    }

    const cached = loadCachedSession();
    if (!forceRefresh && cached) {
      this.session = cached;
      return cached;
    }

    // Refresh headless from profile
    return await this.refreshSessionHeadless();
  }

  async refreshSessionHeadless(timeoutMs = 35_000) {
    const context = await chromium.launchPersistentContext(PROFILE_DIR, {
      executablePath: CHROME_PATH,
      headless: true,
      args: ["--disable-blink-features=AutomationControlled"],
    });

    try {
      const page = context.pages()[0] || (await context.newPage());
      let captured = null;

      context.on("request", (req) => {
        const url = req.url();
        const h = req.headers();
        const auth = h["authorization"];
        if (auth && (url.includes("/owa/") && (url.includes("service.svc") || url.includes("startupdata")))) {
          if (!captured) {
            captured = {
              authorization: auth,
              anchor: h["x-anchormailbox"] || "",
              clientVersion: h["x-client-version"] || "20261112073.08",
              actionSource: h["x-owa-actionsource"] || "Mail",
              updatedAt: new Date().toISOString(),
            };
          }
        }
      });

      await page.goto("https://outlook.live.com/mail/0/", { waitUntil: "domcontentloaded" });

      const start = Date.now();
      while (Date.now() - start < timeoutMs) {
        if (captured) break;
        await new Promise((r) => setTimeout(r, 1000));
      }

      if (!captured) {
        throw new Error("Não foi possível capturar a sessão do Outlook Web.");
      }

      saveCachedSession(captured);
      this.session = captured;
      return captured;
    } finally {
      await context.close();
    }
  }

  async owaCall(actionName, payload, { retryOnAuthError = true } = {}) {
    const session = await this.getValidSession();
    const url = `https://outlook.live.com/owa/0/service.svc?action=${encodeURIComponent(actionName)}&app=Mail&n=1`;

    const headers = {
      authorization: session.authorization,
      "x-anchormailbox": session.anchor,
      action: actionName,
      "x-req-source": "Mail",
      "x-owa-actionsource": actionName,
      "content-type": "application/json; charset=utf-8",
      "x-owa-urlpostdata": encodeURIComponent(JSON.stringify(payload)),
      "x-client-version": session.clientVersion || "20261112073.08",
      "x-owa-hosted-ux": "false",
    };

    const res = await fetch(url, {
      method: "POST",
      headers,
    });

    if ((res.status === 401 || res.status === 440 || res.status === 500) && retryOnAuthError) {
      await this.getValidSession({ forceRefresh: true });
      return await this.owaCall(actionName, payload, { retryOnAuthError: false });
    }

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`Erro na chamada OWA ${actionName} (${res.status}): ${errText.slice(0, 300)}`);
    }

    return await res.json();
  }

  async listFolders() {
    const payload = {
      __type: "FindFolderJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "V2018_01_08",
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: {
            __type: "TimeZoneDefinitionType:#Exchange",
            Id: "E. South America Standard Time",
          },
        },
      },
      Body: {
        __type: "FindFolderRequest:#Exchange",
        FolderShape: {
          __type: "FolderResponseShape:#Exchange",
          BaseShape: "IdOnly",
          AdditionalProperties: [
            { __type: "PropertyUri:#Exchange", FieldURI: "DistinguishedFolderId" },
            { __type: "PropertyUri:#Exchange", FieldURI: "FolderDisplayName" },
            { __type: "PropertyUri:#Exchange", FieldURI: "TotalCount" },
            { __type: "PropertyUri:#Exchange", FieldURI: "UnreadCount" },
          ],
        },
        Paging: {
          __type: "IndexedPageView:#Exchange",
          BasePoint: "Beginning",
          Offset: 0,
          MaxEntriesReturned: 100,
        },
        ParentFolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "msgfolderroot" }],
        Traversal: "Deep",
      },
    };

    const res = await this.owaCall("FindFolder", payload);
    const folders = res?.Body?.ResponseMessages?.Items?.[0]?.RootFolder?.Folders || [];
    this.folderCache = folders.map((f) => ({
      id: f.FolderId?.Id,
      displayName: f.DisplayName,
      totalCount: f.TotalCount ?? 0,
      unreadCount: f.UnreadCount ?? 0,
      distinguishedFolderId: f.DistinguishedFolderId,
    }));
    return this.folderCache;
  }

  async getFolderIdByName(folderName) {
    const folders = await this.listFolders();
    const found = folders.find(
      (f) => f.displayName && f.displayName.toLowerCase() === folderName.toLowerCase()
    );
    if (found) return found.id;

    // Create folder if not found
    const createPayload = {
      __type: "CreateFolderJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "V2018_01_08",
      },
      Body: {
        __type: "CreateFolderRequest:#Exchange",
        ParentFolderId: {
          __type: "ParentFolderId:#Exchange",
          BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "msgfolderroot" },
        },
        Folders: [
          {
            __type: "Folder:#Exchange",
            DisplayName: folderName,
            FolderClass: "IPF.Note",
          },
        ],
      },
    };

    const res = await this.owaCall("CreateFolder", createPayload);
    const newFolder = res?.Body?.ResponseMessages?.Items?.[0]?.Folders?.[0];
    if (newFolder?.FolderId?.Id) {
      return newFolder.FolderId.Id;
    }
    throw new Error(`Pasta '${folderName}' não encontrada e não pôde ser criada.`);
  }

  async listMessages({ folder = "inbox", limit = 25, unreadOnly = false, offset = 0 } = {}) {
    let parentFolderIdObj = { __type: "DistinguishedFolderId:#Exchange", Id: "inbox" };
    if (folder && folder.toLowerCase() !== "inbox") {
      const folderId = await this.getFolderIdByName(folder);
      parentFolderIdObj = { __type: "FolderId:#Exchange", Id: folderId };
    }

    const payload = {
      __type: "FindItemJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "V2018_01_08",
      },
      Body: {
        __type: "FindItemRequest:#Exchange",
        ItemShape: {
          __type: "ItemResponseShape:#Exchange",
          BaseShape: "Default",
          AdditionalProperties: [{ __type: "PropertyUri:#Exchange", FieldURI: "Preview" }],
        },
        ParentFolderIds: [parentFolderIdObj],
        Traversal: "Shallow",
        Paging: {
          __type: "IndexedPageView:#Exchange",
          BasePoint: "Beginning",
          Offset: offset,
          MaxEntriesReturned: limit,
        },
      },
    };

    const res = await this.owaCall("FindItem", payload);
    let items = res?.Body?.ResponseMessages?.Items?.[0]?.RootFolder?.Items || [];
    if (unreadOnly) {
      items = items.filter((it) => it.IsRead === false);
    }

    return items.map((m) => ({
      id: m.ItemId?.Id,
      subject: m.Subject || "(Sem assunto)",
      fromName: m.From?.Mailbox?.Name || m.Sender?.Mailbox?.Name || "",
      fromEmail: m.From?.Mailbox?.EmailAddress || m.Sender?.Mailbox?.EmailAddress || "",
      receivedDateTime: m.DateTimeReceived,
      isRead: m.IsRead,
      bodyPreview: m.Preview || "",
      bodyText: m.Preview || "",
    }));
  }

  async getMessageDetails(messageId) {
    const payload = {
      __type: "GetItemJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "V2018_01_08",
      },
      Body: {
        __type: "GetItemRequest:#Exchange",
        ItemShape: {
          __type: "ItemResponseShape:#Exchange",
          BaseShape: "Default",
          BodyType: "Text",
        },
        ItemIds: [{ __type: "ItemId:#Exchange", Id: messageId }],
      },
    };

    const res = await this.owaCall("GetItem", payload);
    const item = res?.Body?.ResponseMessages?.Items?.[0]?.Items?.[0] || {};
    return {
      id: item.ItemId?.Id || messageId,
      subject: item.Subject || "(Sem assunto)",
      fromName: item.From?.Mailbox?.Name || "",
      fromEmail: item.From?.Mailbox?.EmailAddress || "",
      receivedDateTime: item.DateTimeReceived,
      isRead: item.IsRead,
      bodyText: item.Body?.Value || item.Preview || "",
    };
  }

  /**
   * Moves messages out of the source folder (Inbox) directly into targetFolder (e.g. 'garbage').
   * Guarantees items are removed from Inbox and appended to garbage without overwriting anything.
   */
  async moveMessagesToFolder(messageIds, targetFolderName = "garbage") {
    if (!messageIds || messageIds.length === 0) {
      return { movedCount: 0, results: [] };
    }

    const targetFolderId = await this.getFolderIdByName(targetFolderName);

    const payload = {
      __type: "MoveItemJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "V2018_01_08",
      },
      Body: {
        __type: "MoveItemRequest:#Exchange",
        ToFolderId: {
          __type: "TargetFolderId:#Exchange",
          BaseFolderId: {
            __type: "FolderId:#Exchange",
            Id: targetFolderId,
          },
        },
        ItemIds: messageIds.map((id) => ({
          __type: "ItemId:#Exchange",
          Id: id,
        })),
        ReturnNewItemIds: true,
      },
    };

    const res = await this.owaCall("MoveItem", payload);
    const responseItems = res?.Body?.ResponseMessages?.Items || [];

    const results = messageIds.map((id, index) => {
      const resp = responseItems[index];
      const isSuccess = resp?.ResponseClass === "Success" || resp?.ResponseCode === "NoError";
      return {
        id,
        status: isSuccess ? "moved" : "error",
        error: isSuccess ? null : resp?.MessageText || "Falha ao mover mensagem",
      };
    });

    return {
      targetFolder: targetFolderName,
      targetFolderId,
      movedCount: results.filter((r) => r.status === "moved").length,
      results,
    };
  }
}
