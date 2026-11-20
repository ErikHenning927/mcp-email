#!/usr/bin/env node
import { OutlookSessionManager } from "./outlook-session.js";

const manager = new OutlookSessionManager();
const cmd = process.argv[2] || "status";

try {
  if (cmd === "login") {
    console.log("Abrindo Chrome para login no Outlook Web...");
    const s = await manager.interactiveLogin(240_000);
    console.log(JSON.stringify({ success: true, origin: s.origin, expiresAt: new Date(s.expiresAt).toISOString() }, null, 2));
  } else if (cmd === "folders") {
    const folders = await manager.listFolders();
    console.log(JSON.stringify(folders, null, 2));
  } else if (cmd === "list") {
    const limit = Number(process.argv[3] || 25);
    const unreadOnly = process.argv[4] === "unread";
    const messages = await manager.listMessages({ folder: "inbox", limit, unreadOnly });
    console.log(JSON.stringify({ count: messages.length, messages }, null, 2));
  } else if (cmd === "get") {
    const id = process.argv[3];
    const msg = await manager.getMessageDetails(id);
    console.log(JSON.stringify(msg, null, 2));
  } else if (cmd === "move") {
    const folder = process.argv[3] || "garbage";
    const ids = process.argv.slice(4);
    const res = await manager.moveMessagesToFolder(ids, folder);
    console.log(JSON.stringify(res, null, 2));
  } else {
    const s = await manager.getValidSession();
    console.log(JSON.stringify({ authenticated: true, origin: s.origin, expiresAt: new Date(s.expiresAt).toISOString() }, null, 2));
  }
} catch (err) {
  console.error(JSON.stringify({ error: err.message }, null, 2));
  process.exit(1);
}
