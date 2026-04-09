#!/usr/bin/env node

import * as os from "os";
import * as fs from "fs";
import * as path from "path";

import {
  ExchangeService,
  ExchangeVersion,
  WebCredentials,
  Uri,
  WellKnownFolderName,
  ItemView,
  EmailMessage,
  PropertySet,
  BasePropertySet,
  FileAttachment,
  DeleteMode,
  MessageBody,
} from "ews-javascript-api";

import { command, option, run } from "incur";


function loadConfig() {
  const configPath = path.join(os.homedir(), ".config", "xews", "auth.json");

  if (!fs.existsSync(configPath)) {
    throw new Error(`Config not found: ${configPath}`);
  }

  const raw = fs.readFileSync(configPath, "utf-8");
  const cfg = JSON.parse(raw);

  if (!cfg.email || !cfg.password || !cfg.url) {
    throw new Error("Invalid config: email/password/url required");
  }

  return cfg;
}

function createService() {
  const cfg = loadConfig();

  const service = new ExchangeService(ExchangeVersion.Exchange2016);

  service.Credentials = new WebCredentials(cfg.email, cfg.password);

  service.Url = new Uri(cfg.url);

  return service;
}

async function getOrCreateDraft(service) {
  const res = await service.FindItems(
    WellKnownFolderName.Drafts,
    new ItemView(1),
  );

  if (res.Items.length === 0) {
    const draft = new EmailMessage(service);
    draft.Subject = "Auto draft";
    draft.Body = new MessageBody("Generated");

    await draft.Save(WellKnownFolderName.Drafts);
    return draft;
  }

  const draft = res.Items[0];
  await draft.Load(new PropertySet(BasePropertySet.FirstClassProperties));
  return draft;
}

async function getFirstDraft(service) {
  const res = await service.FindItems(
    WellKnownFolderName.Drafts,
    new ItemView(1),
  );

  if (res.Items.length === 0) return null;

  const draft = res.Items[0];
  await draft.Load(new PropertySet(BasePropertySet.FirstClassProperties));
  return draft;
}

/* ================= COMMANDS ================= */

const upload = command({
  name: "upload",
  args: "<files...>",
  async run({ args }) {
    const service = createService();
    const draft = await getOrCreateDraft(service);

    for (const file of args.files) {
      if (!fs.existsSync(file)) {
        console.warn("Skip missing:", file);
        continue;
      }

      const content = fs.readFileSync(file);
      const name = path.basename(file);

      const att = draft.Attachments.AddFileAttachment(name);
      att.Content = content;

      console.log("Attached:", name);
    }

    await draft.Update(null);
    console.log("Done");
  },
});

const download = command({
  name: "download",
  options: {
    delete: option.boolean().default(false),
  },
  async run({ options }) {
    const service = createService();
    const draft = await getFirstDraft(service);

    if (!draft) {
      console.log("No drafts found");
      return;
    }

    if (!draft.HasAttachments) {
      console.log("No attachments");
      return;
    }

    for (const att of draft.Attachments.Items) {
      if (!(att instanceof FileAttachment)) continue;

      await att.Load();

      if (!att.Content) {
        console.warn("Empty:", att.Name);
        continue;
      }

      fs.writeFileSync(att.Name, Buffer.from(att.Content));
      console.log("Downloaded:", att.Name);
    }

    if (options.delete) {
      await draft.Delete(DeleteMode.MoveToDeletedItems);
      console.log("Draft deleted");
    }
  },
});

/* ================= ENTRY ================= */

run({
  name: "xews",
  commands: [upload, download],
});
