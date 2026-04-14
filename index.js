#!/usr/bin/env node

import * as os from "os";
import * as fs from "fs";
import * as path from "path";
import { spawnSync } from "child_process";

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
  SearchFilter,
  ItemSchema,
} from "ews-javascript-api";

import { Cli, z } from "incur";
import { runSync } from "./sync.js";

const CLI_VERSION = "1.0.0";
const CONFIG_PATH = path.join(os.homedir(), ".config", "xews", "auth.json");
const DEFAULT_CONFIG = `${JSON.stringify(
  {
    email: "user@example.com",
    password: "your-password",
    url: "https://exchange.example.com/EWS/Exchange.asmx",
  },
  null,
  2,
)}\n`;

function getEditor() {
  return (
    process.env.VISUAL || process.env.EDITOR || process.env.GIT_EDITOR || "vi"
  );
}

function ensureConfigFile() {
  fs.mkdirSync(path.dirname(CONFIG_PATH), { recursive: true });

  if (!fs.existsSync(CONFIG_PATH)) {
    fs.writeFileSync(CONFIG_PATH, DEFAULT_CONFIG);
  }

  return CONFIG_PATH;
}

function openInEditor(filePath) {
  const editor = getEditor();

  const result = spawnSync(`${editor} ${JSON.stringify(filePath)}`, {
    shell: true,
    stdio: "inherit",
  });

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error(`Editor exited with code ${result.status}`);
  }
}

function loadConfig() {
  const configPath = CONFIG_PATH;

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

async function getDraftBySubject(service, subject) {
  const filter = new SearchFilter.IsEqualTo(ItemSchema.Subject, subject);
  const res = await service.FindItems(
    WellKnownFolderName.Drafts,
    filter,
    new ItemView(1),
  );

  if (res.Items.length === 0) return null;

  const draft = res.Items[0];
  await draft.Load(new PropertySet(BasePropertySet.FirstClassProperties));
  return draft;
}

async function getOrCreateDraftBySubject(service, subject) {
  const draft = await getDraftBySubject(service, subject);

  if (draft) return draft;

  const created = new EmailMessage(service);
  created.Subject = subject;
  created.Body = new MessageBody(`Generated for ${subject}`);

  await created.Save(WellKnownFolderName.Drafts);
  await created.Load(new PropertySet(BasePropertySet.FirstClassProperties));

  return created;
}

/* ================= ENTRY ================= */

const cli = Cli.create("xews", {
  description:
    "Upload and download draft attachments via Exchange Web Services",
  version: CLI_VERSION,
});

cli.command("init", {
  description: "Create auth config if needed and open it in editor",
  async run() {
    const configPath = ensureConfigFile();
    openInEditor(configPath);
  },
});

cli.command("upload", {
  description: "Attach one or more files to a draft",
  options: z.object({
    file: z
      .array(z.string())
      .default([])
      .describe("File path to attach. Repeat --file for multiple files."),
  }),
  alias: { file: "f" },
  async run(c) {
    if (c.options.file.length === 0) {
      return c.error({
        code: "USAGE",
        message: "Provide at least one file with --file <path>",
      });
    }

    const service = createService();
    const draft = await getOrCreateDraft(service);

    for (const file of c.options.file) {
      if (!fs.existsSync(file)) {
        console.warn("Skip missing:", file);
        continue;
      }

      const content = fs.readFileSync(file);
      const name = path.basename(file);

      draft.Attachments.AddFileAttachment(name, content.toString("base64"));

      console.log("Attached:", name);
    }

    await draft.Update(null);
    console.log("Done");
  },
});

cli.command("download", {
  description: "Download attachments from the first draft",
  options: z.object({
    delete: z
      .boolean()
      .default(false)
      .describe("Delete the draft after download"),
  }),
  async run(c) {
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

      if (!att.Base64Content) {
        console.warn("Empty:", att.Name);
        continue;
      }

      fs.writeFileSync(att.Name, Buffer.from(att.Base64Content, "base64"));
      console.log("Downloaded:", att.Name);
    }

    if (c.options.delete) {
      await draft.Delete(DeleteMode.MoveToDeletedItems);
      console.log("Draft deleted");
    }
  },
});

cli.command("list", {
  description: "List attachments from the first draft",
  aliases: ["ls"],
  async run() {
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
      const size = typeof att.Size === "number" ? att.Size : "-";
      console.log(`${size}\t${att.Name}`);
    }
  },
});

cli.command("clear", {
  description: "Delete all attachments from the first draft",
  async run() {
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

    draft.Attachments.Clear();
    await draft.Update(null);

    console.log("Attachments cleared");
  },
});

cli.command("sync", {
  description: "Synchronize one folder through an EWS draft channel",
  options: z.object({
    dir: z.string().describe("Folder to synchronize"),
    channel: z
      .string()
      .default("default")
      .describe("Sync channel name stored in a dedicated draft subject"),
    interval: z
      .number()
      .int()
      .positive()
      .default(2000)
      .describe("Polling interval in milliseconds"),
    once: z.boolean().default(false).describe("Run one sync cycle and exit"),
  }),
  alias: {
    dir: "d",
    channel: "c",
  },
  async run(c) {
    const service = createService();

    await runSync({
      service,
      rootDir: c.options.dir,
      channel: c.options.channel,
      intervalMs: c.options.interval,
      once: c.options.once,
      getOrCreateDraftBySubject,
    });
  },
});

await cli.serve();
