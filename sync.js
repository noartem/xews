import * as crypto from "crypto";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as zlib from "zlib";

import { FileAttachment } from "ews-javascript-api";

const SHARE_ROOT = path.join(os.homedir(), ".share", "xews");
const MAX_PROCESSED_EVENTS = 20000;
const BATCH_FILE_SUFFIX = ".json.gz";

function log(...args) {
  console.log(new Date().toISOString(), ...args);
}

function sha1(buf) {
  return crypto.createHash("sha1").update(buf).digest("hex");
}

function sha1Text(value) {
  return sha1(Buffer.from(value));
}

function sanitizeChannel(channel) {
  const trimmed = channel.trim();

  if (!trimmed) {
    throw new Error("Channel must not be empty");
  }

  if (!/^[a-zA-Z0-9._-]+$/.test(trimmed)) {
    throw new Error(
      "Channel may only contain letters, numbers, dot, underscore, and dash",
    );
  }

  return trimmed;
}

function getChannelSubject(channel) {
  return `xews-sync:${sanitizeChannel(channel)}`;
}

function getPaths(rootDir, channel) {
  const resolvedRoot = path.resolve(rootDir);
  const channelName = sanitizeChannel(channel);
  const rootHash = sha1Text(resolvedRoot);
  const channelDir = path.join(SHARE_ROOT, "channels", channelName);
  const replicaDir = path.join(SHARE_ROOT, "replicas", channelName);

  return {
    resolvedRoot,
    channelName,
    rootHash,
    objectDir: path.join(channelDir, "objects"),
    statePath: path.join(replicaDir, `${rootHash}.json`),
  };
}

function ensureDirs(paths) {
  fs.mkdirSync(paths.resolvedRoot, { recursive: true });
  fs.mkdirSync(paths.objectDir, { recursive: true });
  fs.mkdirSync(path.dirname(paths.statePath), { recursive: true });
}

function createEmptyState(paths) {
  return {
    version: 1,
    clientId: `client-${sha1Text(`${paths.channelName}:${paths.rootHash}`).slice(0, 12)}`,
    rootDir: paths.resolvedRoot,
    knownFiles: {},
    processedEvents: [],
  };
}

function loadState(paths) {
  if (!fs.existsSync(paths.statePath)) {
    return createEmptyState(paths);
  }

  const parsed = JSON.parse(fs.readFileSync(paths.statePath, "utf-8"));

  return {
    version: 1,
    clientId:
      typeof parsed.clientId === "string" && parsed.clientId
        ? parsed.clientId
        : createEmptyState(paths).clientId,
    rootDir: paths.resolvedRoot,
    knownFiles:
      parsed.knownFiles && typeof parsed.knownFiles === "object"
        ? parsed.knownFiles
        : {},
    processedEvents: Array.isArray(parsed.processedEvents)
      ? parsed.processedEvents.slice(-MAX_PROCESSED_EVENTS)
      : [],
  };
}

function saveState(paths, state) {
  fs.writeFileSync(
    paths.statePath,
    `${JSON.stringify(
      {
        version: state.version,
        clientId: state.clientId,
        rootDir: state.rootDir,
        knownFiles: state.knownFiles,
        processedEvents: state.processedEvents.slice(-MAX_PROCESSED_EVENTS),
      },
      null,
      2,
    )}\n`,
  );
}

function getObjectPath(paths, hash) {
  return path.join(paths.objectDir, hash);
}

function storeObject(paths, hash, buf) {
  const objectPath = getObjectPath(paths, hash);

  if (!fs.existsSync(objectPath)) {
    fs.writeFileSync(objectPath, buf);
  }
}

function loadObject(paths, hash) {
  if (!hash) return null;

  const objectPath = getObjectPath(paths, hash);

  if (!fs.existsSync(objectPath)) {
    return null;
  }

  return fs.readFileSync(objectPath);
}

function rememberProcessed(state, eventId) {
  if (state.processedEvents.includes(eventId)) {
    return;
  }

  state.processedEvents.push(eventId);

  if (state.processedEvents.length > MAX_PROCESSED_EVENTS) {
    state.processedEvents.splice(
      0,
      state.processedEvents.length - MAX_PROCESSED_EVENTS,
    );
  }
}

function createDelta(oldBuf, newBuf) {
  if (oldBuf.length === 0 || newBuf.length === 0) {
    return [];
  }

  const ops = [];
  const maxPrefix = Math.min(oldBuf.length, newBuf.length);
  let prefixLength = 0;

  while (
    prefixLength < maxPrefix &&
    oldBuf[prefixLength] === newBuf[prefixLength]
  ) {
    prefixLength += 1;
  }

  let oldSuffixIndex = oldBuf.length - 1;
  let newSuffixIndex = newBuf.length - 1;

  while (
    oldSuffixIndex >= prefixLength &&
    newSuffixIndex >= prefixLength &&
    oldBuf[oldSuffixIndex] === newBuf[newSuffixIndex]
  ) {
    oldSuffixIndex -= 1;
    newSuffixIndex -= 1;
  }

  if (prefixLength > 0) {
    ops.push({ type: "copy", offset: 0, length: prefixLength });
  }

  if (newSuffixIndex >= prefixLength) {
    ops.push({
      type: "data",
      data: newBuf.slice(prefixLength, newSuffixIndex + 1).toString("base64"),
    });
  }

  const suffixOffset = oldSuffixIndex + 1;
  const suffixLength = oldBuf.length - suffixOffset;

  if (suffixLength > 0) {
    ops.push({
      type: "copy",
      offset: suffixOffset,
      length: suffixLength,
    });
  }

  return ops;
}

function applyDelta(oldBuf, ops) {
  const result = [];

  for (const op of ops) {
    if (op.type === "copy") {
      result.push(oldBuf.slice(op.offset, op.offset + op.length));
      continue;
    }

    result.push(Buffer.from(op.data, "base64"));
  }

  return Buffer.concat(result);
}

function encodeFullPayload(buf) {
  return {
    mode: "full",
    contentBase64: buf.toString("base64"),
  };
}

function createUpsertPayload(relPath, baseHash, nextHash, oldBuf, newBuf) {
  if (!baseHash || !oldBuf || oldBuf.length === 0) {
    return {
      type: "upsert",
      path: relPath,
      baseHash,
      hash: nextHash,
      ...encodeFullPayload(newBuf),
    };
  }

  const delta = createDelta(oldBuf, newBuf);
  const deltaPayload = {
    type: "upsert",
    path: relPath,
    baseHash,
    hash: nextHash,
    mode: "delta",
    ops: delta,
  };
  const fullPayload = {
    type: "upsert",
    path: relPath,
    baseHash,
    hash: nextHash,
    ...encodeFullPayload(newBuf),
  };

  if (
    Buffer.byteLength(JSON.stringify(fullPayload)) <=
    Buffer.byteLength(JSON.stringify(deltaPayload))
  ) {
    return fullPayload;
  }

  return deltaPayload;
}

function validateRelativePath(relPath) {
  if (!relPath || relPath === ".") {
    throw new Error("Invalid relative path");
  }

  if (path.isAbsolute(relPath) || relPath.startsWith("..")) {
    throw new Error(`Path escapes sync root: ${relPath}`);
  }
}

function walkFiles(rootDir, currentDir = rootDir, files = []) {
  const entries = fs.readdirSync(currentDir, { withFileTypes: true });

  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
    const fullPath = path.join(currentDir, entry.name);

    if (entry.isDirectory()) {
      walkFiles(rootDir, fullPath, files);
      continue;
    }

    if (!entry.isFile()) {
      continue;
    }

    files.push(fullPath);
  }

  return files;
}

function scanDirectory(paths) {
  const files = walkFiles(paths.resolvedRoot);
  const snapshot = new Map();

  for (const filePath of files) {
    const relPath = path.relative(paths.resolvedRoot, filePath);
    validateRelativePath(relPath);
    const buf = fs.readFileSync(filePath);
    const hash = sha1(buf);
    storeObject(paths, hash, buf);
    snapshot.set(relPath, { hash, buf });
  }

  return snapshot;
}

function nextEventId(state) {
  return `${new Date().toISOString()}-${state.clientId}-${crypto.randomBytes(4).toString("hex")}`;
}

function encodeBatch(batch) {
  return zlib.gzipSync(Buffer.from(JSON.stringify(batch), "utf-8"));
}

function decodeBatch(buf) {
  return JSON.parse(zlib.gunzipSync(buf).toString("utf-8"));
}

async function sendBatch({ draft, batch }) {
  const name = `${batch.id}${BATCH_FILE_SUFFIX}`;
  const payload = encodeBatch(batch).toString("base64");

  draft.Attachments.AddFileAttachment(name, payload);
  await draft.Update(null);
}

async function pollEvents(draft) {
  if (!draft.HasAttachments) {
    return [];
  }

  const attachments = [];

  for (const att of draft.Attachments.Items) {
    if (!(att instanceof FileAttachment)) {
      continue;
    }

    await att.Load();

    if (!att.Base64Content) {
      continue;
    }

    const content = Buffer.from(att.Base64Content, "base64");
    let payloadEvents;

    if (att.Name.endsWith(BATCH_FILE_SUFFIX)) {
      const batch = decodeBatch(content);

      if (!batch || !Array.isArray(batch.events)) {
        throw new Error(`Invalid batch payload: ${att.Name}`);
      }

      payloadEvents = batch.events;
    } else if (att.Name.endsWith(".json")) {
      payloadEvents = [JSON.parse(content.toString("utf-8"))];
    } else {
      continue;
    }

    for (const event of payloadEvents) {
      if (!event || typeof event.id !== "string") {
        throw new Error(`Invalid event payload: ${att.Name}`);
      }
    }

    attachments.push({
      attachment: att,
      attachmentName: att.Name,
      events: payloadEvents,
    });
  }

  attachments.sort((a, b) => a.attachmentName.localeCompare(b.attachmentName));
  return attachments;
}

function resolveBaseBuffer(paths, state, event) {
  if (event.baseHash) {
    const stored = loadObject(paths, event.baseHash);

    if (stored) {
      return stored;
    }
  }

  const filePath = path.join(paths.resolvedRoot, event.path);

  if (fs.existsSync(filePath)) {
    return fs.readFileSync(filePath);
  }

  const knownHash = state.knownFiles[event.path];

  if (knownHash) {
    return loadObject(paths, knownHash) || Buffer.alloc(0);
  }

  return Buffer.alloc(0);
}

function materializeEventBuffer(paths, state, event) {
  if (event.mode === "full") {
    return Buffer.from(event.contentBase64, "base64");
  }

  if (event.mode === "delta") {
    const oldBuf = resolveBaseBuffer(paths, state, event);
    return applyDelta(oldBuf, event.ops);
  }

  throw new Error(`Unsupported event mode: ${event.mode}`);
}

function applyRemoteEvent(paths, state, event) {
  validateRelativePath(event.path);

  if (event.type === "delete") {
    const filePath = path.join(paths.resolvedRoot, event.path);

    if (fs.existsSync(filePath)) {
      fs.rmSync(filePath, { force: true });
    }

    delete state.knownFiles[event.path];
    rememberProcessed(state, event.id);
    log("[APPLY DELETE]", event.path);
    return;
  }

  if (event.type !== "upsert") {
    throw new Error(`Unsupported event type: ${event.type}`);
  }

  const nextBuf = materializeEventBuffer(paths, state, event);
  const actualHash = sha1(nextBuf);

  if (actualHash !== event.hash) {
    throw new Error(
      `Hash mismatch for ${event.path}: ${actualHash} !== ${event.hash}`,
    );
  }

  const filePath = path.join(paths.resolvedRoot, event.path);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, nextBuf);
  storeObject(paths, event.hash, nextBuf);
  state.knownFiles[event.path] = event.hash;
  rememberProcessed(state, event.id);
  log("[APPLY]", event.path, event.mode);
}

async function applyRemoteEvents({ paths, state, draft }) {
  const polled = await pollEvents(draft);
  let changed = false;
  const attachmentsToDelete = [];

  for (const entry of polled) {
    for (const event of entry.events) {
      if (state.processedEvents.includes(event.id)) {
        continue;
      }

      applyRemoteEvent(paths, state, event);
      changed = true;
    }

    attachmentsToDelete.push(entry.attachment);
  }

  if (changed) {
    saveState(paths, state);
  }

  if (attachmentsToDelete.length > 0) {
    for (const attachment of attachmentsToDelete) {
      draft.Attachments.Remove(attachment);
    }

    await draft.Update(null);
    log("[CLEANUP]", `removed=${attachmentsToDelete.length}`);
  }
}

function createLocalUpsertEvent({ paths, state, relPath, buf, hash }) {
  const baseHash = state.knownFiles[relPath] || null;
  const baseBuf = loadObject(paths, baseHash);
  return {
    version: 1,
    id: nextEventId(state),
    clientId: state.clientId,
    ...createUpsertPayload(relPath, baseHash, hash, baseBuf, buf),
  };
}

function createLocalDeleteEvent({ state, relPath }) {
  return {
    version: 1,
    id: nextEventId(state),
    clientId: state.clientId,
    type: "delete",
    path: relPath,
    baseHash: state.knownFiles[relPath] || null,
  };
}

async function sendLocalBatch({ draft, state, changes }) {
  if (changes.length === 0) {
    return false;
  }

  const batch = {
    version: 1,
    id: nextEventId(state),
    clientId: state.clientId,
    events: changes.map((change) => change.event),
  };

  await sendBatch({ draft, batch });

  for (const change of changes) {
    if (change.kind === "upsert") {
      storeObject(change.paths, change.hash, change.buf);
      state.knownFiles[change.relPath] = change.hash;
      log("[SEND]", change.relPath, change.event.mode);
    } else {
      delete state.knownFiles[change.relPath];
      log("[SEND DELETE]", change.relPath);
    }

    rememberProcessed(state, change.event.id);
  }

  return true;
}

async function syncLocalChanges({ paths, state, draft }) {
  const snapshot = scanDirectory(paths);
  const changes = [];

  for (const [relPath, entry] of snapshot) {
    if (state.knownFiles[relPath] === entry.hash) {
      continue;
    }

    changes.push({
      kind: "upsert",
      paths,
      relPath,
      buf: entry.buf,
      hash: entry.hash,
      event: createLocalUpsertEvent({
        paths,
        state,
        relPath,
        buf: entry.buf,
        hash: entry.hash,
      }),
    });
  }

  for (const relPath of Object.keys(state.knownFiles).sort()) {
    if (snapshot.has(relPath)) {
      continue;
    }

    changes.push({
      kind: "delete",
      relPath,
      event: createLocalDeleteEvent({ state, relPath }),
    });
  }

  const changed = await sendLocalBatch({ draft, state, changes });

  if (changed) {
    saveState(paths, state);
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function runSync({
  service,
  rootDir,
  channel,
  intervalMs,
  once,
  getOrCreateDraftBySubject,
}) {
  const paths = getPaths(rootDir, channel);
  ensureDirs(paths);

  const state = loadState(paths);
  const subject = getChannelSubject(paths.channelName);

  log("[SYNC START]", paths.resolvedRoot, subject, `client=${state.clientId}`);

  while (true) {
    const draft = await getOrCreateDraftBySubject(service, subject);

    await applyRemoteEvents({ paths, state, draft });
    await syncLocalChanges({ paths, state, draft });

    if (once) {
      break;
    }

    await sleep(intervalMs);
  }
}
