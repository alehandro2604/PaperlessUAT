// Generic key/value cache backed by Redis. Mirrors the interface of the
// Angular CacheService that previously used IndexedDB:
//   get / set / clear / entriesByPrefix / deleteByPrefix / clearAll
// Most keys are namespaced by the caller's Entra object id (req.userOid), so
// one user can never read or wipe another's personal entries (e.g. My Tasks).
// Shared across every signed-in account/browser (see SHARED_KEY_PREFIXES):
//   fc:files:folder:… / fc:files:lib:… / fc:tasks:lib:v7:… / fc:tasks:hr-folder:…
// Every read, write and delete of a shared key is checked against the
// caller's SharePoint permissions (permissionCheck.js), so a user only ever
// sees or changes snapshots of folders/lists they can open themselves.
const express = require('express');
const redisClient = require('./redisClient');
const { filterAllowedKeys, canAccessKey, getSiteId } = require('./permissionCheck');
const { getGraphClient } = require('./graphAuth');
const router = express.Router();

// Matches the old client-side PERSIST_MAX_AGE (7 days); Redis evicts for us.
const ENTRY_TTL_SECONDS = Number(process.env.CACHE_ENTRY_TTL_SECONDS || 7 * 24 * 60 * 60);

/**
 * Shared across every signed-in account/browser so one crawl warms Redis for all:
 * - fc:files:folder:…    All Files folder listings
 * - fc:files:lib:…       All Files library root listings (guarded by access to that library)
 * - fc:tasks:lib:v7:…    All Files folder Comments
 * - fc:tasks:hr-folder:… HR Files person Comments
 * NOT shared: everything else, e.g. fc:tasks:hr-user:… (My Tasks — per assignee)
 */
const SHARED_KEY_PREFIXES = [
  'fc:files:folder:',
  'fc:files:lib:',          // one root listing per document library, stored once for everyone
  'fc:tasks:lib:v7:',
  'fc:tasks:hr-folder:',
];

const isSharedKey = (key = '') =>
  SHARED_KEY_PREFIXES.some((prefix) => String(key).startsWith(prefix));

const ns = (req, key = '') => {
  if (key && isSharedKey(key)) return `shared:${key}`;
  return `cache:${req.userOid}:${key}`;
};

/** Redis SCAN patterns for a client prefix. clearAll (empty) is personal-only. */
const patternsForPrefix = (req, prefix = '') => {
  if (!prefix) return [`cache:${req.userOid}:*`];
  const patterns = [`cache:${req.userOid}:${prefix}*`];
  // Only touch shared keys when the prefix is in the shared family — never on
  // logout / owner-swap wipes that use the broad `fc:` prefix.
  if (isSharedKey(prefix)) {
    patterns.push(`shared:${prefix}*`);
  }
  return patterns;
};

const clientKeyFromRedis = (req, redisKey) => {
  if (redisKey.startsWith('shared:')) return redisKey.slice('shared:'.length);
  const personal = `cache:${req.userOid}:`;
  if (redisKey.startsWith(personal)) return redisKey.slice(personal.length);
  return redisKey;
};

/** Personal keys always pass; shared keys pass only if the user can open the source. */
const canUseKey = async (req, key) => !isSharedKey(key) || canAccessKey(req, key);

/** Drops shared Redis keys the caller has no SharePoint access to. */
async function visibleKeys(req, redisKeys) {
  const sharedClientKeys = redisKeys
    .filter((k) => k.startsWith('shared:'))
    .map((k) => clientKeyFromRedis(req, k));
  if (sharedClientKeys.length === 0) return redisKeys;
  const allowed = await filterAllowedKeys(req, sharedClientKeys);
  return redisKeys.filter((k) => !k.startsWith('shared:') || allowed.has(clientKeyFromRedis(req, k)));
}

// Express 4 does not forward async rejections to error handlers on its own.
const wrap = (fn) => (req, res) => fn(req, res).catch((err) => {
  console.error(`${req.method} ${req.originalUrl} failed:`, err.message || err);
  res.status(500).json({ error: 'Cache operation failed', detail: err.message });
});

async function keysByPattern(pattern) {
  const keys = [];
  for await (const key of redisClient.scanIterator({ MATCH: pattern, COUNT: 200 })) {
    // node-redis v5 yields batches (arrays); v4 yields single keys.
    if (Array.isArray(key)) keys.push(...key);
    else keys.push(key);
  }
  return keys;
}

async function keysForPrefix(req, prefix) {
  const sets = await Promise.all(patternsForPrefix(req, prefix).map(keysByPattern));
  return [...new Set(sets.flat())];
}

// Returns { data, timestamp } or JSON null when the key is absent.
// Shared keys: prefer shared:…, then fall back to this user's old personal copy
// so pre-migration caches still paint until the next write upgrades them.
router.get('/entry', wrap(async (req, res) => {
  const key = String(req.query.key || '');
  if (!(await canUseKey(req, key))) return res.status(403).json({ error: 'Access denied' });
  let raw = await redisClient.get(ns(req, key));
  if (!raw && key && isSharedKey(key)) {
    raw = await redisClient.get(`cache:${req.userOid}:${key}`);
  }
  res.json(raw ? JSON.parse(raw) : null);
}));

router.put('/entry', wrap(async (req, res) => {
  const { key, data } = req.body;
  if (!key) return res.status(400).json({ error: 'Missing key' });
  if (!(await canUseKey(req, key))) return res.status(403).json({ error: 'Access denied' });
  const entry = JSON.stringify({ data, timestamp: Date.now() });
  await redisClient.setEx(ns(req, key), ENTRY_TTL_SECONDS, entry);
  // Drop legacy personal copy once upgraded to shared.
  if (isSharedKey(key)) {
    await redisClient.del(`cache:${req.userOid}:${key}`);
  }
  res.json({ ok: true });
}));

router.delete('/entry', wrap(async (req, res) => {
  const key = String(req.query.key || '');
  if (!(await canUseKey(req, key))) return res.status(403).json({ error: 'Access denied' });
  await redisClient.del(ns(req, key));
  res.json({ ok: true });
}));

// Entries whose key starts with ?prefix= (personal; shared for folder-task families).
router.get('/entries', wrap(async (req, res) => {
  const keys = await visibleKeys(req, await keysForPrefix(req, String(req.query.prefix || '')));
  if (keys.length === 0) return res.json([]);
  const values = await redisClient.mGet(keys);
  const entries = keys
    .map((key, i) => ({
      key: clientKeyFromRedis(req, key),
      entry: values[i] ? JSON.parse(values[i]) : null,
    }))
    .filter((e) => e.entry !== null);
  res.json(entries);
}));

// Deletes by prefix. Empty prefix wipes this user's personal keys only (clearAll).
router.delete('/entries', wrap(async (req, res) => {
  const keys = await visibleKeys(req, await keysForPrefix(req, String(req.query.prefix || '')));
  if (keys.length > 0) await redisClient.del(keys);
  res.json({ ok: true, deleted: keys.length });
}));

// ─── Per-task To Do cache ──────────────────────────────────────────────────
// Replaces the single `fc:tasks:hr-user:v5` snapshot, which outgrew the upload
// limit and had to be re-sent whole after every change.
//   cache:<oid>:todo:task:<list>:<id>  one To Do task (personal: only its owner reads/writes it)
//   cache:<oid>:todo:index             Redis SET of the task keys in this user's To Do
//   cache:<oid>:todo:updated           when this user's To Do was last saved (ms)
// Personal on purpose: the browser uploads these rows, so they must never be shared.
// Keys sit under cache:<oid>: so logout's clearAll (DELETE /entries) wipes them too.

const TODO_TASK_TTL_SECONDS = ENTRY_TTL_SECONDS;
const MAX_TODO_UPSERTS_PER_CALL = 500;

const todoIndexKey = (req) => `cache:${req.userOid}:todo:index`;
const todoUpdatedKey = (req) => `cache:${req.userOid}:todo:updated`;
const todoTaskKey = (req, key) => `cache:${req.userOid}:todo:task:${key}`;

// "<listName>:<itemId>", the same key AppComponent.getHrTaskKey builds.
const isValidTodoTaskKey = (key) =>
  typeof key === 'string' && key.length > 0 && key.length <= 300 && !/[\r\n*?[\]]/.test(key);

// This user's To Do: { updatedAt, tasks: [{ key, data }] }. Expired tasks drop out of the index.
router.get('/todo', wrap(async (req, res) => {
  const [keys, updated] = await Promise.all([
    redisClient.sMembers(todoIndexKey(req)),
    redisClient.get(todoUpdatedKey(req)),
  ]);
  if (keys.length === 0) return res.json({ updatedAt: 0, tasks: [] });

  const values = await redisClient.mGet(keys.map((k) => todoTaskKey(req, k)));
  const expired = keys.filter((_, i) => !values[i]);
  if (expired.length) await redisClient.sRem(todoIndexKey(req), expired);

  res.json({
    updatedAt: Number(updated) || 0,
    tasks: keys.flatMap((key, i) => (values[i] ? [{ key, data: JSON.parse(values[i]) }] : [])),
  });
}));

// Save changed tasks and drop removed ones: { upsert: [{ key, data }], remove: [key] }.
router.put('/tasks', wrap(async (req, res) => {
  const upsert = Array.isArray(req.body?.upsert) ? req.body.upsert : [];
  const remove = Array.isArray(req.body?.remove) ? req.body.remove : [];
  if (upsert.length > MAX_TODO_UPSERTS_PER_CALL) {
    return res.status(413).json({ error: `At most ${MAX_TODO_UPSERTS_PER_CALL} tasks per call` });
  }

  const multi = redisClient.multi();
  let upserted = 0;
  let removed = 0;
  for (const item of upsert) {
    if (!isValidTodoTaskKey(item?.key) || item.data === undefined) continue;
    multi.setEx(todoTaskKey(req, item.key), TODO_TASK_TTL_SECONDS, JSON.stringify(item.data));
    multi.sAdd(todoIndexKey(req), item.key);
    upserted += 1;
  }
  for (const key of remove) {
    if (!isValidTodoTaskKey(key)) continue;
    multi.del(todoTaskKey(req, key));
    multi.sRem(todoIndexKey(req), key);
    removed += 1;
  }
  multi.setEx(todoUpdatedKey(req), TODO_TASK_TTL_SECONDS, String(Date.now()));
  multi.expire(todoIndexKey(req), TODO_TASK_TTL_SECONDS);
  await multi.exec();
  res.json({ ok: true, upserted, removed });
}));

// Forget this user's whole To Do cache (the old invalidate()).
router.delete('/todo', wrap(async (req, res) => {
  const keys = await redisClient.sMembers(todoIndexKey(req));
  await redisClient.del([...keys.map((k) => todoTaskKey(req, k)), todoIndexKey(req), todoUpdatedKey(req)]);
  res.json({ ok: true, deleted: keys.length });
}));

// Step 0 (read-only): does this list support Graph delta queries?
//   1st call: asks for a change token (token=latest) without listing the whole list.
//   Change an item in SharePoint, call again: should return only that item.
//   ?reset=1 starts over. Remove this route once step 3 is built.
router.get('/delta-test', wrap(async (req, res) => {
  const list = String(req.query.list || '').trim();
  if (!list) return res.status(400).json({ error: 'Pass ?list=<list name>' });

  const stateKey = `cache:${req.userOid}:todo:delta-test:${list}`;
  if (req.query.reset) await redisClient.del(stateKey);

  const client = await getGraphClient(req.userToken);   // acts as the signed-in user
  const siteId = await getSiteId(client);
  const saved = await redisClient.get(stateKey);
  let url = saved || `/sites/${siteId}/lists/${encodeURIComponent(list)}/items/delta?token=latest`;

  const changed = [];
  let deltaLink = null;
  let pages = 0;
  const started = Date.now();
  try {
    while (url && pages < 5) {
      const page = await client.api(url).get();
      for (const item of page.value ?? []) {
        changed.push({ id: item.id, modified: item.lastModifiedDateTime, deleted: !!item.deleted });
      }
      deltaLink = page['@odata.deltaLink'] ?? deltaLink;
      url = page['@odata.nextLink'] ?? null;
      pages += 1;
    }
  } catch (err) {
    return res.json({ list, supported: false, status: err.statusCode ?? err.code, message: err.message });
  }

  if (deltaLink) await redisClient.setEx(stateKey, 24 * 60 * 60, deltaLink);
  res.json({
    list,
    supported: !!deltaLink,
    firstRun: !saved,
    pages,
    ms: Date.now() - started,
    changedCount: changed.length,
    morePages: !!url,
    sample: changed.slice(0, 10),
  });
}));

module.exports = router;