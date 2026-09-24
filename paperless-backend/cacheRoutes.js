// Generic key/value cache backed by Redis. Mirrors the interface of the
// Angular CacheService that previously used IndexedDB:
//   get / set / clear / entriesByPrefix / deleteByPrefix / clearAll
// Most keys are namespaced by the caller's Entra object id (req.userOid), so
// one user can never read or wipe another's personal entries (e.g. My Tasks).
// Shared across every signed-in account/browser:
//   fc:files:… (Attachments, HR people list, All Files library listings)
//   fc:tasks:lib:v7:… / fc:tasks:hr-folder:… / fc:tasks:all-files-comments:…
// so any authenticated user reuses the same Redis snapshot.
const express = require('express');
const redisClient = require('./redisClient');

const router = express.Router();

// Matches the old client-side PERSIST_MAX_AGE (7 days); Redis evicts for us.
const ENTRY_TTL_SECONDS = Number(process.env.CACHE_ENTRY_TTL_SECONDS || 7 * 24 * 60 * 60);

/**
 * Shared across every signed-in account/browser so one crawl warms Redis for all:
 * - fc:files:…          Attachments + HR people list + All Files library listings
 * - fc:tasks:lib:v7:…   All Files folder Comments
 * - fc:tasks:hr-folder:… HR Files person Comments
 * - fc:tasks:all-files-comments:… All Files comment-search snapshot
 * NOT shared: fc:tasks:hr-user:… (My Tasks — per assignee)
 */

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

// Express 4 does not forward async rejections to error handlers on its own.
const wrap = (fn) => (req, res) => fn(req, res).catch((err) => {
  console.error(`${req.method} ${req.originalUrl} failed:`, err.message || err);
  res.status(500).json({ error: 'Cache operation failed', detail: err.message });
});

const { filterAllowedKeys, canAccessKey } = require('./permissionCheck');

const SHARED_KEY_PREFIXES = [
  'fc:files:folder:',
  'fc:tasks:lib:v7:',
  'fc:tasks:hr-folder:',
];

router.get('/entry', wrap(async (req, res) => {
  const key = String(req.query.key || '');
  const allowed = await filterAllowedKeys(req, [key]);
  if (!allowed.has(key)) return res.status(403).json({ error: 'Access denied' });
  const raw = await redisClient.get(ns(req, key));
  res.json(raw ? JSON.parse(raw) : null);
}));


const keys = await keysForPrefix(req, prefix);
const sharedClientKeys = keys.filter((k) => k.startsWith('shared:')).map((k) => clientKeyFromRedis(req, k));
const allowed = await filterAllowedKeys(req, sharedClientKeys);
const visible = keys.filter((k) => !k.startsWith('shared:') || allowed.has(clientKeyFromRedis(req, k)));
// use `visible` instead of `keys` for mGet + mapping


isSharedKey = (key) => SHARED_KEY_PREFIXES.some((prefix) => String(key).startsWith(prefix));

async function keysByPattern(pattern) {
  const keys = [];
  for await (const key of redisClient.scanIterator({ MATCH: pattern, COUNT: 200 })) {
    keys.push(key);
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
  let raw = await redisClient.get(ns(req, key));
  if (!raw && key && isSharedKey(key)) {
    raw = await redisClient.get(`cache:${req.userOid}:${key}`);
  }
  res.json(raw ? JSON.parse(raw) : null);
}));

router.put('/entry', wrap(async (req, res) => {
  const { key, data } = req.body;
  if (!key) return res.status(400).json({ error: 'Missing key' });
  const entry = JSON.stringify({ data, timestamp: Date.now() });
  const redisKey = ns(req, key);
  await redisClient.setEx(redisKey, ENTRY_TTL_SECONDS, entry);
  // Drop legacy personal copy once upgraded to shared.
  if (isSharedKey(key)) {
    await redisClient.del(`cache:${req.userOid}:${key}`);
  }
  res.json({ ok: true });
}));

router.delete('/entry', wrap(async (req, res) => {
  await redisClient.del(ns(req, String(req.query.key || '')));
  res.json({ ok: true });
}));

// Entries whose key starts with ?prefix= (personal; shared for folder-task families).
router.get('/entries', wrap(async (req, res) => {
  const prefix = String(req.query.prefix || '');
  const keys = await keysForPrefix(req, prefix);
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
  const keys = await keysForPrefix(req, String(req.query.prefix || ''));
  if (keys.length > 0) await redisClient.del(keys);
  res.json({ ok: true, deleted: keys.length });
}));

module.exports = router;