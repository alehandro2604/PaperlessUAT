//What this will do is it will check if the user has access to the shared cache
//It will also cache the result for 15 minutes if the user has access
//It will also cache the result for 5 minutes if the user does not have access
const redisClient = require('./redisClient');
const { getGraphClient } = require('./graphAuth');

const ALLOW_TTL = 15 * 60;   // remember "yes" for 15 min
const DENY_TTL = 5 * 60;     // remember "no" for 5 min
const SITE_PATH = process.env.SITE_PATH;
const HR_LIBRARY = process.env.HR_LIBRARY_NAME || 'HRPersonal';
const norm = (s) => String(s ?? '').replace(/[\s_-]+/g, '').toLowerCase();

let siteIdPromise = null;
function getSiteId(client) {
  siteIdPromise ??= client.api(`/sites/${SITE_PATH}`).select('id').get()
    .then((s) => s.id)
    .catch((e) => { siteIdPromise = null; throw e; });
  return siteIdPromise;
}

// Libraries this user can see (also how we find the HR library's drive id)
async function getUserDrives(req, client, siteId) {
  const key = `perm:${req.userOid}:drives`;
  const cached = await redisClient.get(key);
  if (cached) return JSON.parse(cached);
  const page = await client.api(`/sites/${siteId}/drives`).select('id,name').get();
  const drives = (page.value ?? []).map((d) => ({ id: d.id, name: d.name }));
  await redisClient.setEx(key, ALLOW_TTL, JSON.stringify(drives));
  return drives;
}

// Key → which SharePoint object guards it. null = unknown shape → deny.
function describeKey(key, ctx) {
  let m = key.match(/^fc:files:folder:([^:]+):([^:]+)$/);
  if (m) return { resource: `item:${m[1]}:${m[2]}`, path: `/drives/${m[1]}/items/${m[2]}?$select=id` };

  // Library root listing: allowed only if the library is one of this user's drives, and
  // Graph confirms they can open its root (a library missing from their list is denied).
  m = key.match(/^fc:files:lib:[^:]+:(.+)$/);
  if (m) {
    const drive = (ctx.drives ?? []).find((d) => norm(d.name) === norm(m[1]));
    if (!drive) return null;
    return { resource: `drive:${drive.id}`, path: `/drives/${drive.id}/root?$select=id` };
  }

  m = key.match(/^fc:tasks:lib:v\d+:([^:]+):/);
  if (m) return { resource: `list:${m[1]}`, path: `/sites/${ctx.siteId}/lists/${encodeURIComponent(m[1])}?$select=id` };

  // `ids:` = cached { email, lookupId } for the person — same folder guards it.
  m = key.match(/^fc:tasks:hr-folder:(?:v\d+:|ids:)?(.+)$/);
  if (m && ctx.hrDriveId) {
    return { resource: `hr:${m[1]}`, path: `/drives/${ctx.hrDriveId}/root:/${encodeURIComponent(m[1])}?$select=id` };
  }
  return null;
}

// Returns the Set of keys this user is allowed to read.
async function filterAllowedKeys(req, keys) {
  const allowed = new Set();
  if (!keys.length) return allowed;

  let client, ctx;
  try {
    client = await getGraphClient(req.userToken);          // OBO: acts as the user
    const siteId = await getSiteId(client);
    const drives = await getUserDrives(req, client, siteId);
    ctx = { siteId, drives, hrDriveId: drives.find((d) => norm(d.name) === norm(HR_LIBRARY))?.id };
  } catch (err) {
    console.warn('[perm] cannot verify, denying shared cache:', err.message);
    return allowed;                                        // fail closed
  }

  // Group keys by the SharePoint object they depend on
  const groups = new Map();
  for (const key of keys) {
    const d = describeKey(key, ctx);
    if (!d) continue;
    if (!groups.has(d.resource)) groups.set(d.resource, { ...d, keys: [] });
    groups.get(d.resource).keys.push(key);
  }

  // 1) Answers we already know
  const resources = [...groups.keys()];
  const cached = resources.length
    ? await redisClient.mGet(resources.map((r) => `perm:${req.userOid}:${r}`))
    : [];
  const verdict = new Map();
  resources.forEach((r, i) => { if (cached[i]) verdict.set(r, cached[i] === '1'); });

  // 2) Ask Graph about the rest, 20 at a time with $batch
  const unknown = resources.filter((r) => !verdict.has(r));
  for (let i = 0; i < unknown.length; i += 20) {
    const chunk = unknown.slice(i, i + 20);
    try {
      const resp = await client.api('/$batch').post({
        requests: chunk.map((r, j) => ({ id: String(j), method: 'GET', url: groups.get(r).path })),
      });
      for (const r of resp.responses ?? []) {
        const resource = chunk[Number(r.id)];
        if (r.status >= 200 && r.status < 300) verdict.set(resource, true);
        else if ([401, 403, 404].includes(r.status)) verdict.set(resource, false);
        else continue;                                     // 429/5xx: deny now, don't cache
        const ok = verdict.get(resource);
        await redisClient.setEx(`perm:${req.userOid}:${resource}`, ok ? ALLOW_TTL : DENY_TTL, ok ? '1' : '0');
      }
    } catch (err) {
      console.warn('[perm] $batch failed:', err.message);
    }
  }

  for (const [resource, g] of groups) {
    if (verdict.get(resource) === true) g.keys.forEach((k) => allowed.add(k));
  }
  return allowed;
}

const canAccessKey = async (req, key) => (await filterAllowedKeys(req, [key])).has(key);

module.exports = { filterAllowedKeys, canAccessKey, getSiteId };