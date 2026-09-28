require('dotenv').config();

const express = require('express');
const cors = require('cors');
const redisClient = require('./redisClient');
const { getGraphClient } = require('./graphAuth');
const { requireUser } = require('./authMiddleware');
const cacheRoutes = require('./cacheRoutes');

const app = express();

// Allow the Angular origins that call this API. In production prefer same-origin
// reverse proxy (then CORS is unused); CORS_ORIGINS is a comma-separated list.
const corsOrigins = (process.env.CORS_ORIGINS || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);
app.use(
  cors(
    corsOrigins.length
      ? { origin: corsOrigins, credentials: true }
      : undefined // default: reflect request origin (fine for local / same-network UAT)
  )
);
// Crawl snapshots cached by the Angular app can be several MB.
app.use(express.json({ limit: '25mb' }));

const PORT = process.env.PORT || 3000;
const CACHE_TTL_SECONDS = Number(process.env.CACHE_TTL_SECONDS || 300);
const SITE_PATH = process.env.SITE_PATH; // e.g. emoffice365.sharepoint.com:/sites/PaperlessUAT

// Health check — no auth needed, useful to confirm the server + Redis are up.
app.get('/api/health', async (req, res) => {
  try {
    const pong = await redisClient.ping();
    res.json({ server: 'ok', redis: pong });
  } catch (err) {
    res.status(500).json({ server: 'ok', redis: 'unreachable', error: err.message });
  }
});

// Everything below requires a validated user token from Angular
// (sets req.userToken for OBO and req.userOid for per-user cache keys).
app.use(requireUser);

// Generic per-user cache — Redis replacement for the app's IndexedDB store.
app.use('/api/cache', cacheRoutes);

/**
 * Cached per-person task lookup.
 * GET /api/tasks/:personEmail?list=<listNameOrId>
 * The list can also default via TASKS_LIST in .env.
 */
app.get('/api/tasks/:personEmail', async (req, res) => {
  try {
    const { personEmail } = req.params;
    const list = req.query.list || process.env.TASKS_LIST;
    if (!list) return res.status(400).json({ error: 'No list specified (use ?list= or TASKS_LIST env var)' });

    const cacheKey = `tasks:${list}:${personEmail.toLowerCase()}`;

    // 1. Check cache first
    const cached = await redisClient.get(cacheKey);
    if (cached) {
      res.set('X-Cache', 'HIT');
      return res.json(JSON.parse(cached));
    }

    // 2. Cache miss — fetch from Graph using the user's own token (OBO),
    //    so Graph itself enforces what this user is allowed to see.
    const graphClient = await getGraphClient(req.userToken);
    const freshData = await graphClient
      .api(`/sites/${SITE_PATH}:/lists/${list}/items`)
      .expand('fields')
      .filter(`fields/SubmittedByEmail eq '${personEmail.replace(/'/g, "''")}'`)
      .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
      .get();

    // 3. Cache it with an expiry
    await redisClient.setEx(cacheKey, CACHE_TTL_SECONDS, JSON.stringify(freshData));

    res.set('X-Cache', 'MISS');
    res.json(freshData);
  } catch (err) {
    console.error('GET /api/tasks failed:', err.message || err);
    const status = err.statusCode === 401 || err.errorCode === 'invalid_grant' ? 401 : 500;
    res.status(status).json({ error: 'Failed to fetch tasks', detail: err.message });
  }
});

/**
 * Invalidate the Redis read-through cache for a person/list so the next GET hits Graph.
 * DELETE /api/tasks/:personEmail?list=<listNameOrId>
 */
app.delete('/api/tasks/:personEmail', async (req, res) => {
  try {
    const { personEmail } = req.params;
    const list = req.query.list || process.env.TASKS_LIST;
    if (!list) return res.status(400).json({ error: 'No list specified (use ?list= or TASKS_LIST env var)' });

    const cacheKey = `tasks:${list}:${personEmail.toLowerCase()}`;
    const deleted = await redisClient.del(cacheKey);
    res.json({ ok: true, deleted });
  } catch (err) {
    console.error('DELETE /api/tasks failed:', err.message || err);
    res.status(500).json({ error: 'Failed to invalidate tasks cache', detail: err.message });
  }
});

app.listen(PORT, () => console.log(`Backend running on http://localhost:${PORT}`));
