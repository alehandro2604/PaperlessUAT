// Redis RAM report for capacity planning. Does not return cache contents.

function parseInfoMemory(infoText) {
  const map = {};
  for (const line of String(infoText || '').split(/\r?\n/)) {
    if (!line || line.startsWith('#')) continue;
    const idx = line.indexOf(':');
    if (idx === -1) continue;
    map[line.slice(0, idx)] = line.slice(idx + 1).trim();
  }
  return {
    usedMemoryBytes: Number(map.used_memory || 0),
    usedMemoryHuman: map.used_memory_human || null,
    usedMemoryPeakBytes: Number(map.used_memory_peak || 0),
    usedMemoryPeakHuman: map.used_memory_peak_human || null,
    maxMemoryBytes: Number(map.maxmemory || 0),
    maxMemoryHuman: map.maxmemory_human || null,
    fragmentationRatio: Number(map.mem_fragmentation_ratio || 0),
  };
}

function bytesToHuman(bytes) {
  const n = Number(bytes) || 0;
  if (n < 1024) return `${n} B`;
  const units = ['KB', 'MB', 'GB', 'TB'];
  let value = n / 1024;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${value.toFixed(value >= 10 ? 1 : 2)} ${units[unit]}`;
}

function classifyKey(key) {
  if (key.startsWith('shared:')) return { bucket: 'shared', userOid: null };
  const personal = key.match(/^cache:([0-9a-f-]{36}):/i);
  if (personal) return { bucket: 'personal', userOid: personal[1].toLowerCase() };
  if (key.startsWith('tasks:')) return { bucket: 'tasks', userOid: null };
  return { bucket: 'other', userOid: null };
}

async function scanAllKeys(redisClient) {
  const keys = [];
  for await (const key of redisClient.scanIterator({ MATCH: '*', COUNT: 200 })) {
    keys.push(key);
  }
  return keys;
}

async function memoryUsageBytes(redisClient, keys) {
  const sizes = [];
  const chunkSize = 100;
  for (let i = 0; i < keys.length; i += chunkSize) {
    const chunk = keys.slice(i, i + chunkSize);
    const results = await Promise.all(
      chunk.map(async (key) => {
        try {
          const usage = await redisClient.sendCommand(['MEMORY', 'USAGE', key]);
          return Number(usage) || 0;
        } catch {
          const raw = await redisClient.get(key);
          return raw ? Buffer.byteLength(raw) : 0;
        }
      })
    );
    sizes.push(...results);
  }
  return sizes;
}

function recommendRam({ usedMemoryBytes, personalUsers, avgPersonalBytes, sharedBytes }) {
  // Redis: 2x current use, rounded up to 1 GB steps (fragmentation + growth).
  const redisGb = Math.max(1, Math.ceil((usedMemoryBytes * 2) / (1024 ** 3)));
  // Node holds request JSON (limit 25mb) plus Graph responses; 1 GB is enough for typical UAT.
  const nodeGb = 1;
  const osGb = 2;
  const projected100Users =
    sharedBytes + avgPersonalBytes * Math.max(personalUsers, 100);
  return {
    redisGb,
    nodeGb,
    osHeadroomGb: osGb,
    serverVmGb: redisGb + nodeGb + osGb,
    note:
      'Buy RAM for the machine (or Azure Cache for Redis) that hosts Redis + Node, not for the user browser. Shared file/task crawls are stored once; My Tasks and other personal keys grow with each signed-in user and live up to 7 days.',
    if100UsersRedisGb: Math.max(1, Math.ceil((projected100Users * 2) / (1024 ** 3))),
  };
}

async function collectRamMetrics(redisClient) {
  const [infoText, keys] = await Promise.all([
    redisClient.info('memory'),
    scanAllKeys(redisClient),
  ]);
  const sizes = await memoryUsageBytes(redisClient, keys);

  const buckets = {
    shared: { keys: 0, bytes: 0 },
    personal: { keys: 0, bytes: 0 },
    tasks: { keys: 0, bytes: 0 },
    other: { keys: 0, bytes: 0 },
  };
  const byUser = new Map();

  keys.forEach((key, i) => {
    const bytes = sizes[i] || 0;
    const { bucket, userOid } = classifyKey(key);
    buckets[bucket].keys += 1;
    buckets[bucket].bytes += bytes;
    if (userOid) {
      const current = byUser.get(userOid) || { keys: 0, bytes: 0 };
      current.keys += 1;
      current.bytes += bytes;
      byUser.set(userOid, current);
    }
  });

  const users = [...byUser.entries()]
    .map(([userOid, stats]) => ({
      userOid,
      keys: stats.keys,
      bytes: stats.bytes,
      human: bytesToHuman(stats.bytes),
    }))
    .sort((a, b) => b.bytes - a.bytes);

  const personalUsers = users.length;
  const avgPersonalBytes = personalUsers
    ? Math.round(buckets.personal.bytes / personalUsers)
    : 0;
  const maxPersonalBytes = users.reduce((m, u) => Math.max(m, u.bytes), 0);
  const info = parseInfoMemory(infoText);
  const measuredBytes = sizes.reduce((sum, n) => sum + n, 0);

  return {
    generatedAt: new Date().toISOString(),
    redis: {
      ...info,
      keyCount: keys.length,
      measuredKeyBytes: measuredBytes,
      measuredKeyHuman: bytesToHuman(measuredBytes),
    },
    buckets: Object.fromEntries(
      Object.entries(buckets).map(([name, stats]) => [
        name,
        { ...stats, human: bytesToHuman(stats.bytes) },
      ])
    ),
    perUser: {
      userCount: personalUsers,
      averageBytes: avgPersonalBytes,
      averageHuman: bytesToHuman(avgPersonalBytes),
      maxBytes: maxPersonalBytes,
      maxHuman: bytesToHuman(maxPersonalBytes),
      users,
    },
    recommendation: recommendRam({
      usedMemoryBytes: info.usedMemoryBytes || measuredBytes,
      personalUsers,
      avgPersonalBytes,
      sharedBytes: buckets.shared.bytes + buckets.tasks.bytes,
    }),
  };
}

module.exports = { collectRamMetrics, bytesToHuman };
