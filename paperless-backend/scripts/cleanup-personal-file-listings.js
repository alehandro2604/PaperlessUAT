// One-off cleanup after All Files listings moved to shared Redis keys.
//
// Removes the old per-user copies:
//   cache:<oid>:fc:files:all-root:*   merged all-libraries listing (no longer saved)
//   cache:<oid>:fc:files:lib:*        per-library listing, now stored once as shared:fc:files:lib:*
//
// The shared copies are rebuilt the next time any user opens that library, so deleting
// the old ones loses nothing. Dry run by default; nothing is deleted without --delete.
//
//   node scripts/cleanup-personal-file-listings.js                    dry run, .env
//   node scripts/cleanup-personal-file-listings.js --env .env.uat     dry run, UAT
//   node scripts/cleanup-personal-file-listings.js --delete           delete
const path = require('path');
const args = process.argv.slice(2);
const envFlag = args.indexOf('--env');
const envFile = envFlag > -1 ? args[envFlag + 1] : '.env';
require('dotenv').config({ path: path.resolve(__dirname, '..', envFile) });

const redisClient = require('../redisClient');

const PATTERNS = ['cache:*:fc:files:all-root:*', 'cache:*:fc:files:lib:*'];
const doDelete = args.includes('--delete');

function waitUntilReady(client) {
  if (client.isReady) return Promise.resolve();
  return new Promise((resolve, reject) => {
    client.once('ready', resolve);
    client.once('error', reject);
  });
}

async function main() {
  await waitUntilReady(redisClient);
  console.log(`${doDelete ? 'DELETE' : 'DRY RUN'} against ${process.env.REDIS_URL || 'redis://localhost:6379'}`);

  let keys = 0;
  let bytes = 0;
  const batch = [];
  for (const match of PATTERNS) {
    for await (const found of redisClient.scanIterator({ MATCH: match, COUNT: 500 })) {
      for (const key of [].concat(found)) {
        let size = 0;
        try { size = Number(await redisClient.sendCommand(['MEMORY', 'USAGE', key])) || 0; } catch { /* ignore */ }
        keys += 1;
        bytes += size;
        batch.push(key);
        if (batch.length >= 200 && doDelete) await redisClient.del(batch.splice(0));
      }
    }
  }
  if (doDelete && batch.length) await redisClient.del(batch.splice(0));

  console.log(`${doDelete ? 'Deleted' : 'Would delete'} ${keys} key(s), about ${(bytes / 1048576).toFixed(1)} MB.`);
  if (!doDelete && keys) console.log('Re-run with --delete to remove them.');
  await redisClient.quit();
}

main().catch((err) => {
  console.error('Cleanup failed:', err.message || err);
  process.exit(1);
});
