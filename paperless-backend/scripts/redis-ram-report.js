require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const redisClient = require('../redisClient');
const { collectRamMetrics } = require('../ramMetrics');

function waitUntilReady(client) {
  if (client.isReady) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const onReady = () => {
      cleanup();
      resolve();
    };
    const onError = (err) => {
      cleanup();
      reject(err);
    };
    const cleanup = () => {
      client.off('ready', onReady);
      client.off('error', onError);
    };
    client.once('ready', onReady);
    client.once('error', onError);
  });
}

(async () => {
  try {
    await waitUntilReady(redisClient);
    const report = await collectRamMetrics(redisClient);
    console.log(JSON.stringify(report, null, 2));
    const rec = report.recommendation;
    console.error(
      `\nSummary: ${report.perUser.userCount} users, avg ${report.perUser.averageHuman} personal cache each, shared ${report.buckets.shared.human}.`
    );
    console.error(
      `Suggested: Redis ${rec.redisGb} GB + Node ${rec.nodeGb} GB + OS ${rec.osHeadroomGb} GB → VM ${rec.serverVmGb} GB (2x headroom on current Redis use).`
    );
  } catch (err) {
    console.error('Failed to collect Redis RAM metrics:', err.message || err);
    process.exitCode = 1;
  } finally {
    try {
      await redisClient.quit();
    } catch {
      process.exit();
    }
  }
})();
