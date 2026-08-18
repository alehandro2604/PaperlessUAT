const { createClient } = require('redis');

const client = createClient({ url: process.env.REDIS_URL || 'redis://localhost:6379' });
client.on('error', (err) => console.error('Redis connection error:', err.message));

(async () => {
  await client.connect();
  console.log('Connected to Redis at', process.env.REDIS_URL || 'redis://localhost:6379');
})();

module.exports = client;
