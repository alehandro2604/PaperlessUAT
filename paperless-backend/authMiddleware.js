const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

// Entra ID signing keys (same endpoint serves v1 and v2 tokens), cached.
const jwks = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${process.env.TENANT_ID}/discovery/v2.0/keys`,
  cache: true,
  rateLimit: true,
});

function getKey(header, callback) {
  jwks.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

/**
 * Verifies the bearer token was issued by our tenant for THIS API
 * (the api://<clientId>/access_as_user scope exposed in Azure).
 * Sets req.userToken (raw, for OBO) and req.userOid (stable user id,
 * used to namespace Redis keys per user).
 */
function requireUser(req, res, next) {
  const authHeader = req.headers.authorization || '';
  const token = authHeader.replace(/^Bearer\s+/i, '');
  if (!token) return res.status(401).json({ error: 'No token provided' });

  const validAudiences = [`api://${process.env.CLIENT_ID}`, process.env.CLIENT_ID];

  jwt.verify(token, getKey, { audience: validAudiences }, (err, decoded) => {
    if (err) return res.status(401).json({ error: 'Invalid token', detail: err.message });
    if (decoded.tid !== process.env.TENANT_ID) {
      return res.status(401).json({ error: 'Token from wrong tenant' });
    }
    req.userToken = token;
    req.userOid = decoded.oid;
    next();
  });
}

module.exports = { requireUser };
