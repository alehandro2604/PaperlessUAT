const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const msalClient = new ConfidentialClientApplication(msalConfig);

/**
 * On-behalf-of flow: exchanges the user's token (issued for THIS API, i.e. the
 * api://<clientId>/access_as_user scope) for a Graph token representing the
 * same user. Graph then enforces that user's own SharePoint permissions.
 */
async function getGraphClient(userAccessToken) {
  const result = await msalClient.acquireTokenOnBehalfOf({
    oboAssertion: userAccessToken,
    scopes: ['https://graph.microsoft.com/.default'],
  });

  return Client.init({
    authProvider: (done) => done(null, result.accessToken),
  });
}

module.exports = { getGraphClient };
