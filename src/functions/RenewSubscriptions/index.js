const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const fetch = require('node-fetch');

const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';

async function getAccessToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('Missing TENANT_ID, CLIENT_ID or CLIENT_SECRET');
  const cred = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const scope = 'https://graph.microsoft.com/.default';
  const token = await cred.getToken(scope);
  if (!token || !token.token) throw new Error('Failed to acquire access token');
  return token.token;
}

async function graphRequest(method, path, token, body) {
  const url = `${GRAPH_ENDPOINT}${path}`;
  const opts = {
    method,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }
  };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(url, opts);
  const text = await res.text();
  let parsed;
  try { parsed = text ? JSON.parse(text) : null; } catch (e) { parsed = text; }
  if (!res.ok) {
    const err = new Error(`Graph request failed ${res.status} ${res.statusText}: ${typeof parsed === 'object' ? JSON.stringify(parsed) : parsed}`);
    err.status = res.status;
    throw err;
  }
  return parsed;
}

function getDefaultExpirationDate() {
  const now = new Date();
  now.setMinutes(now.getMinutes() + 43200); // 43200 minutes = 30 days
  return now.toISOString().replace(/\.\d{3}Z$/, 'Z'); // Remove milliseconds for Graph API
}

app.timer('RenewSubscriptions', {
  schedule: '0 0 * * * *', // Every day at midnight UTC
  handler: async (timerContext, context) => {
    context.log('🔄 RenewSubscriptions timer triggered');

    try {
      const token = await getAccessToken();
      const subs = await graphRequest('GET', '/subscriptions', token);

      if (!subs.value || !Array.isArray(subs.value) || subs.value.length === 0) {
        context.log('No subscriptions found to renew.');
        return;
      }

      const expiration = getDefaultExpirationDate();
      for (const sub of subs.value) {
        try {
          context.log(`Renewing subscription ${sub.id} (expires ${sub.expirationDateTime})`);
          const result = await graphRequest('PATCH', `/subscriptions/${encodeURIComponent(sub.id)}`, token, { expirationDateTime: expiration });
          context.log(`✅ Renewed ${sub.id}: new expiration ${result.expirationDateTime}`);
        } catch (err) {
          context.log(`❌ Failed to renew ${sub.id}: ${err.message || err}`);
        }
      }
    } catch (err) {
      context.log(`❌ Error in renewal process: ${err.message || err}`);
    }
  }
});