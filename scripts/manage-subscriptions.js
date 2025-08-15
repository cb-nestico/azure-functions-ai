/**
 * CLI for managing Microsoft Graph webhook subscriptions.
 * Usage:
 *   node scripts/manage-subscriptions.js create --resource "/sites/{site-id}/drive/root" --notificationUrl "<url>" --expiration "YYYY-MM-DDTHH:MM:SSZ" --clientState "<state>"
 *   node scripts/manage-subscriptions.js renew --id "<subscription-id>" --expiration "YYYY-MM-DDTHH:MM:SSZ"
 *   node scripts/manage-subscriptions.js list
 *   node scripts/manage-subscriptions.js delete --id "<subscription-id>"
 */

const { ClientSecretCredential } = require('@azure/identity');
const fetch = require('node-fetch');

const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';

function parseArgs() {
  const args = {};
  const argv = process.argv.slice(2);
  args.cmd = argv[0];
  for (let i = 1; i < argv.length; i++) {
    if (argv[i].startsWith('--')) {
      const key = argv[i].slice(2);
      const val = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : 'true';
      args[key] = val;
    }
  }
  return args;
}

async function getAccessToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('Missing TENANT_ID, CLIENT_ID or CLIENT_SECRET environment variables');
  }
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

async function createSubscription(opts) {
  const token = await getAccessToken();
  const body = {
    changeType: opts.changeType || 'updated',
    notificationUrl: opts.notificationUrl,
    resource: opts.resource,
    expirationDateTime: opts.expiration, // <-- fix here
    clientState: opts.clientState || undefined
  };
  if (!body.notificationUrl || !body.resource || !body.expirationDateTime) { // <-- fix here
    throw new Error('create requires --notificationUrl, --resource, --expiration');
  }
  if (!body.clientState) delete body.clientState;
  const res = await graphRequest('POST', '/subscriptions', token, body);
  return res;
}

async function renewSubscription(id, expiration) {
  const token = await getAccessToken();
  if (!id || !expiration) throw new Error('renew requires --id and --expiration');
  const res = await graphRequest('PATCH', `/subscriptions/${encodeURIComponent(id)}`, token, { expirationDateTime: expiration });
  return res;
}

async function listSubscriptions() {
  const token = await getAccessToken();
  const res = await graphRequest('GET', '/subscriptions', token);
  return res;
}

async function deleteSubscription(id) {
  const token = await getAccessToken();
  if (!id) throw new Error('delete requires --id');
  await graphRequest('DELETE', `/subscriptions/${encodeURIComponent(id)}`, token);
  return { success: true };
}

function getDefaultExpirationDate() {
  const now = new Date();
  now.setMinutes(now.getMinutes() + 43200); // 43200 minutes = 30 days
  return now.toISOString().replace(/\.\d{3}Z$/, 'Z'); // Remove milliseconds for Graph API
}

(async function main() {
  try {
    
    const args = parseArgs();
    console.log('Parsed args:', args);

     // Auto-calculate expiration if not provided
    if ((args.cmd === 'create' || args.cmd === 'renew') && !args.expiration) {
      args.expiration = getDefaultExpirationDate();
      console.log(`Auto-set expiration to: ${args.expiration}`);
    }

    if (!args.cmd) {
      console.log('Usage: node manage-subscriptions.js <create|renew|list|delete> [options]');
      process.exit(1);
    }

    if (args.cmd === 'create') {
      const result = await createSubscription({
        changeType: args.changeType,
        notificationUrl: args.notificationUrl,
        resource: args.resource,
        expiration: args.expiration,
        clientState: args.clientState
      });
      console.log('Subscription created:', JSON.stringify(result, null, 2));
      return;
    }

    if (args.cmd === 'renew') {
      const result = await renewSubscription(args.id, args.expiration);
      console.log('Subscription renewed (response):', JSON.stringify(result, null, 2));
      return;
    }

    if (args.cmd === 'list') {
      const result = await listSubscriptions();
      console.log('Subscriptions:', JSON.stringify(result, null, 2));
      return;
    }

    if (args.cmd === 'delete') {
      const result = await deleteSubscription(args.id);
      console.log('Delete result:', result);
      return;
    }

    console.log('Unknown command');
    process.exit(1);
  } catch (err) {
    console.error('Error:', err.message || err);
    process.exit(1);
  }


})();