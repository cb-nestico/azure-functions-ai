/**
 * SharePointWebhook - classic Azure Functions handler (module.exports)
 * - Handles Graph webhook validation, notification processing, and forwards .vtt files.
 */

const DEFAULT_LOG = console;

module.exports = async function (context, req) {
  context.log = context.log || DEFAULT_LOG;
  context.log('🔔 SharePointWebhook invoked');

  // Subscription validation (Graph sends validationToken)
  const validationToken = safeGetValidationToken(req);
  if (validationToken) {
    context.log('🧾 Validation token received, replying to Graph');
    // Graph expects plain text body containing the token
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'text/plain' },
      body: validationToken
    };
    return;
  }

  // Parse payload
  const body = req && req.body;
  if (!body) {
    context.log('❌ No payload received');
    context.res = { status: 400, body: 'No payload' };
    return;
  }
  context.log('📬 Notification payload received');

  // Lazy-load fetch
  let fetchFn = globalThis.fetch;
  if (!fetchFn) {
    try { fetchFn = require('node-fetch'); } catch (e) { context.log('⚠️ node-fetch not installed; internal calls skipped'); }
  }

  // Lazy-load Graph libs
  let ClientSecretCredential = null;
  let GraphClient = null;
  try { ClientSecretCredential = require('@azure/identity').ClientSecretCredential; } catch (e) { /* not available */ }
  try { GraphClient = require('@microsoft/microsoft-graph-client').Client; } catch (e) { /* not available */ }

  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const sharepointDriveId = process.env.SHAREPOINT_DRIVE_ID;

  // Initialize Graph client if possible
  let graphClient = null;
  if (ClientSecretCredential && GraphClient && tenantId && clientId && clientSecret) {
    try {
      const cred = new ClientSecretCredential(tenantId, clientId, clientSecret);
      const authProvider = {
        getAccessToken: async () => {
          const t = await cred.getToken('https://graph.microsoft.com/.default');
          return t ? t.token : null;
        }
      };
      graphClient = GraphClient.initWithMiddleware ? GraphClient.initWithMiddleware({ authProvider }) : null;
      context.log('✅ Graph client initialized');
    } catch (err) {
      context.log('❌ Failed to initialize Graph client:', err?.message || err);
      graphClient = null;
    }
  } else {
    context.log('ℹ️ Graph client not initialized (missing libs or env vars)');
  }

  // Internal ProcessVttFile endpoint resolution
  const processUrlOverride = process.env.PROCESS_VTT_URL;
  const websiteHostname = process.env.WEBSITE_HOSTNAME; // present when running in Azure
  const internalBase = processUrlOverride || (websiteHostname ? `https://${websiteHostname}` : null);
  const processEndpoint = internalBase ? `${internalBase}/api/ProcessVttFile` : null;
  if (!processEndpoint) context.log('⚠️ No internal ProcessVttFile endpoint configured; set PROCESS_VTT_URL or rely on WEBSITE_HOSTNAME in Azure');

  // Validate payload shape: expect body.value array (Graph webhook)
  if (!Array.isArray(body.value)) {
    context.log('⚠️ Unexpected payload shape; expected body.value array');
    context.res = { status: 202, body: 'No notifications' };
    return;
  }

  // Process notifications
  for (const notification of body.value) {
    try {
      context.log(`🔎 Notification: ${JSON.stringify(notification)}`);

      const itemId = notification?.resourceData?.id || null;
      let itemName = inferNameFromResource(notification?.resource);

      // If name not present, try to resolve via Graph
      if (!itemName && itemId && graphClient && sharepointDriveId) {
        try {
          const item = await graphClient.api(`/drives/${sharepointDriveId}/items/${itemId}`).select('id,name').get();
          itemName = item?.name;
          context.log(`✅ Resolved item name via Graph: ${itemName}`);
        } catch (err) {
          context.log('❌ Error resolving item via Graph:', err?.message || err);
        }
      }

      if (!itemName) {
        context.log('ℹ️ Could not determine item name — skipping notification');
        continue;
      }

      // Only handle .vtt files
      if (!itemName.toLowerCase().endsWith('.vtt')) {
        context.log(`ℹ️ Skipping non-VTT file: ${itemName}`);
        continue;
      }

      // Log changeType and other details
      const changeType = notification?.changeType || 'unknown';
      context.log(`📄 ChangeType: ${changeType}`);
      context.log(`📄 SubscriptionId: ${notification?.subscriptionId}`);
      context.log(`📄 TenantId: ${notification?.tenantId}`);
      context.log(`📄 SiteUrl: ${notification?.siteUrl}`);
      context.log(`📄 UserId: ${notification?.userId}`);
      context.log(`📄 Expiration: ${notification?.expirationDateTime}`);
      context.log(`📄 ClientState: ${notification?.clientState}`);

      // Handle different change types
      switch (changeType) {
        case 'created':
          context.log(`🟢 File created: ${itemName}`);
          // TODO: Trigger downstream processing for new files
          break;
        case 'updated':
          context.log(`🟡 File updated: ${itemName}`);
          // TODO: Trigger downstream processing for updated files
          break;
        case 'deleted':
          context.log(`🔴 File deleted: ${itemName}`);
          // TODO: Handle file deletion if needed
          break;
        default:
          context.log(`⚪ Unknown changeType: ${changeType}`);
      }

      // Forward to ProcessVttFile if possible
      if (processEndpoint && fetchFn) {
        const payload = { batchMode: false, name: itemName, outputFormat: 'json' };
        context.log(`➡️ Forwarding to ProcessVttFile: ${processEndpoint} payload=${JSON.stringify(payload)}`);
        try {
          const res = await fetchFn(processEndpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          });
          const text = res && res.text ? await res.text() : '';
          context.log(`⬅️ ProcessVttFile responded ${res.status}: ${String(text).slice(0,500)}`);
        } catch (err) {
          context.log('❌ Error calling ProcessVttFile endpoint:', err?.message || err);
        }
      } else {
        context.log('ℹ️ No process endpoint or fetch available — notification logged only');
      }
    } catch (err) {
      context.log('❌ Error processing notification:', err?.message || err);
    }
  }

  // Acknowledge receipt to Graph
  context.res = { status: 202, body: 'Webhook processed' };
};

/* ----------------- Helper functions ----------------- */

function safeGetValidationToken(req) {
  try {
    if (!req) return null;
    if (req.query) {
      if (typeof req.query.get === 'function') return req.query.get('validationToken') || null;
      if (req.query.validationToken) return req.query.validationToken;
    }
    if (req.body && req.body.validationToken) return req.body.validationToken;
  } catch (e) { /* ignore */ }
  return null;
}

function inferNameFromResource(resource) {
  try {
    if (!resource) return null;
    const parts = String(resource).split('/');
    return parts[parts.length - 1] || null;
  } catch (e) { return null; }  
}