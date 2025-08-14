const { app } = require('@azure/functions');

// Register SharePointWebhook using the new programming model
try {
  const sharePointHandlerClassic = require('./src/functions/SharePointWebhook/index.js');

  app.http('SharePointWebhook', {
    methods: ['POST', 'GET', 'OPTIONS'],
    authLevel: 'function',
    handler: async (request, context) => {
      let body = null;
      try { body = await request.json(); } catch (_) { /* ignore */ }

      const classicReq = {
        query: {
          get: (key) => (request.query && typeof request.query.get === 'function') ? request.query.get(key) : (request.query && request.query[key]),
          validationToken: (request.query && (typeof request.query.get === 'function' ? request.query.get('validationToken') : request.query.validationToken))
        },
        body
      };

      await sharePointHandlerClassic(context, classicReq);

      if (context && context.res) {
        return { status: context.res.status || 200, body: context.res.body, headers: context.res.headers };
      }
      return { status: 202, body: 'Webhook processed' };
    }
  });
} catch (err) {
  console.error('Could not register SharePointWebhook wrapper:', err && err.stack ? err.stack : err);
}

// Register ProcessVttFile using the new programming model
try {
  const processVttFileHandlerClassic = require('./src/functions/ProcessVttFile/index.js');

  app.http('ProcessVttFile', {
    methods: ['POST', 'GET', 'OPTIONS'],
    authLevel: 'function',
    handler: async (request, context) => {
      let body = null;
      try { body = await request.json(); } catch (_) { /* ignore */ }

      const classicReq = {
        query: {
          get: (key) => (request.query && typeof request.query.get === 'function') ? request.query.get(key) : (request.query && request.query[key])
        },
        body
      };

      await processVttFileHandlerClassic(context, classicReq);

      if (context && context.res) {
        return { status: context.res.status || 200, body: context.res.body, headers: context.res.headers };
      }
      return { status: 202, body: 'VTT file processed' };
    }
  });
} catch (err) {
  console.error('Could not register ProcessVttFile wrapper:', err && err.stack ? err.stack : err);
}

// ...existing code...