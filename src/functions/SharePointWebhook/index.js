const axios = require('axios');

/**
 * Enhanced SharePointWebhook Handler
 * - Structured logging
 * - Event validation and parsing
 * - Aggregated output
 */
module.exports = async function (req, context) {
    function logEvent(level, message, meta = {}) {
        if (context.log && typeof context.log[level] === 'function') {
            context.log[level](`[SharePointWebhook] ${message}`, meta);
        } else {
            context.log(`[${level.toUpperCase()}] [SharePointWebhook] ${message}`, meta);
        }
    }

    // Parse body if it's a stream
    let body = req.body;
    if (body && typeof body.getReader === 'function') {
        // Node.js stream: read and parse
        const chunks = [];
        const reader = body.getReader();
        let done, value;
        while (true) {
            ({ done, value } = await reader.read());
            if (done) break;
            chunks.push(value);
        }
        const raw = Buffer.concat(chunks).toString();
        try {
            body = JSON.parse(raw);
        } catch (err) {
            logEvent('error', 'Failed to parse streamed body', { error: err.message, raw });
            return {
                status: 400,
                body: { success: false, error: 'Invalid JSON in streamed body' }
            };
        }
    }

    // 1. Webhook validation
    if (req.query && req.query.validationToken) {
        logEvent('info', 'Validation token received');
        return {
            status: 200,
            body: req.query.validationToken
        };
    }

    // 2. Parse and validate notifications
    const notifications = body?.value;
    // --- Recommended code: Only log warning for external webhook requests ---
    if (!Array.isArray(notifications) || notifications.length === 0) {
        if (body?.value !== undefined) {
            logEvent('warn', 'No notifications received', { body });
        }
        return {
            status: 400,
            body: { success: false, error: 'No notifications received' }
        };
    }
    // -----------------------------------------------------------------------

    // 3. Process each notification
    const results = [];
    for (const notification of notifications) {
        const { changeType, resource, subscriptionId, clientState } = notification;
        const eventMeta = { changeType, resource, subscriptionId, clientState, timestamp: new Date().toISOString() };

        try {
            logEvent('info', 'Processing notification', eventMeta);

            // Only process .vtt files
            if (!resource || !resource.toLowerCase().endsWith('.vtt')) {
                logEvent('info', 'Skipping non-VTT resource', eventMeta);
                results.push({ resource, processed: false, reason: 'Not a VTT file' });
                continue;
            }

            // Handle created/updated events
            if (changeType === 'created' || changeType === 'updated') {
                // Call ProcessVttFile endpoint (no changes to its logic)
                try {
                    const response = await axios.post(
                        process.env.PROCESS_VTT_ENDPOINT || 'http://localhost:7071/api/ProcessVttFile',
                        { name: resource }
                    );
                    logEvent('info', `Triggered VTT processing for ${resource}`, { status: response.status });
                    results.push({
                        resource,
                        processed: true,
                        event: changeType,
                        status: 'Triggered',
                        processVttStatus: response.status,
                        processVttResult: response.data
                    });
                } catch (err) {
                    logEvent('error', `Error calling ProcessVttFile for ${resource}`, { error: err.message });
                    results.push({
                        resource,
                        processed: false,
                        event: changeType,
                        status: 'Error triggering VTT processing',
                        error: err.message
                    });
                }
            }
            // Handle deleted events
            else if (changeType === 'deleted') {
                logEvent('info', `Resource deleted: ${resource}`, eventMeta);
                results.push({ resource, processed: true, event: changeType, status: 'Deleted' });
            }
            // Unknown event
            else {
                logEvent('warn', 'Unknown changeType', eventMeta);
                results.push({ resource, processed: false, event: changeType, status: 'Unknown event type' });
            }
        } catch (err) {
            logEvent('error', 'Error processing notification', { ...eventMeta, error: err.message });
            results.push({ resource, processed: false, error: err.message });
        }
    }

    // 4. Aggregate and return results
    return {
        status: 200,
        body: {
            success: true,
            processedCount: results.filter(r => r.processed).length,
            skippedCount: results.filter(r => !r.processed).length,
            results,
            timestamp: new Date().toISOString()
        }
    };
};