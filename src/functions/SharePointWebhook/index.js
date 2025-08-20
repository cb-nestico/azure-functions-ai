const axios = require('axios');

// Helper: Send notification to Teams and Email (SendGrid)
async function sendNotification(type, fileName, changeType, details) {
    // ...existing notification logic...
}

// Helper: Validate incoming request
function validateRequest(req) {
    // ...existing validation logic...
    return true;
}

// Export only the handler function
module.exports = async (req, context) => {
    // Handle Graph webhook validation
    if (req.query && req.query.validationToken) {
        context.log('Validation token received');
        return { status: 200, body: req.query.validationToken };
    }

    // Parse incoming notification
    let notifications = [];
    try {
        notifications = req.body.value || [];
    } catch (err) {
        context.log.error('Failed to parse notification payload', err);
        return { status: 400, body: 'Invalid payload' };
    }

    for (const notification of notifications) {
        const {
            changeType,
            resource,
            subscriptionId,
            tenantId,
            siteUrl,
            userId,
            expirationDateTime,
            clientState
        } = notification;

        // Validate request
        if (!validateRequest(req)) {
            context.log.warn('Request validation failed', { subscriptionId, tenantId });
            continue;
        }

        // Extract file name (assume resource ends with file name)
        const fileName = resource.split('/').pop();

        // Only process .vtt files
        if (!fileName.endsWith('.vtt')) {
            context.log(`Skipping non-VTT file: ${fileName}`);
            continue;
        }

        // Log notification details
        context.log(`[Webhook] ${changeType} - ${fileName}`, {
            subscriptionId, tenantId, siteUrl, userId, expirationDateTime, clientState
        });

        // Downstream processing
        try {
            if (changeType === 'created' || changeType === 'updated') {
                // Call ProcessVttFile endpoint
                await axios.post(process.env.PROCESS_VTT_ENDPOINT, {
                    name: fileName,
                    siteUrl,
                    subscriptionId,
                    tenantId,
                    userId
                });
                await sendNotification('FileProcessed', fileName, changeType, { siteUrl, userId });
            } else if (changeType === 'deleted') {
                // TODO: Handle deleted files (e.g., cleanup, notify)
                await sendNotification('FileDeleted', fileName, changeType, { siteUrl, userId });
            }
        } catch (err) {
            context.log.error(`Error processing ${changeType} for ${fileName}`, err);
            await sendNotification('ProcessingError', fileName, changeType, { error: err.message });
        }
    }

    // Acknowledge receipt
    return { status: 202, body: 'Notifications processed' };
};