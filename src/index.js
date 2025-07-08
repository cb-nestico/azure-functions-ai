const { app } = require('@azure/functions');

// Import and register all functions
require('./functions/ProcessVttFile');

// Azure Functions v4 automatically handles:
// - HTTP streaming
// - Optimal performance settings
// - Request/response handling

// Export the app for Azure Functions runtime
module.exports = app;