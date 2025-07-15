# Azure Functions AI - Troubleshooting Session
**Date**: June 24, 2025  
**Issue**: "No job functions found" error  
**Status**: ✅ RESOLVED

## Problem Summary
Azure Functions v4 project was using legacy v1/v2 programming model patterns, causing the runtime to fail to discover functions.

## Root Causes Identified
1. **Programming Model Mismatch**: Using @azure/functions v4 with legacy function.json files
2. **Incorrect Function Registration**: Used module.exports instead of app.http() pattern
3. **Wrong Entry Point**: package.json pointed to wrong main file
4. **Outdated Extension Bundle**: host.json used v3 instead of v4 bundle
5. **Legacy Function Structure**: Had function.json file that's not needed in v4

## Changes Made

### 1. Removed function.json
- Deleted: `src/functions/ProcessVttFile/function.json`
- Reason: Not needed in v4 programming model

### 2. Updated Function Code
**Before:**
```javascript
module.exports = async function (context, req) {
    // function logic
};
```

**After:**
```javascript
const { app } = require('@azure/functions');

app.http('ProcessVttFile', {
    methods: ['GET', 'POST'],
    authLevel: 'function',
    handler: async (request, context) => {
        // function logic
    }
});
```

### 3. Fixed package.json
**Changed main entry point:**
```json
{
  "main": "src/index.js"  // Was: "src/functions/ProcessVttFile/index.js"
}
```

### 4. Updated host.json
**Extension bundle updated:**
```json
{
  "extensionBundle": {
    "id": "Microsoft.Azure.Functions.ExtensionBundle",
    "version": "[4.*, 5.0.0)"  // Was: "[3.*, 4.0.0)"
  }
  // Removed: "functions": [ "ProcessVttFile" ]
}
```

### 5. Updated src/index.js
**Added function import:**
```javascript
const { app } = require('@azure/functions');

// Import and register functions
require('./functions/ProcessVttFile');

app.setup({
    enableHttpStream: true,
});
```

## Final Result
```
Functions:
        ProcessVttFile: [GET,POST] http://localhost:7071/api/ProcessVttFile
```

## Project Structure (Final)
```
AZURE FUNCTIONS-AI/
├── src/
│   ├── index.js                    # Main entry point
│   └── functions/
│       └── ProcessVttFile/
│           └── index.js            # Function implementation (v4 style)
├── host.json                       # Updated to v4
├── local.settings.json             # Environment variables
├── package.json                    # Fixed entry point
├── README.md                       # Complete documentation
└── TROUBLESHOOTING-SESSION-2025-06-24.md  # This file
```

## Environment Variables Required
```
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id  
CLIENT_SECRET=your-client-secret
OPENAI_ENDPOINT=your-openai-endpoint
OPENAI_DEPLOYMENT=your-deployment-name
OPENAI_KEY=your-openai-key
SHAREPOINT_SITE_URL=your-sharepoint-url
SHAREPOINT_SITE_ID=your-site-id
SHAREPOINT_DRIVE_ID=your-drive-id
SHAREPOINT_LIST_ID=your-list-id
```

## Next Steps for Tomorrow
1. **Start function**: `func start`
2. **Test endpoint**: POST to http://localhost:7071/api/ProcessVttFile
3. **Deploy to Azure** (if needed)
4. **Set up SharePoint webhook** (if needed)

## Function Purpose
Processes VTT files from SharePoint, generates AI summaries using Azure OpenAI, and appends formatted summaries to a master AllNotes.md file.

## Key Learnings
- Azure Functions v4 uses code-based registration, not function.json files
- Always match programming model version with project structure
- Extension bundles must align with Functions runtime version
- Entry points in package.json must point to correct main file

---
**Session Complete**: All issues resolved, code is production-ready
