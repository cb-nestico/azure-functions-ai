const fetch = require('node-fetch');
const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { OpenAI } = require('openai');

process.on('unhandledRejection', (reason) => {
    console.error('Unhandled Rejection:', reason);
});

// Logging shim to support context.log.error/warn/info on runtimes where they are not functions
function setupLogging(context) {
    try {
        if (!context) return;

        // Ensure level methods exist
        if (typeof context.error !== 'function') context.error = (...args) => console.error(...args);
        if (typeof context.warn !== 'function')  context.warn  = (...args) => console.warn(...args);
        if (typeof context.info !== 'function')  context.info  = (...args) => console.info(...args);

        // Ensure context.log exists and attach level helpers
        if (typeof context.log !== 'function') {
            const base = (...args) => console.log(...args);
            base.error = (...args) => context.error(...args);
            base.warn  = (...args) => context.warn(...args);
            base.info  = (...args) => context.info(...args);
            context.log = base;
        } else {
            if (typeof context.log.error !== 'function') context.log.error = (...args) => context.error(...args);
            if (typeof context.log.warn  !== 'function') context.log.warn  = (...args) => context.warn(...args);
            if (typeof context.log.info  !== 'function') context.log.info  = (...args) => context.info(...args);
        }
    } catch {
        // no-op
    }
}

// ✅ HTTP Trigger Registration
app.http('ProcessVttFile', {
    methods: ['GET', 'POST'],
    authLevel: 'function',
    route: 'ProcessVttFile',
    handler: async (request, context) => {
        setupLogging(context);
        const startTime = Date.now();
        context.log('🎯 ProcessVttFile function triggered');

        try {
            let fileName, batchMode = false, fileNames = [], outputFormat = 'json';

            if (request.method === 'GET') {
                fileName = request.query.get('name');
                outputFormat = request.query.get('format') || 'json';
                context.log(`📥 GET request - fileName: ${fileName}, format: ${outputFormat}`);
            } else {
                const body = await request.text();
                context.log(`📥 POST request - body length: ${body?.length || 0}`);

                if (!body || body.trim() === '') {
                    throw new Error('Request body is empty');
                }

                try {
                    const requestData = JSON.parse(body);
                    fileName = requestData.name;
                    batchMode = requestData.batchMode || false;
                    fileNames = requestData.fileNames || [];
                    outputFormat = requestData.outputFormat || 'json';

                    context.log(`📥 Parsed request - batchMode: ${batchMode}, files: ${fileNames.length || 1}, format: ${outputFormat}`);
                } catch (parseError) {
                    throw new Error(`Invalid JSON format: ${parseError.message}`);
                }
            }

            // Batch processing with concurrency (Promise.allSettled)
            if (batchMode && fileNames.length > 1) {
                context.log(`🔄 Starting batch processing for ${fileNames.length} files`);
                return await processBatchFiles(context, fileNames, outputFormat);
            } else {
                const singleFile = fileName || (fileNames.length > 0 ? fileNames[0] : null);
                if (!singleFile) {
                    throw new Error('File name is required (provide "name" parameter or fileNames array)');
                }
                context.log(`🎥 Processing single file: ${singleFile}`);
                return await processSingleFile(context, singleFile, outputFormat);
            }

        } catch (error) {
            context.log.error('❌ Function execution failed:', error?.message || error);
            context.log.error('❌ Function error stack:', error?.stack || 'No stack trace');

            return {
                status: 500,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    success: false,
                    error: 'Function execution failed',
                    message: error?.message || String(error),
                    stack: error?.stack || 'No stack trace',
                    timestamp: new Date().toISOString(),
                    processingTimeMs: Date.now() - startTime
                })
            };
        }
    }
});

// ✅ Single File Handler
async function processSingleFile(context, fileName, outputFormat = 'json') {
    const result = await processSingleVttFile(context, fileName, outputFormat);
    const status = result && result.status ? result.status : (result?.success ? 200 : 500);

    if (outputFormat.toLowerCase() === 'html' && result.htmlContent) {
        return {
            status: status,
            headers: { 'Content-Type': 'text/html' },
            body: result.htmlContent
        };
    }

    return {
        status: status,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(result)
    };
}

// ✅ Batch Handler with concurrency and token aggregation
async function processBatchFiles(context, fileNames, outputFormat = 'json') {
    const batchStartTime = Date.now();
    const concurrencyLimit = Number(process.env.BATCH_CONCURRENCY || 3);

    context.log(`🧪 BATCH MODE: concurrencyLimit=${concurrencyLimit}, totalFiles=${fileNames.length}`);

    // Process files in parallel with Promise.allSettled
    const batches = chunkArray(fileNames, concurrencyLimit);
    const results = [];

    for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
        const batch = batches[batchIndex];
        context.log(`▶️ Processing batch ${batchIndex + 1}/${batches.length} (size=${batch.length})`);

        const batchResults = await Promise.allSettled(
            batch.map(fileName =>
                processSingleVttFile(context, fileName, outputFormat)
                    .then(fileResult => ({
                        fileName,
                        success: fileResult.success === true,
                        processingTimeMs: fileResult?.metadata?.processingTimeMs || 0,
                        ...fileResult
                    }))
                    .catch(error => ({
                        fileName,
                        success: false,
                        error: error?.message || String(error),
                        stack: error?.stack || 'No stack trace',
                        processingTimeMs: 0
                    }))
            )
        );

        results.push(...batchResults.map(r => r.value || r.reason));
        if (batchIndex < batches.length - 1) {
            context.log('⏸️ Waiting 1s before next batch...');
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
    }

    const batchTotalTime = Date.now() - batchStartTime;
    const successfulFiles = results.filter(r => r.success);
    const anySuccess = results.some(r => r.success);
    const allSuccess = results.every(r => r.success);

    // Aggregate OpenAI token usage across files
    const tokenTotals = results.reduce((acc, r) => {
        const t = r?.metadata?.openaiTokens;
        if (t) {
            acc.prompt += t.prompt || 0;
            acc.completion += t.completion || 0;
            acc.total += t.total || 0;
        }
        return acc;
    }, { prompt: 0, completion: 0, total: 0 });

    return {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            batchId: `batch_${Date.now()}`,
            success: anySuccess,
            partialSuccess: anySuccess && !allSuccess,
            batchMode: true,
            processedFiles: results.length,
            successfulFiles: successfulFiles.length,
            failedFiles: results.length - successfulFiles.length,
            results,
            metadata: {
                batchProcessingTimeMs: batchTotalTime,
                averageTimePerFile: Math.round(batchTotalTime / Math.max(1, results.length)),
                concurrencyLimit,
                totalBatches: batches.length,
                outputFormat,
                timestamp: new Date().toISOString(),
                openaiTokensTotal: tokenTotals
            }
        })
    };
}

// ...existing processSingleVttFile and helper functions remain unchanged...
// (You can keep your previous implementation for processSingleVttFile, applyOutputFormat, chunkArray, etc.)


// ✅ Granular Error Logging & Debug Statements in processSingleVttFile
async function processSingleVttFile(context, fileName, outputFormat = 'json') {
    const processingStartTime = Date.now();
    try {
        context.log(`🎬 Starting VTT processing for: ${fileName}`);

        // Configuration
        let config;
        try {
            config = {
                tenantId: process.env.TENANT_ID,
                clientId: process.env.CLIENT_ID,
                clientSecret: process.env.CLIENT_SECRET,
                openaiEndpoint: process.env.OPENAI_ENDPOINT,
                openaiKey: process.env.OPENAI_KEY,
                deployment: process.env.OPENAI_DEPLOYMENT || 'gpt-4o-text',
                sharepointDriveId: process.env.SHAREPOINT_DRIVE_ID,
                sharepointSiteUrl: process.env.SHAREPOINT_SITE_URL
            };
            const loggedConfig = { ...config, clientSecret: '***', openaiKey: '***' };
            context.log('🔧 Loaded configuration:', JSON.stringify(loggedConfig));
        } catch (configError) {
            context.log.error('❌ Error loading configuration:', configError?.message || configError);
            throw configError;
        }

        const requiredConfig = ['tenantId', 'clientId', 'clientSecret', 'openaiKey', 'openaiEndpoint', 'sharepointDriveId'];
        const missingConfig = requiredConfig.filter(key => !config[key]);
        if (missingConfig.length > 0) {
            context.log.error('❌ Missing required configuration:', missingConfig.join(', '));
            return {
                success: false,
                status: 500,
                error: `Missing required configuration: ${missingConfig.join(', ')}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }
        context.log('✅ Configuration validated');

        let openaiClient;
        try {
            openaiClient = new OpenAI({
                apiKey: config.openaiKey,
                baseURL: `${config.openaiEndpoint}/openai/deployments/${config.deployment}`,
                defaultQuery: { 'api-version': '2024-08-01-preview' },
                defaultHeaders: { 'api-key': config.openaiKey }
            });
            context.log('✅ OpenAI client initialized');
        } catch (openaiError) {
            context.log.error('❌ Error initializing OpenAI client:', openaiError?.message || openaiError);
            throw openaiError;
        }

        let graphClient;
        try {
            const credential = new ClientSecretCredential(
                config.tenantId,
                config.clientId,
                config.clientSecret
            );
            const authProvider = new TokenCredentialAuthenticationProvider(credential, {
                scopes: ['https://graph.microsoft.com/.default']
            });
            graphClient = Client.initWithMiddleware({ authProvider });
            context.log('✅ Graph client initialized');
        } catch (graphError) {
            context.log.error('❌ Error initializing Graph client:', graphError?.message || graphError);
            throw graphError;
        }

        let driveItems;
        try {
            context.log(`🔍 Searching for VTT files in drive: ${config.sharepointDriveId}`);
            driveItems = await graphClient.api(`/drives/${config.sharepointDriveId}/root/children`).get();
        } catch (driveError) {
            context.log.error('❌ Error fetching drive items:', driveError?.message || driveError);
            return {
                success: false,
                status: 500,
                error: `Error fetching drive items: ${driveError?.message || driveError}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        const vttFiles = [];
        for (const item of driveItems.value) {
            if (item.file && item.name.toLowerCase().endsWith('.vtt')) {
                vttFiles.push(item);
                context.log(`  📄 VTT: ${item.name} (${item.size} bytes)`);
            } else if (item.folder) {
                try {
                    const folderItems = await graphClient
                        .api(`/drives/${config.sharepointDriveId}/items/${item.id}/children`)
                        .get();
                    for (const subItem of folderItems.value) {
                        if (subItem.file && subItem.name.toLowerCase().endsWith('.vtt')) {
                            vttFiles.push(subItem);
                            context.log(`    📄 VTT in ${item.name}: ${subItem.name} (${subItem.size} bytes)`);
                        }
                    }
                } catch (folderError) {
                    context.log.error(`    ❌ Cannot access folder ${item.name}: ${folderError?.message || folderError}`);
                }
            }
        }
        context.log(`🎬 Total VTT files found: ${vttFiles.length}`);

        context.log(`🔎 Selecting target file for request: ${fileName}`);
        let targetFile = vttFiles.find(file => file.name.toLowerCase() === fileName.toLowerCase());
        if (!targetFile) {
            targetFile = vttFiles.find(file =>
                file.name.toLowerCase().includes(fileName.replace('.vtt', '').toLowerCase())
            );
        }
        if (!targetFile) {
            const availableFiles = vttFiles.map(f => f.name).slice(0, 10);
            context.log.error(`❌ File not found: ${fileName}`);
            return {
                success: false,
                status: 404,
                error: `File not found: ${fileName}`,
                availableFiles,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }
        context.log(`✅ Found file: ${targetFile.name} (${targetFile.size} bytes)`);

        let vttContent;
        let wasTruncated = false;
        try {
            context.log(`🔎 Fetching file details for download URL (id: ${targetFile.id})`);
            const fileDetails = await graphClient
                .api(`/drives/${config.sharepointDriveId}/items/${targetFile.id}`)
                .select('@microsoft.graph.downloadUrl,name,size,id')
                .get();

            const downloadUrl = fileDetails['@microsoft.graph.downloadUrl'] || targetFile['@microsoft.graph.downloadUrl'];
            context.log(`⬇️ Download URL present: ${Boolean(downloadUrl)} for ${fileDetails.name}`);
            if (!downloadUrl) {
                return {
                    success: false,
                    status: 502,
                    error: 'Download URL not available for the selected file.',
                    fileId: targetFile.id,
                    fileName: targetFile.name,
                    processedAt: new Date().toISOString(),
                    processingTimeMs: Date.now() - processingStartTime
                };
            }

            const httpFetch = globalThis.fetch ? globalThis.fetch.bind(globalThis) : fetch;
            const response = await httpFetch(downloadUrl);
            context.log(`⬇️ HTTP GET ${response.status} for VTT content (${targetFile.name})`);
            if (!response.ok) {
                throw new Error(`Failed to download VTT: HTTP ${response.status}`);
            }

            const raw = await response.text();
            const MAX_CHARS = Number(process.env.MAX_VTT_CHARS || 32000);
            wasTruncated = raw.length > MAX_CHARS;
            vttContent = wasTruncated ? raw.slice(0, MAX_CHARS) : raw;
            context.log(`✅ Downloaded content: ${raw.length} characters`);
            if (wasTruncated) {
                context.log(`⚠️ Content truncated from ${raw.length} to ${MAX_CHARS} characters`);
            }
        } catch (downloadError) {
            context.log.error('❌ Error downloading VTT file:', downloadError);
            context.log.error('❌ Download error stack:', downloadError?.stack || 'No stack trace');
            return {
                success: false,
                status: 500,
                error: `Error downloading VTT file: ${downloadError?.message || downloadError}`,
                stack: downloadError?.stack || 'No stack trace',
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        let timestampBlocks;
        try {
            timestampBlocks = parseVttTimestamps(vttContent);
            context.log(`✅ Parsed VTT timestamps, blocks: ${timestampBlocks.length}`);
        } catch (parseError) {
            context.log.error('❌ Error parsing VTT timestamps:', parseError?.message || parseError);
            context.log.error('❌ Parse error stack:', parseError?.stack || 'No stack trace');
            return {
                success: false,
                status: 500,
                error: `Error parsing VTT timestamps: ${parseError?.message || parseError}`,
                stack: parseError?.stack || 'No stack trace',
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        let meetingMetadata;
        try {
            meetingMetadata = extractMeetingMetadata(vttContent, targetFile, config.sharepointSiteUrl);
            context.log(`✅ Extracted meeting metadata: ${JSON.stringify(meetingMetadata)}`);
        } catch (metaError) {
            context.log.error('❌ Error extracting meeting metadata:', metaError?.message || metaError);
            context.log.error('❌ Metadata error stack:', metaError?.stack || 'No stack trace');
            return {
                success: false,
                status: 500,
                error: `Error extracting meeting metadata: ${metaError?.message || metaError}`,
                stack: metaError?.stack || 'No stack trace',
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        const transcriptText = timestampBlocks.map(b => `${b.timestamp || ""} ${b.content || ""}`).join("\n");

        // --- Refined Key Point Extraction Logic ---
        const aiPrompt = `
You are a service that outputs ONLY strict JSON. No prose. No Markdown. No code fences.
Analyze the transcript and return exactly this JSON schema:

{
  "summary": "2-3 sentences executive summary",
  "keyPoints": [
    { "title": "short topic or action", "timestamp": "HH:MM:SS", "speaker": "name if known", "videoLink": "" }
  ]
}

Rules:
- Output a single JSON object only.
- keyPoints: 5–12 items when possible.
- If a field is unknown, use an empty string.
Transcript:
${transcriptText}
`;

        let aiParsed = {};
        let summary = "";
        let keyPoints = [];
        // Token usage log holder
        let tokensLog = { prompt: 0, completion: 0, total: 0 };
        try {
            const aiResponse = await openaiClient.chat.completions.create({
                model: config.deployment,
                messages: [
                    { role: 'system', content: 'You output only strict JSON objects that match the user schema.' },
                    { role: 'user', content: aiPrompt }
                ],
                temperature: 0.2,
                max_tokens: 1024,
                // Force structured JSON from Azure OpenAI (2024-08-01-preview)
                response_format: { type: 'json_object' }
            });
            context.log('🧠 Raw AI response:', aiResponse);

            // Token usage logging
            const usage = aiResponse?.usage || {};
            tokensLog = {
                prompt: usage.prompt_tokens || 0,
                completion: usage.completion_tokens || 0,
                total: usage.total_tokens || 0
            };
            context.log(`🧾 OpenAI tokens: ${JSON.stringify(tokensLog)}`);

            const aiContent = aiResponse?.choices?.[0]?.message?.content ?? '';
            aiParsed = safeParseModelJson(aiContent);
            context.log('🧠 Parsed AI response:', aiParsed);

            summary = (aiParsed && typeof aiParsed.summary === 'string') ? aiParsed.summary : "";
            keyPoints = Array.isArray(aiParsed?.keyPoints) ? aiParsed.keyPoints.filter(Boolean) : [];
        } catch (err) {
            context.log.error('❌ Error calling or parsing OpenAI:', err);
            context.log.error('❌ OpenAI error stack:', err?.stack || 'No stack trace');
            summary = "";
            keyPoints = [];
        }

        // Build video links if available
        if (keyPoints.length > 0) {
            keyPoints = keyPoints.map(point => ({
                ...point,
                videoLink: point?.timestamp && meetingMetadata.videoUrl
                    ? `${meetingMetadata.videoUrl}#t=${(point.timestamp || '').replace(/:/g, 'h').replace(/h(\d{2})$/, 'm$1s')}`
                    : (point?.videoLink || "")
            }));
        }

        // Fallbacks: never fail the request just because AI format varied
        if (!summary || summary.trim().length < 20) {
            summary = generateFallbackSummary(transcriptText);
        }
        if (!Array.isArray(keyPoints) || keyPoints.length < 3) {
            const fallback = deriveKeyPointsFallbackFromText(transcriptText);
            keyPoints = fallback.slice(0, 8).map((title, idx) => ({
                title,
                timestamp: timestampBlocks[idx]?.timestamp || "",
                speaker: timestampBlocks[idx]?.speaker || "",
                videoLink: timestampBlocks[idx]?.timestamp && meetingMetadata.videoUrl
                    ? `${meetingMetadata.videoUrl}#t=${(timestampBlocks[idx].timestamp || '').replace(/:/g, 'h').replace(/h(\d{2})$/, 'm$1s')}`
                    : ""
            }));
        }

        const metadata = {
            endpoint: config.openaiEndpoint,
            deployment: config.deployment,
            fileSize: targetFile.size,
            originalContentLength: vttContent.length,
            processedContentLength: vttContent.length,
            truncated: wasTruncated,
            estimatedTokens: Math.round(vttContent.length / 4),
            totalTimestamps: timestampBlocks.length,
            totalKeyPoints: keyPoints.length,
            processedAt: new Date().toISOString(),
            processingTimeMs: Date.now() - processingStartTime,
            // Per-file OpenAI token usage
            openaiTokens: tokensLog
        };

        let result;
        try {
            result = {
                success: true,
                meetingTitle: meetingMetadata.title,
                date: meetingMetadata.date,
                videoUrl: meetingMetadata.videoUrl,
                file: fileName,
                actualFile: targetFile.name,
                summary,
                keyPoints,
                timestampBlocks,
                metadata
            };
            result = await applyOutputFormat(context, result, outputFormat);
            context.log('✅ Output formatted');
        } catch (formatError) {
            context.log.error('❌ Error formatting output:', formatError?.message || formatError);
            context.log.error('❌ Format error stack:', formatError?.stack || 'No stack trace');
            return {
                success: false,
                status: 500,
                error: `Error formatting output: ${formatError?.message || formatError}`,
                stack: formatError?.stack || 'No stack trace',
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        return result;

    } catch (error) {
        context.log.error(`❌ Error in processSingleVttFile for ${fileName}:`, error?.message || error);
        context.log.error('❌ Single file error stack:', error?.stack || 'No stack trace');
        return {
            success: false,
            status: 500,
            error: error?.message || String(error),
            stack: error?.stack || 'No stack trace',
            file: fileName,
            processedAt: new Date().toISOString(),
            processingTimeMs: Date.now() - processingStartTime
        };
    }
}

// ✅ Helper and formatting functions
async function applyOutputFormat(context, result, outputFormat) {
    switch (outputFormat.toLowerCase()) {
        case 'html':
            return generateHtmlOutput(context, result);
        case 'markdown':
            return generateMarkdownOutput(context, result);
        case 'summary':
            return generateSummaryOutput(context, result);
        case 'json':
        default:
            return result;
    }
}

function generateHtmlOutput(context, result) {
    const { meetingTitle, keyPoints, summary, metadata } = result;

    const html = `<!DOCTYPE html>
<html>
<head>
    <title>Meeting Analysis: ${meetingTitle}</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; color: #333; }
        .header { text-align: center; border-bottom: 2px solid #007acc; padding-bottom: 20px; margin-bottom: 30px; }
        .summary { background: #f5f5f5; padding: 20px; border-radius: 8px; margin: 20px 0; }
        .key-points { margin: 30px 0; }
        ul.keypoint-list { margin: 0 0 0 20px; padding: 0; }
        li.keypoint-item { margin-bottom: 14px; padding: 10px 0; border-bottom: 1px solid #eee; }
        .timestamp { font-weight: bold; color: #007acc; font-family: monospace; margin-right: 10px; }
        .speaker { font-style: italic; color: #666; margin-right: 10px; }
        .title { font-weight: bold; }
        .metadata { background: #e8f4f8; padding: 20px; border-radius: 8px; margin-top: 30px; }
        h1 { color: #007acc; margin: 0; }
        h2 { color: #005a9e; border-bottom: 1px solid #ddd; padding-bottom: 10px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Meeting Analysis: ${meetingTitle}</h1>
        <p><strong>Generated:</strong> ${new Date().toLocaleDateString()} at ${new Date().toLocaleTimeString()}</p>
        <p><strong>Processing Time:</strong> ${metadata.processingTimeMs}ms | <strong>Key Points:</strong> ${keyPoints.length}</p>
    </div>
    <div class="summary">
        <h2>📋 Executive Summary</h2>
        <p>${summary.replace(/\n/g, '<br>')}</p>
    </div>
    <div class="key-points">
        <h2>🎯 Key Discussion Points (${keyPoints.length} items)</h2>
        <ul class="keypoint-list">
            ${keyPoints
                .filter(point => point.title && point.title.trim() !== '')
                .map(point => `
                <li class="keypoint-item">
                    ${point.timestamp ? `<span class="timestamp">${point.timestamp}</span>` : ''}
                    ${point.speaker ? `<span class="speaker">${point.speaker}</span>` : ''}
                    <span class="title">${point.title}</span>
                    ${point.videoLink ? `<a class="video-link" href="${point.videoLink}" target="_blank">🔗 Video</a>` : ''}
                </li>
            `).join('')}
        </ul>
    </div>
    <div class="metadata">
        <h3>📊 Processing Information</h3>
        <p><strong>File:</strong> ${metadata.fileSize} bytes | <strong>Timestamps:</strong> ${metadata.totalTimestamps} | <strong>Processing:</strong> ${metadata.processingTimeMs}ms</p>
        <p><strong>Tokens:</strong> prompt ${metadata.openaiTokens?.prompt || 0}, completion ${metadata.openaiTokens?.completion || 0}, total ${metadata.openaiTokens?.total || 0}</p>
    </div>
    <footer style="text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; color: #666;">
        <p><em>Generated by Azure Functions VTT Meeting Transcript Processor</em></p>
        <p>File: ${result.actualFile} | Processed: ${metadata.processedAt}</p>
    </footer>
</body>
</html>`;

    return {
        success: true,
        outputFormat: 'html',
        htmlContent: html,
        downloadable: {
            contentType: 'text/html',
            fileName: `${meetingTitle.replace(/[^a-z0-9]/gi, '_')}_Analysis.html`,
            content: html,
            size: html.length
        }
    };
}

function generateMarkdownOutput(context, result) {
    const { meetingTitle, keyPoints, summary, metadata } = result;

    const markdown = `# Meeting Analysis: ${meetingTitle}

**Generated:** ${new Date().toLocaleDateString()} at ${new Date().toLocaleTimeString()}  
**Processing Time:** ${metadata.processingTimeMs}ms | **Key Points:** ${keyPoints.length}

## 📋 Executive Summary

${summary}

## 🎯 Key Discussion Points (${keyPoints.length} items)

${keyPoints.map((point, index) => `### ${index + 1}. ${point.timestamp} - ${point.title}

**Speaker:** ${point.speaker}

---`).join('\n\n')}

## 📊 Processing Information

- **File Size:** ${Math.round(metadata.fileSize / 1024)}KB
- **Timestamps:** ${metadata.totalTimestamps}
- **Processing Time:** ${metadata.processingTimeMs}ms
- **Tokens:** prompt ${metadata.openaiTokens?.prompt || 0}, completion ${metadata.openaiTokens?.completion || 0}, total ${metadata.openaiTokens?.total || 0}

---

**File:** ${result.actualFile} | **Processed:** ${metadata.processedAt}  
*Generated by Azure Functions VTT Meeting Transcript Processor*`;

    return {
        ...result,
        outputFormat: 'markdown',
        markdownContent: markdown,
        downloadable: {
            contentType: 'text/markdown',
            fileName: `${meetingTitle.replace(/[^a-z0-9]/gi, '_')}_Analysis.md`,
            content: markdown,
            size: markdown.length
        }
    };
}

function generateSummaryOutput(context, result) {
    const { meetingTitle, keyPoints, summary, metadata } = result;

    return {
        success: true,
        meetingTitle,
        summary,
        keyPointsCount: keyPoints.length,
        topKeyPoints: keyPoints.slice(0, 5).map(point => ({
            timestamp: point.timestamp,
            title: point.title,
            speaker: point.speaker
        })),
        processingTimeMs: metadata.processingTimeMs,
        fileSize: metadata.fileSize,
        tokens: metadata.openaiTokens,
        outputFormat: 'summary',
        processedAt: metadata.processedAt
    };
}

function deriveKeyPointsFallbackFromText(text) {
    if (!text) return [];
    const bullets = Array.from(new Set(
        text.split(/\r?\n+/)
            .filter(l => /^\s*[-•–]/.test(l))
            .map(l => l.replace(/^\s*[-•–]\s*/, "").trim())
            .filter(Boolean)
    ));
    if (bullets.length >= 3) return bullets.slice(0, 12);
    return text
        .split(/(?<=[.!?])\s+/)
        .map(s => s.trim())
        .filter(s => /^[A-Z][a-z]+/.test(s) || /action|decision|important|key/i.test(s))
        .slice(0, 8);
}

function parseVttTimestamps(vttContent) {
    if (!vttContent) return [];

    const contentBlocks = [];
    const lines = vttContent.split('\n');
    let currentBlock = null;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        const timestampMatch = line.match(/(\d{2}:\d{2}:\d{2})\.\d{3}/);
        if (timestampMatch) {
            if (currentBlock) contentBlocks.push(currentBlock);
            currentBlock = { timestamp: timestampMatch[1], content: '', speaker: null };
        } else if (currentBlock && line.length > 0) {
            const speakerMatch = line.match(/<v\s+([^>]+)>(.+)<\/v>/);
            if (speakerMatch) {
                currentBlock.speaker = speakerMatch[1];
                currentBlock.content += speakerMatch[2] + ' ';
            } else {
                currentBlock.content += line + ' ';
            }
        }
    }

    if (currentBlock) contentBlocks.push(currentBlock);
    return contentBlocks;
}

function extractMeetingMetadata(vttContent, fileMetadata, sharepointSiteUrl) {
    const noteMatch = vttContent.match(/NOTE\s+(.+)/);
    const meetingTitle = noteMatch ? noteMatch[1].trim()
        : fileMetadata.name.replace('.vtt', '').replace(/[-_]/g, ' ').replace(/\b\w/g, l => l.toUpperCase());

    const videoUrl = sharepointSiteUrl
        ? `${sharepointSiteUrl}/Shared%20Documents/${fileMetadata.name.replace('.vtt', '')}`
        : "https://yourtenant.sharepoint.com/video-placeholder";

    return {
        title: meetingTitle,
        videoUrl,
        date: new Date().toISOString().split('T')[0],
        filename: fileMetadata.name
    };
}

function chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
        chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
}

// Safe JSON parse for model outputs that may include code fences
function safeParseModelJson(text) {
    if (!text) return {};
    let cleaned = String(text).trim();
    cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/```$/i, '').trim();
    try { return JSON.parse(cleaned); } catch {}
    const start = cleaned.indexOf('{');
    const end = cleaned.lastIndexOf('}');
    if (start >= 0 && end > start) {
        const candidate = cleaned.slice(start, end + 1);
        try { return JSON.parse(candidate); } catch {}
    }
    return {};
}

function generateFallbackSummary(text) {
    if (!text) return "Meeting transcript processed. Key topics extracted.";
    const clean = text
        .replace(/^\s*\d{2}:\d{2}:\d{2}\s*/gm, '')
        .replace(/<[^>]+>/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    const sentences = clean.split(/(?<=[.!?])\s+/).filter(s => s.length > 0);
    const picked = sentences.slice(0, 3).join(' ');
    return picked || "Meeting transcript processed. Key topics extracted.";
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        parseVttTimestamps,
        extractMeetingMetadata,
        deriveKeyPointsFallbackFromText
    };
}