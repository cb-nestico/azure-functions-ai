const fetch = require('node-fetch');
const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { OpenAI } = require('openai');

process.on('unhandledRejection', (reason) => {
    console.error('Unhandled Rejection:', reason);
});

// ✅ HTTP Trigger Registration
app.http('ProcessVttFile', {
    methods: ['GET', 'POST'],
    authLevel: 'function',
    route: 'ProcessVttFile',
    handler: async (request, context) => {
        const startTime = Date.now();
        context.log('🎯 ProcessVttFile function triggered');

        try {
            // Parse request - handle both batch and single file processing
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

            // Determine processing mode
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
            context.log.error('Stack trace:', error?.stack);

            return {
                status: 500,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    success: false,
                    error: 'Function execution failed',
                    message: error?.message || String(error),
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
    const status = result && result.status ? result.status : 200;
    return {
        status: status,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(result)
    };
}

// ✅ Batch Handler with per-file success/failure (sequential, resilient)
async function processBatchFiles(context, fileNames, outputFormat = 'json') {
    const results = [];
    const batchStartTime = Date.now();

    // Sequential during troubleshooting
    const concurrencyLimit = 1;
    const batches = chunkArray(fileNames, concurrencyLimit);

    context.log(`🧪 BATCH MODE: concurrencyLimit=${concurrencyLimit}, totalFiles=${fileNames.length}, totalBatches=${batches.length}`);

    try {
        for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
            const batch = batches[batchIndex];
            context.log(`▶️ Processing batch ${batchIndex + 1}/${batches.length} (size=${batch.length})`);

            for (const fileName of batch) {
                const fileStartTime = Date.now();
                context.log(`  • Processing file in batch: ${fileName}`);
                try {
                    const fileResult = await processSingleVttFile(context, fileName, outputFormat);
                    results.push({
                        fileName,
                        success: fileResult.success === true,
                        processingTimeMs: Date.now() - fileStartTime,
                        ...fileResult
                    });
                } catch (error) {
                    context.log.error(`  ❌ Unhandled error for ${fileName}:`, error);
                    results.push({
                        fileName,
                        success: false,
                        error: error?.message || String(error),
                        processingTimeMs: Date.now() - fileStartTime
                    });
                }
            }

            if (batchIndex < batches.length - 1) {
                context.log('⏸️ Waiting 1s before next batch...');
                await new Promise(resolve => setTimeout(resolve, 1000));
            }
        }
    } catch (err) {
        context.log.error('❌ Unhandled error in processBatchFiles:', err);
        // Do not fail the whole request; return what we have
    }

    const batchTotalTime = Date.now() - batchStartTime;
    const successfulFiles = results.filter(r => r.success);

    return {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            success: successfulFiles.length === results.length,
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
                timestamp: new Date().toISOString()
            }
        })
    };
}

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

        // Validate configuration
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

        // Initialize OpenAI client
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

        // Initialize Microsoft Graph client
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

        // Find VTT files in SharePoint
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

        // Find target file
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

        // Download file content (refresh download URL and log HTTP status)
        let vttContent;
        try {
            context.log(`🔎 Fetching file details for download URL (id: ${targetFile.id})`);
            const fileDetails = await graphClient
                // Note: do NOT encode the id here; Graph expects the raw driveItem id
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
            const MAX_CHARS = 32000;
            const wasTruncated = raw.length > MAX_CHARS;
            vttContent = wasTruncated ? raw.slice(0, MAX_CHARS) : raw;
            context.log(`✅ Downloaded content: ${raw.length} characters`);
            if (wasTruncated) {
                context.log(`⚠️ Content truncated from ${raw.length} to ${MAX_CHARS} characters`);
            }
        } catch (downloadError) {
            context.log.error('❌ Error downloading VTT file:', downloadError);
            return {
                success: false,
                status: 500,
                error: `Error downloading VTT file: ${downloadError?.message || downloadError}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        // Parse VTT timestamps
        let timestampBlocks;
        try {
            timestampBlocks = parseVttTimestamps(vttContent);
            context.log(`✅ Parsed VTT timestamps, blocks: ${timestampBlocks.length}`);
        } catch (parseError) {
            context.log.error('❌ Error parsing VTT timestamps:', parseError?.message || parseError);
            return {
                success: false,
                status: 500,
                error: `Error parsing VTT timestamps: ${parseError?.message || parseError}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        // Extract meeting metadata
        let meetingMetadata;
        try {
            meetingMetadata = extractMeetingMetadata(vttContent, targetFile, config.sharepointSiteUrl);
            context.log(`✅ Extracted meeting metadata: ${JSON.stringify(meetingMetadata)}`);
        } catch (metaError) {
            context.log.error('❌ Error extracting meeting metadata:', metaError?.message || metaError);
            return {
                success: false,
                status: 500,
                error: `Error extracting meeting metadata: ${metaError?.message || metaError}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        // Generate summary using OpenAI
        let summary;
        try {
            context.log('🧠 Generating summary with OpenAI...');
            const aiResponse = await openaiClient.chat.completions.create({
                model: config.deployment,
                messages: [
                    { role: "system", content: "You are an expert meeting transcript summarizer." },
                    { role: "user", content: vttContent }
                ],
                max_tokens: 2048,
                temperature: 0.3
            });
            summary = aiResponse.choices[0].message.content;
            context.log('✅ OpenAI summary generated');
        } catch (aiError) {
            context.log.error('❌ Error generating summary with OpenAI:', aiError?.message || aiError);
            return {
                success: false,
                status: 500,
                error: `Error generating summary with OpenAI: ${aiError?.message || aiError}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        // Extract key points
        let keyPoints;
        try {
            keyPoints = extractKeyPoints(summary, timestampBlocks, meetingMetadata.videoUrl);
            context.log(`✅ Extracted key points: ${keyPoints.length}`);
        } catch (kpError) {
            context.log.error('❌ Error extracting key points:', kpError?.message || kpError);
            return {
                success: false,
                status: 500,
                error: `Error extracting key points: ${kpError?.message || kpError}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        // Prepare metadata
        const metadata = {
            endpoint: config.openaiEndpoint,
            deployment: config.deployment,
            fileSize: targetFile.size,
            originalContentLength: vttContent.length,
            processedContentLength: vttContent.length,
            truncated: false,
            estimatedTokens: Math.round(vttContent.length / 4),
            totalTimestamps: timestampBlocks.length,
            totalKeyPoints: keyPoints.length,
            processedAt: new Date().toISOString(),
            processingTimeMs: Date.now() - processingStartTime
        };

        // Format output
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
            return {
                success: false,
                status: 500,
                error: `Error formatting output: ${formatError?.message || formatError}`,
                processedAt: new Date().toISOString(),
                processingTimeMs: Date.now() - processingStartTime
            };
        }

        return result;

    } catch (error) {
        context.log.error(`❌ Error in processSingleVttFile for ${fileName}:`, error?.message || error);
        return {
            success: false,
            status: 500,
            error: error?.message || String(error),
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
        .key-point { margin: 15px 0; padding: 15px; border-left: 4px solid #007acc; background: #fafafa; border-radius: 4px; }
        .timestamp { font-weight: bold; color: #007acc; font-family: monospace; }
        .speaker { font-style: italic; color: #666; margin-left: 10px; }
        .title { font-weight: bold; margin: 5px 0; }
        .video-link { color: #007acc; text-decoration: none; font-size: 14px; }
        .video-link:hover { text-decoration: underline; }
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
        ${keyPoints.map((point) => `
            <div class="key-point">
                <div class="timestamp">${point.timestamp}</div>
                <span class="speaker">${point.speaker}</span>
                <div class="title">${point.title}</div>
                <a href="${point.videoLink}" class="video-link" target="_blank">🎥 Watch this moment in video</a>
            </div>
        `).join('')}
    </div>
    <div class="metadata">
        <h3>📊 Processing Information</h3>
        <p><strong>File:</strong> ${metadata.fileSize} bytes | <strong>Timestamps:</strong> ${metadata.totalTimestamps} | <strong>Processing:</strong> ${metadata.processingTimeMs}ms</p>
    </div>
    <footer style="text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; color: #666;">
        <p><em>Generated by Azure Functions VTT Meeting Transcript Processor</em></p>
        <p>File: ${result.actualFile} | Processed: ${metadata.processedAt}</p>
    </footer>
</body>
</html>`;

    return {
        ...result,
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
**Video Link:** [🎥 Watch this moment](${point.videoLink})

---`).join('\n\n')}

## 📊 Processing Information

- **File Size:** ${Math.round(metadata.fileSize / 1024)}KB
- **Timestamps:** ${metadata.totalTimestamps}
- **Processing Time:** ${metadata.processingTimeMs}ms

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
            speaker: point.speaker,
            videoLink: point.videoLink
        })),
        processingTimeMs: metadata.processingTimeMs,
        fileSize: metadata.fileSize,
        outputFormat: 'summary',
        processedAt: metadata.processedAt
    };
}

function parseVttTimestamps(vttContent) {
    if (!vttContent) return [];

    const contentBlocks = [];
    const lines = vttContent.split('\n');
    let currentBlock = null;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // Check if line contains timestamp
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

function extractKeyPoints(summary, timestampBlocks, videoUrl) {
    if (!summary || !timestampBlocks) return [];
    const keyPoints = [];
    const summaryLines = summary.split('\n').filter(line =>
        line.trim().length > 0 && (line.includes('**') || line.includes('###'))
    );

    summaryLines.forEach((line, index) => {
        if (index < timestampBlocks.length) {
            const title = line.replace(/[#*]/g, '').trim();
            const block = timestampBlocks[index];
            if (block && title.length > 10) {
                keyPoints.push({
                    title,
                    description: `Key discussion point from ${block.speaker || 'meeting'}`,
                    timestamp: block.timestamp,
                    videoLink: createVideoLink(block.timestamp, videoUrl),
                    speaker: block.speaker
                });
            }
        }
    });
    return keyPoints;
}

function createVideoLink(timestamp, videoUrl) {
    const [hours, minutes, seconds] = timestamp.split(':');
    return `${videoUrl}#t=${hours}h${minutes}m${seconds}s`;
}

function chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
        chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
}

// Export functions for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        parseVttTimestamps,
        extractMeetingMetadata,
        extractKeyPoints,
        createVideoLink
    };
}