const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { OpenAI } = require('openai');

// ✅ ADD THIS REGISTRATION AT THE TOP
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
            context.log.error('❌ Function execution failed:', error.message);
            context.log.error('Stack trace:', error.stack);
            
            return {
                status: 500,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    success: false,
                    error: 'Function execution failed',
                    message: error.message,
                    timestamp: new Date().toISOString(),
                    processingTimeMs: Date.now() - startTime
                })
            };
        }
    }
});

// ✅ YOUR EXISTING FUNCTIONS CONTINUE HERE (no changes needed)
async function processSingleFile(context, fileName, outputFormat = 'json') {
    const result = await processSingleVttFile(context, fileName, outputFormat);
    return {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(result)
    };
}

async function processBatchFiles(context, fileNames, outputFormat = 'json') {
    const results = [];
    const batchStartTime = Date.now();
    const concurrencyLimit = 3;
    
    // Split files into batches
    const batches = chunkArray(fileNames, concurrencyLimit);
    context.log(`Processing ${fileNames.length} files in ${batches.length} batches`);
    
    for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
        const batch = batches[batchIndex];
        context.log(`Processing batch ${batchIndex + 1}/${batches.length}`);
        
        const batchPromises = batch.map(async (fileName) => {
            const fileStartTime = Date.now();
            try {
                const fileResult = await processSingleVttFile(context, fileName, outputFormat);
                const fileProcessingTime = Date.now() - fileStartTime;
                
                return {
                    fileName: fileName,
                    success: true,
                    processingTimeMs: fileProcessingTime,
                    ...fileResult
                };
            } catch (error) {
                const fileProcessingTime = Date.now() - fileStartTime;
                context.log.error(`Batch processing failed for ${fileName}:`, error);
                
                return {
                    fileName: fileName,
                    success: false,
                    error: error.message,
                    processingTimeMs: fileProcessingTime
                };
            }
        });
        
        const batchResults = await Promise.all(batchPromises);
        results.push(...batchResults);
        
        // Add delay between batches
        if (batchIndex < batches.length - 1) {
            context.log('Waiting 2 seconds before next batch...');
            await new Promise(resolve => setTimeout(resolve, 2000));
        }
    }
    
    const batchTotalTime = Date.now() - batchStartTime;
    const successfulFiles = results.filter(r => r.success);
    
    return {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            success: true,
            batchMode: true,
            processedFiles: results.length,
            successfulFiles: successfulFiles.length,
            failedFiles: results.length - successfulFiles.length,
            results: results,
            metadata: {
                batchProcessingTimeMs: batchTotalTime,
                averageTimePerFile: Math.round(batchTotalTime / fileNames.length),
                concurrencyLimit: concurrencyLimit,
                totalBatches: batches.length,
                outputFormat: outputFormat,
                timestamp: new Date().toISOString()
            }
        })
    };
}

// ✅ KEEP ALL YOUR EXISTING FUNCTIONS UNCHANGED
async function processSingleVttFile(context, fileName, outputFormat = 'json') {
    const processingStartTime = Date.now();
    context.log(`🎬 Starting VTT processing for: ${fileName}`);
    
    // Configuration
    const config = {
        tenantId: process.env.TENANT_ID,
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        openaiEndpoint: process.env.OPENAI_ENDPOINT,
        openaiKey: process.env.OPENAI_KEY,
        deployment: process.env.OPENAI_DEPLOYMENT || 'gpt-4o-text',
        sharepointDriveId: process.env.SHAREPOINT_DRIVE_ID,
        sharepointSiteUrl: process.env.SHAREPOINT_SITE_URL
    };

    // Validate configuration
    const requiredConfig = ['tenantId', 'clientId', 'clientSecret', 'openaiKey', 'openaiEndpoint', 'sharepointDriveId'];
    const missingConfig = requiredConfig.filter(key => !config[key]);
    
    if (missingConfig.length > 0) {
        throw new Error(`Missing required configuration: ${missingConfig.join(', ')}`);
    }

    context.log('✅ Configuration validated');

    // Initialize OpenAI client
    const openaiClient = new OpenAI({
        apiKey: config.openaiKey,
        baseURL: `${config.openaiEndpoint}/openai/deployments/${config.deployment}`,
        defaultQuery: { 'api-version': '2024-08-01-preview' },
        defaultHeaders: {
            'api-key': config.openaiKey,
        },
    });

    // Initialize Microsoft Graph client
    const credential = new ClientSecretCredential(
        config.tenantId,
        config.clientId,
        config.clientSecret
    );

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: ['https://graph.microsoft.com/.default']
    });

    const graphClient = Client.initWithMiddleware({ authProvider });
    context.log('✅ Graph client initialized');

    // Find VTT files in SharePoint
    context.log(`🔍 Searching for VTT files in drive: ${config.sharepointDriveId}`);
    
    const driveItems = await graphClient
        .api(`/drives/${config.sharepointDriveId}/root/children`)
        .get();

    const vttFiles = [];
    
    // Search root level and folders
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
                context.log(`    ❌ Cannot access folder ${item.name}: ${folderError.message}`);
            }
        }
    }

    context.log(`🎬 Total VTT files found: ${vttFiles.length}`);

    // Find target file
    let targetFile = vttFiles.find(file => 
        file.name.toLowerCase() === fileName.toLowerCase()
    );

    if (!targetFile) {
        // Try partial match
        targetFile = vttFiles.find(file => 
            file.name.toLowerCase().includes(fileName.replace('.vtt', '').toLowerCase())
        );
    }

    if (!targetFile) {
        const availableFiles = vttFiles.map(f => f.name).slice(0, 10).join(', ');
        throw new Error(`File not found: ${fileName}. Available VTT files: ${availableFiles}${vttFiles.length > 10 ? '...' : ''}`);
    }

    context.log(`✅ Found file: ${targetFile.name} (${targetFile.size} bytes)`);

    // Download VTT file content
    const downloadUrlResponse = await graphClient
        .api(`/drives/${config.sharepointDriveId}/items/${targetFile.id}`)
        .select('@microsoft.graph.downloadUrl')
        .get();

    const downloadUrl = downloadUrlResponse['@microsoft.graph.downloadUrl'];
    const response = await fetch(downloadUrl);
    
    if (!response.ok) {
        throw new Error(`Download failed: ${response.status} ${response.statusText}`);
    }

    let vttContent = await response.text();
    context.log(`✅ Downloaded content: ${vttContent.length} characters`);

    // Handle large files
    const MAX_TOKENS = 8000;
    const CHARS_PER_TOKEN = 4;
    const maxChars = MAX_TOKENS * CHARS_PER_TOKEN;

    let truncated = false;
    const originalLength = vttContent.length;

    if (vttContent.length > maxChars) {
        vttContent = vttContent.substring(0, maxChars);
        truncated = true;
        context.log(`⚠️ Content truncated from ${originalLength} to ${vttContent.length} characters`);
    }

    // Parse VTT timestamps
    const timestampBlocks = parseVttTimestamps(vttContent);
    context.log(`✅ Parsed ${timestampBlocks.length} timestamp blocks`);

    // Extract metadata
    const metadata = extractMeetingMetadata(vttContent, targetFile, config.sharepointSiteUrl);
    context.log(`Meeting: ${metadata.title}`);

    // Generate AI summary
    context.log('🤖 Generating AI-powered meeting summary...');

    const trainingPrompt = `You are an expert in Dynamics 365 CRM training analysis. Analyze this meeting transcript and provide:

1. **Training Topics Covered**: Identify specific Dynamics 365 CRM features, functions, or processes that were taught or discussed.

2. **Key Learning Points**: For each topic, provide:
   - A clear, concise title (e.g., "XRM Toolbox Usage", "Environment Access Management")
   - A brief 1-2 sentence description of what was taught or demonstrated
   - Any best practices or tips shared

3. **Action Items**: Identify any homework, practice exercises, or follow-up tasks assigned

4. **Q&A Highlights**: Note important questions asked and answers provided

5. **Next Steps**: Any upcoming training sessions or topics mentioned

Focus on actionable learning content for CRM training reference.

Meeting Title: ${metadata.title}
Transcript Content:
${vttContent}`;

    const completion = await openaiClient.chat.completions.create({
        model: config.deployment,
        messages: [
            {
                role: "user",
                content: trainingPrompt
            }
        ],
        max_tokens: 2000,
        temperature: 0.3
    });

    const summary = completion.choices[0].message.content;
    context.log('✅ AI summary generated successfully');

    // Extract key points with timestamps
    const keyPoints = extractKeyPoints(summary, timestampBlocks, metadata.videoUrl);

    // Create base result
    const baseResult = {
        success: true,
        meetingTitle: metadata.title,
        date: metadata.date,
        videoUrl: metadata.videoUrl,
        file: fileName,
        actualFile: targetFile.name,
        summary: summary,
        keyPoints: keyPoints,
        timestampBlocks: timestampBlocks.slice(0, 10), // First 10 for reference
        metadata: {
            endpoint: config.openaiEndpoint,
            deployment: config.deployment,
            fileSize: targetFile.size,
            originalContentLength: originalLength,
            processedContentLength: vttContent.length,
            truncated: truncated,
            estimatedTokens: Math.ceil(vttContent.length / CHARS_PER_TOKEN),
            totalTimestamps: timestampBlocks.length,
            totalKeyPoints: keyPoints.length,
            processedAt: new Date().toISOString(),
            processingTimeMs: Date.now() - processingStartTime
        }
    };

    // Apply output formatting
    return await applyOutputFormat(context, baseResult, outputFormat);
}

// ✅ KEEP ALL YOUR HELPER FUNCTIONS UNCHANGED
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
        ${keyPoints.map((point, index) => `
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
        meetingTitle: meetingTitle,
        summary: summary,
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
            if (currentBlock) {
                contentBlocks.push(currentBlock);
            }
            currentBlock = {
                timestamp: timestampMatch[1], // HH:MM:SS format
                content: '',
                speaker: null
            };
        } else if (currentBlock && line.length > 0) {
            // Extract speaker and content
            const speakerMatch = line.match(/<v\s+([^>]+)>(.+)<\/v>/);
            if (speakerMatch) {
                currentBlock.speaker = speakerMatch[1];
                currentBlock.content += speakerMatch[2] + ' ';
            } else {
                currentBlock.content += line + ' ';
            }
        }
    }
    
    if (currentBlock) {
        contentBlocks.push(currentBlock);
    }
    
    return contentBlocks;
}

function extractMeetingMetadata(vttContent, fileMetadata, sharepointSiteUrl) {
    // Extract meeting title from NOTE line
    const noteMatch = vttContent.match(/NOTE\s+(.+)/);
    const meetingTitle = noteMatch ? noteMatch[1].trim() : 
                        fileMetadata.name.replace('.vtt', '').replace(/[-_]/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
    
    // Create video URL from SharePoint site
    const videoUrl = sharepointSiteUrl ? 
                    `${sharepointSiteUrl}/Shared%20Documents/${fileMetadata.name.replace('.vtt', '')}` :
                    "https://yourtenant.sharepoint.com/video-placeholder";
    
    return {
        title: meetingTitle,
        videoUrl: videoUrl,
        date: new Date().toISOString().split('T')[0], // YYYY-MM-DD format
        filename: fileMetadata.name
    };
}

function extractKeyPoints(summary, timestampBlocks, videoUrl) {
    if (!summary || !timestampBlocks) return [];
    
    const keyPoints = [];
    
    // Extract titles with markdown formatting
    const summaryLines = summary.split('\n').filter(line => 
        line.trim().length > 0 && (line.includes('**') || line.includes('###'))
    );
    
    summaryLines.forEach((line, index) => {
        if (index < timestampBlocks.length) {
            const title = line.replace(/[#*]/g, '').trim();
            const block = timestampBlocks[index];
            
            if (block && title.length > 10) {
                keyPoints.push({
                    title: title,
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