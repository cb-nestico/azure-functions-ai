const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { OpenAI } = require('openai');

app.http('ProcessVttFile', {
    methods: ['GET', 'POST'],
    authLevel: 'function', // Changed from 'anonymous' to 'function' for production
    handler: async (request, context) => {
        const startTime = Date.now();
        context.log('🎯 Azure Function triggered: ProcessVttFile');

        try {
            // 1. Parse request
            let fileName;
            
            if (request.method === 'GET') {
                fileName = request.query.get('name');
                context.log(`📥 GET request - fileName: ${fileName}`);
            } else {
                const body = await request.text();
                context.log(`📥 POST request - raw body length: ${body?.length || 0}`);
                
                if (!body || body.trim() === '') {
                    throw new Error('Request body is empty');
                }
                
                try {
                    const requestData = JSON.parse(body);
                    fileName = requestData.name;
                    context.log(`📥 Parsed JSON - fileName: ${fileName}`);
                } catch (parseError) {
                    throw new Error(`Invalid JSON format: ${parseError.message}`);
                }
            }

            if (!fileName) {
                throw new Error('File name is required (provide "name" parameter)');
            }

            context.log(`🎥 Processing file: ${fileName}`);

            // 2. Configuration - FIXED to match your Azure environment variables
            const config = {
                tenantId: process.env.TENANT_ID,
                clientId: process.env.CLIENT_ID,
                clientSecret: process.env.CLIENT_SECRET,
                
                // Use East US 2 endpoint and key (your working configuration)
                openaiEndpoint: process.env.OPENAI_ENDPOINT, // https://ai-teams-eastus2.openai.azure.com/
                openaiKey: process.env.OPENAI_KEY, // East US 2 key
                deployment: process.env.OPENAI_DEPLOYMENT || 'gpt-4o-text',
                
                sharepointDriveId: process.env.SHAREPOINT_DRIVE_ID,
                sharepointSiteUrl: process.env.SHAREPOINT_SITE_URL
            };

            // 3. Validate configuration
            const requiredConfig = ['tenantId', 'clientId', 'clientSecret', 'openaiKey', 'openaiEndpoint', 'sharepointDriveId'];
            const missingConfig = requiredConfig.filter(key => !config[key]);
            
            if (missingConfig.length > 0) {
                throw new Error(`Missing required configuration: ${missingConfig.join(', ')}`);
            }

            context.log('✅ Configuration validated');
            context.log(`🔧 Using OpenAI endpoint: ${config.openaiEndpoint}`);
            context.log(`🔧 Using deployment: ${config.deployment}`);

            // 4. Initialize OpenAI client with correct Azure configuration
            const openaiClient = new OpenAI({
                apiKey: config.openaiKey,
                baseURL: `${config.openaiEndpoint}/openai/deployments/${config.deployment}`,
                defaultQuery: { 'api-version': '2024-08-01-preview' },
                defaultHeaders: {
                    'api-key': config.openaiKey,
                },
            });

            context.log('✅ OpenAI client initialized');

            // 5. Initialize Microsoft Graph client
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

            // 6. Find VTT files in SharePoint
            context.log(`🔍 Searching for VTT files in drive: ${config.sharepointDriveId}`);
            
            const driveItems = await graphClient
                .api(`/drives/${config.sharepointDriveId}/root/children`)
                .get();

            context.log(`📋 Found ${driveItems.value.length} items in drive root`);
            
            const vttFiles = [];
            
            // Search root level
            for (const item of driveItems.value) {
                if (item.file && item.name.toLowerCase().endsWith('.vtt')) {
                    vttFiles.push(item);
                    context.log(`  📄 VTT: ${item.name} (${item.size} bytes)`);
                } else if (item.folder) {
                    context.log(`  📁 Folder: ${item.name}`);
                    
                    // Search subfolders for VTT files
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

            // 7. Find target file
            let targetFile = vttFiles.find(file => 
                file.name.toLowerCase() === fileName.toLowerCase()
            );

            if (!targetFile) {
                // Try partial match
                targetFile = vttFiles.find(file => 
                    file.name.toLowerCase().includes(fileName.replace('.vtt', '').toLowerCase())
                );
                
                if (targetFile) {
                    context.log(`📄 Found partial match: ${targetFile.name}`);
                }
            }

            if (!targetFile) {
                const availableFiles = vttFiles.map(f => f.name).slice(0, 10).join(', ');
                throw new Error(`File not found: ${fileName}. Available VTT files: ${availableFiles}${vttFiles.length > 10 ? '...' : ''}`);
            }

            context.log(`✅ Found file: ${targetFile.name} (${targetFile.size} bytes)`);

            // 8. Download VTT file content using two-step process
            context.log('📥 Downloading VTT file content...');
            
            // Get download URL
            const downloadUrlResponse = await graphClient
                .api(`/drives/${config.sharepointDriveId}/items/${targetFile.id}`)
                .select('@microsoft.graph.downloadUrl')
                .get();

            const downloadUrl = downloadUrlResponse['@microsoft.graph.downloadUrl'];
            context.log('✅ Got download URL');

            // Download content using fetch
            const response = await fetch(downloadUrl);
            
            if (!response.ok) {
                throw new Error(`Download failed: ${response.status} ${response.statusText}`);
            }

            let vttContent = await response.text();
            context.log(`✅ Downloaded content: ${vttContent.length} characters`);

            // Log content preview for verification
            const previewLines = vttContent.split('\n').slice(0, 5);
            context.log('Content preview:', previewLines);

            // 9. Process content for OpenAI (handle large files)
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

            // 10. Parse VTT timestamps (NEW FEATURE)
            context.log('🔍 Parsing VTT timestamps...');
            const timestampBlocks = parseVttTimestamps(vttContent);
            context.log(`✅ Parsed ${timestampBlocks.length} timestamp blocks`);

            // 11. Extract metadata with NOTE parsing (NEW FEATURE)
            const metadata = extractMeetingMetadata(vttContent, targetFile, config.sharepointSiteUrl);
            context.log(`Meeting: ${metadata.title}`);

            // 12. Generate AI summary with Dynamics 365 CRM training focus (ENHANCED)
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

            // 13. Extract key points with timestamps (NEW FEATURE)
            const keyPoints = extractKeyPoints(summary, timestampBlocks, metadata.videoUrl);

            // 14. Format final response with enhanced structure
            const result = {
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
                    processingTimeMs: Date.now() - startTime
                }
            };

            return {
                status: 200,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(result)
            };

        } catch (error) {
            context.log.error('❌ Function execution failed:', error.message);
            context.log.error('Stack trace:', error.stack);
            
            return {
                status: 500,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    error: 'VTT processing failed',
                    message: error.message,
                    endpoint: process.env.OPENAI_ENDPOINT || 'not configured',
                    deployment: process.env.OPENAI_DEPLOYMENT || 'not configured',
                    timestamp: new Date().toISOString()
                })
            };
        }
    }
});

// NEW: Parse VTT timestamps
function parseVttTimestamps(vttContent) {
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

// NEW: Extract meeting metadata with NOTE parsing
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

// NEW: Extract key points with video links
function extractKeyPoints(summary, timestampBlocks, videoUrl) {
    const keyPoints = [];
    
    // Simple extraction - match summary sections to timestamps
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

// NEW: Convert timestamp to video link format
function createVideoLink(timestamp, videoUrl) {
    const [hours, minutes, seconds] = timestamp.split(':');
    return `${videoUrl}#t=${hours}h${minutes}m${seconds}s`;
}