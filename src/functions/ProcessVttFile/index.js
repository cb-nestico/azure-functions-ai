const { app } = require('@azure/functions');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const { OpenAIClient, AzureKeyCredential } = require('@azure/openai');

app.http('ProcessVttFile', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log('🎯 Azure Function triggered: ProcessVttFile');

        try {
            let fileName;

            // Handle both GET and POST requests
            if (request.method === 'GET') {
                fileName = request.query.get('name');
                context.log(`📥 GET request - fileName: ${fileName}`);
            } else {
                // Parse POST request body
                const body = await request.text();
                context.log(`📥 POST request - raw body: ${body}`);
                
                if (!body || body.trim() === '') {
                    throw new Error('Request body is empty');
                }

                try {
                    const requestData = JSON.parse(body);
                    fileName = requestData.name;
                    context.log(`📥 Parsed JSON - fileName: ${fileName}`);
                } catch (parseError) {
                    context.log(`❌ JSON parse error: ${parseError.message}`);
                    throw new Error(`Invalid JSON format: ${parseError.message}`);
                }
            }

            if (!fileName) {
                throw new Error('File name is required (provide "name" parameter)');
            }

            context.log(`🎥 Processing file: ${fileName}`);

            // Initialize authentication
            const credential = new ClientSecretCredential(
                process.env.TENANT_ID,
                process.env.CLIENT_ID,
                process.env.CLIENT_SECRET
            );

            // Initialize Graph client
            const accessToken = await credential.getToken(['https://graph.microsoft.com/.default']);
            const graphClient = Client.init({
                authProvider: async (done) => done(null, accessToken.token)
            });

            context.log('✅ Graph client initialized');

            // List files in the drive with debugging
            context.log(`🔍 Listing files in drive: ${process.env.SHAREPOINT_DRIVE_ID}`);
            
            const driveItems = await graphClient
                .api(`/drives/${process.env.SHAREPOINT_DRIVE_ID}/root/children`)
                .get();

            context.log(`📋 Found ${driveItems.value.length} items in drive root:`);
            
            const vttFiles = [];
            for (const item of driveItems.value) {
                if (item.file && item.name.toLowerCase().endsWith('.vtt')) {
                    vttFiles.push(item);
                    context.log(`  📄 VTT: ${item.name} (${item.size} bytes)`);
                } else if (item.folder) {
                    context.log(`  📁 Folder: ${item.name}`);
                    
                    // Check subfolder for VTT files
                    try {
                        const folderItems = await graphClient
                            .api(`/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${item.id}/children`)
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

            // Try to find the exact file
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
                // List available VTT files for debugging
                const availableVttFiles = vttFiles.map(f => f.name).join(', ');
                throw new Error(`File not found: ${fileName}. Available VTT files: ${availableVttFiles || 'none'}`);
            }

            context.log(`✅ Found file: ${targetFile.name} (${targetFile.size} bytes)`);

            // Download file content using HTTP fetch approach
            context.log(`🔄 Downloading file content for: ${targetFile.name}`);
            
            try {
                // Get the download URL from Graph API
                const downloadUrlResponse = await graphClient
                    .api(`/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${targetFile.id}`)
                    .select('@microsoft.graph.downloadUrl')
                    .get();

                const downloadUrl = downloadUrlResponse['@microsoft.graph.downloadUrl'];
                context.log(`🔗 Download URL obtained: ${downloadUrl ? 'Yes' : 'No'}`);

                if (!downloadUrl) {
                    throw new Error('Could not obtain download URL from Microsoft Graph');
                }

                // Use fetch to download the file content directly
                const response = await fetch(downloadUrl);
                
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }

                let vttContent = await response.text();
                context.log(`✅ Downloaded VTT file (${vttContent.length} characters)`);

                // Validate file content
                if (vttContent.length < 100) {
                    context.log(`⚠️ Warning: File content seems too short. Content preview: ${vttContent.substring(0, 200)}`);
                    throw new Error(`File content is too short (${vttContent.length} characters). Expected ${targetFile.size} bytes.`);
                }

                // Show a preview of the content for debugging
                const preview = vttContent.substring(0, 300).replace(/\n/g, '\\n');
                context.log(`📄 VTT Content Preview: ${preview}...`);

                // RATE LIMITING OPTIMIZATION: Truncate large VTT files
                const MAX_TOKENS = 8000; // Conservative limit for your tier
                const CHARS_PER_TOKEN = 4; // Approximate ratio
                const maxChars = MAX_TOKENS * CHARS_PER_TOKEN;

                if (vttContent.length > maxChars) {
                    context.log(`⚠️ Large file detected (${vttContent.length} chars). Truncating to ${maxChars} chars to avoid rate limits.`);
                    vttContent = vttContent.substring(0, maxChars);
                    context.log(`✂️ Truncated content to ${vttContent.length} characters`);
                }

                // Now process with Azure OpenAI
                context.log('🤖 Processing with Azure OpenAI...');
                
                const openaiClient = new OpenAIClient(
                    process.env.OPENAI_ENDPOINT,
                    new AzureKeyCredential(process.env.OPENAI_KEY)
                );

                const messages = [
                    {
                        role: "system",
                        content: "You are an expert meeting analyst. Analyze VTT transcripts and provide comprehensive summaries with key points, action items, and next steps."
                    },
                    {
                        role: "user",
                        content: `Please analyze this VTT meeting transcript and provide a detailed summary:

## Meeting Analysis Request
- Extract key discussion points
- Identify action items and assignments  
- Note important decisions made
- Highlight next steps and follow-ups
- Summarize participant contributions

## VTT Transcript:
${vttContent}`
                    }
                ];

                context.log(`🤖 Calling OpenAI deployment: ${process.env.OPENAI_DEPLOYMENT}`);
                context.log(`📊 Estimated tokens: ~${Math.ceil(vttContent.length / CHARS_PER_TOKEN)}`);
                
                const result = await openaiClient.getChatCompletions(
                    process.env.OPENAI_DEPLOYMENT,
                    messages,
                    {
                        maxTokens: 1500, // Reduced to stay within limits
                        temperature: 0.3
                    }
                );

                const summary = result.choices[0].message.content;
                context.log(`✅ Summary generated (${summary.length} characters)`);

                return {
                    status: 200,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        success: true,
                        file: fileName,
                        actualFile: targetFile.name,
                        summary: summary,
                        metadata: {
                            endpoint: process.env.OPENAI_ENDPOINT,
                            deployment: process.env.OPENAI_DEPLOYMENT,
                            fileSize: targetFile.size,
                            originalContentLength: vttContent.length,
                            truncated: targetFile.size > maxChars,
                            estimatedTokens: Math.ceil(vttContent.length / CHARS_PER_TOKEN),
                            processedAt: new Date().toISOString()
                        }
                    })
                };

            } catch (downloadError) {
                context.log(`❌ Download Error: ${downloadError.message}`);
                throw new Error(`File download failed: ${downloadError.message}`);
            }

        } catch (error) {
            context.log('❌ Error:', error.message);
            context.log('❌ Error stack:', error.stack);
            
            return {
                status: 500,
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    error: 'Processing failed',
                    details: error.message,
                    endpoint: process.env.OPENAI_ENDPOINT,
                    deployment: process.env.OPENAI_DEPLOYMENT
                })
            };
        }
    }
});