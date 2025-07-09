const { app } = require('@azure/functions');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const { OpenAIClient, AzureKeyCredential } = require('@azure/openai');

// ✨ NEW: VTT timestamp extraction and parsing
function parseVttTimestamps(vttContent) {
    const contentBlocks = [];
    const lines = vttContent.split('\n');
    let currentBlock = null;
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Check if line contains timestamp (VTT format: HH:MM:SS.mmm --> HH:MM:SS.mmm)
        const timestampMatch = line.match(/(\d{2}:\d{2}:\d{2})\.\d{3}\s*-->\s*(\d{2}:\d{2}:\d{2})\.\d{3}/);
        if (timestampMatch) {
            // Save previous block if exists and has content
            if (currentBlock && currentBlock.content.trim()) {
                contentBlocks.push(currentBlock);
            }
            
            // Start new block
            currentBlock = {
                startTime: timestampMatch[1], // HH:MM:SS format
                endTime: timestampMatch[2],   // HH:MM:SS format
                content: '',
                speaker: null
            };
        } else if (currentBlock && line.length > 0) {
            // Extract speaker and content from VTT speaker tags
            const speakerMatch = line.match(/<v\s+([^>]+)>(.+)<\/v>/);
            if (speakerMatch) {
                currentBlock.speaker = speakerMatch[1].trim();
                currentBlock.content += speakerMatch[2].trim() + ' ';
            } else if (!line.match(/^\d+$/)) { // Skip sequence numbers
                currentBlock.content += line + ' ';
            }
        }
    }
    
    // Don't forget the last block
    if (currentBlock && currentBlock.content.trim()) {
        contentBlocks.push(currentBlock);
    }
    
    return contentBlocks;
}

// ✨ NEW: Convert timestamp to video link format
function createVideoLink(timestamp, videoUrl) {
    const [hours, minutes, seconds] = timestamp.split(':');
    return `${videoUrl}#t=${hours}h${minutes}m${seconds}s`;
}

// ✨ NEW: Extract meeting metadata from VTT content and file metadata
function extractMeetingMetadata(vttContent, fileMetadata) {
    // Extract meeting title from NOTE line (VTT format: NOTE Title goes here)
    const noteMatch = vttContent.match(/^NOTE\s+(.+)$/m);
    const meetingTitle = noteMatch ? noteMatch[1].trim() : "Dynamics 365 CRM Training";
    
    // Extract video URL from SharePoint metadata or construct from file URL
    let videoUrl = "https://yourtenant.sharepoint.com/video-placeholder";
    
    // Try to extract video URL from various SharePoint metadata fields
    if (fileMetadata) {
        videoUrl = fileMetadata.VideoURL || 
                  fileMetadata.videoUrl ||
                  fileMetadata.webUrl ||
                  fileMetadata.parentReference?.path || 
                  videoUrl;
    }
    
    return {
        title: meetingTitle,
        videoUrl: videoUrl,
        date: new Date().toISOString().split('T')[0], // YYYY-MM-DD format
        filename: fileMetadata?.name || 'unknown.vtt'
    };
}

// ✨ NEW: Enhanced AI prompt for Dynamics 365 CRM training focus
function createTrainingAnalysisPrompt(vttContent) {
    return `You are an expert in Dynamics 365 CRM training analysis. Analyze this meeting transcript and extract specific training content.

**Instructions:**
1. **Identify Training Topics**: Focus on Dynamics 365 CRM features, functions, or processes discussed
2. **Extract Key Learning Points**: For each topic, provide:
   - A clear, concise title (e.g., "Creating Custom Fields", "Lead Management Process")
   - A brief 1-2 sentence description of what was taught or demonstrated
   - Any best practices or tips shared
3. **Action Items**: Identify homework, practice exercises, or follow-up tasks
4. **Q&A Highlights**: Note important questions and answers
5. **Next Steps**: Upcoming training sessions or topics mentioned

**Output Format:**
Structure your response as clearly organized sections with specific topics that can be linked to timestamps. Focus on actionable learning content for CRM training reference.

**Meeting Transcript:**
${vttContent}`;
}

// ✨ OPTIMIZED: Helper functions for better content analysis
function getTimeDifference(time1, time2) {
    const [h1, m1, s1] = time1.split(':').map(Number);
    const [h2, m2, s2] = time2.split(':').map(Number);
    
    const seconds1 = h1 * 3600 + m1 * 60 + s1;
    const seconds2 = h2 * 3600 + m2 * 60 + s2;
    
    return Math.abs(seconds2 - seconds1);
}

function extractTopicsFromSummary(summary) {
    const topics = [];
    const lines = summary.split('\n');
    
    lines.forEach(line => {
        // Look for topic headers in AI summary
        if (line.match(/^#+\s+/) || line.match(/^\d+\.\s+\*\*.*\*\*/)) {
            topics.push(line.replace(/^#+\s+/, '').replace(/^\d+\.\s+\*\*/, '').replace(/\*\*.*$/, '').trim());
        }
    });
    
    return topics;
}

function findMatchingTopic(content, topics) {
    const contentLower = content.toLowerCase();
    
    for (const topic of topics) {
        const topicWords = topic.toLowerCase().split(' ');
        const matchCount = topicWords.filter(word => 
            word.length > 3 && contentLower.includes(word)
        ).length;
        
        if (matchCount >= Math.min(2, topicWords.length * 0.5)) {
            return topic;
        }
    }
    
    return null;
}

function generateTopicTitle(content) {
    // Extract key phrases for topic titles
    const keyPhrases = [
        'access', 'permission', 'login', 'security', 'database', 'azure', 'data studio',
        'crm', 'dynamics', 'environment', 'production', 'development', 'migration',
        'document', 'license', 'filter', 'view', 'query', 'sql', 'cds', 'sharepoint'
    ];
    
    const contentLower = content.toLowerCase();
    const foundPhrases = keyPhrases.filter(phrase => contentLower.includes(phrase));
    
    if (foundPhrases.length > 0) {
        return `Discussion on ${foundPhrases.slice(0, 3).join(', ')}`;
    }
    
    return 'General Discussion';
}

function categorizeContent(content) {
    const contentLower = content.toLowerCase();
    
    if (contentLower.includes('access') || contentLower.includes('permission') || contentLower.includes('security')) {
        return 'Security & Access';
    }
    if (contentLower.includes('crm') || contentLower.includes('dynamics')) {
        return 'CRM Features';
    }
    if (contentLower.includes('database') || contentLower.includes('azure') || contentLower.includes('data')) {
        return 'Data Management';
    }
    if (contentLower.includes('license') || contentLower.includes('document')) {
        return 'Licensing & Tools';
    }
    if (contentLower.includes('filter') || contentLower.includes('view') || contentLower.includes('query')) {
        return 'UI & Navigation';
    }
    
    return 'General Training';
}

function truncateDescription(content, maxLength) {
    if (content.length <= maxLength) return content;
    
    const truncated = content.substring(0, maxLength);
    const lastSpace = truncated.lastIndexOf(' ');
    
    return lastSpace > maxLength * 0.8 
        ? truncated.substring(0, lastSpace) + '...'
        : truncated + '...';
}

// ✨ OPTIMIZED: Enhanced key points extraction with meaningful filtering
function extractKeyPointsFromSummary(summary, timestampBlocks) {
    const keyPoints = [];
    
    // Filter for meaningful content blocks (longer conversations, not short utterances)
    const meaningfulBlocks = timestampBlocks.filter(block => {
        if (!block.content || block.content.trim().length < 80) return false;
        
        // Filter out common filler words and short responses
        const content = block.content.toLowerCase().trim();
        const fillerPhrases = ['yeah', 'ok', 'uh', 'um', 'right', 'sure', 'thanks', 'thank you', 'mm-hmm', 'uh-huh'];
        const isMainlyFiller = fillerPhrases.some(phrase => 
            content.includes(phrase) && content.length < 100
        );
        
        // Skip if mostly numbers, IDs, or technical gibberish
        if (content.match(/^[0-9a-f-]{20,}/) || content.match(/^\d+[-\d]*$/)) return false;
        
        return !isMainlyFiller;
    });
    
    // Group consecutive blocks by speaker for better context
    const groupedBlocks = [];
    let currentGroup = null;
    
    for (const block of meaningfulBlocks) {
        const timeDiff = currentGroup ? getTimeDifference(currentGroup.endTime, block.startTime) : 60;
        
        if (!currentGroup || currentGroup.speaker !== block.speaker || timeDiff > 30) {
            
            if (currentGroup && currentGroup.content.trim().length > 150) {
                groupedBlocks.push(currentGroup);
            }
            
            currentGroup = {
                speaker: block.speaker,
                startTime: block.startTime,
                endTime: block.endTime,
                content: block.content,
                topics: []
            };
        } else {
            currentGroup.content += ' ' + block.content;
            currentGroup.endTime = block.endTime;
        }
    }
    if (currentGroup && currentGroup.content.trim().length > 150) {
        groupedBlocks.push(currentGroup);
    }
    
    // Extract key topics from AI summary
    const summaryTopics = extractTopicsFromSummary(summary);
    
    // Match grouped blocks with AI-identified topics and limit to most meaningful points
    groupedBlocks.slice(0, 15).forEach((group, index) => { // Limit to 15 most meaningful points
        const relevantTopic = findMatchingTopic(group.content, summaryTopics);
        const topicCategory = categorizeContent(group.content);
        
        // Generate meaningful title
        const topicTitle = relevantTopic || generateTopicTitle(group.content);
        const speakerName = group.speaker || 'Participant';
        
        keyPoints.push({
            title: `${speakerName}: ${topicTitle}`,
            description: truncateDescription(group.content, 200),
            timestamp: group.startTime,
            speaker: group.speaker,
            content: group.content.trim(),
            topicType: topicCategory,
            duration: getTimeDifference(group.startTime, group.endTime)
        });
    });
    
    // Sort by timestamp for chronological order
    keyPoints.sort((a, b) => {
        const timeA = a.timestamp.split(':').map(Number);
        const timeB = b.timestamp.split(':').map(Number);
        
        for (let i = 0; i < 3; i++) {
            if (timeA[i] !== timeB[i]) {
                return timeA[i] - timeB[i];
            }
        }
        return 0;
    });
    
    return keyPoints;
}

// ✨ NEW: Format enhanced training output matching requirements
function formatTrainingOutput(summary, timestamps, metadata, keyPoints) {
    return {
        success: true,
        meetingTitle: metadata.title,
        date: metadata.date,
        videoUrl: metadata.videoUrl,
        keyPoints: keyPoints.map(point => ({
            title: point.title,
            description: point.description,
            timestamp: point.timestamp, // HH:MM:SS format
            videoLink: createVideoLink(point.timestamp, metadata.videoUrl),
            speaker: point.speaker,
            topicType: point.topicType,
            duration: `${point.duration}s`
        })),
        summary: summary,
        metadata: {
            originalFile: metadata.filename,
            fileSize: metadata.fileSize,
            processedAt: new Date().toISOString(),
            totalKeyPoints: keyPoints.length,
            totalTimestamps: timestamps.length,
            endpoint: process.env.OPENAI_ENDPOINT,
            deployment: process.env.OPENAI_DEPLOYMENT,
            trainingFocused: true,
            optimized: true
        }
    };
}

// ✨ NEW: Enhanced file metadata retrieval
async function getEnhancedFileMetadata(graphClient, driveId, fileId) {
    try {
        const fileData = await graphClient
            .api(`/drives/${driveId}/items/${fileId}`)
            .select('name,size,lastModifiedDateTime,webUrl,@microsoft.graph.downloadUrl')
            .expand('listItem($select=fields)')
            .get();
            
        return {
            ...fileData,
            VideoURL: fileData.listItem?.fields?.VideoURL,
            customFields: fileData.listItem?.fields
        };
    } catch (error) {
        // Return basic metadata if enhanced retrieval fails
        return { name: 'unknown', size: 0, webUrl: '' };
    }
}

app.http('ProcessVttFile', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        context.log('🎯 Azure Function triggered: ProcessVttFile - Enhanced Training Analysis v2.0');

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

            // ✨ NEW: Get enhanced file metadata for video URL extraction
            const enhancedMetadata = await getEnhancedFileMetadata(graphClient, process.env.SHAREPOINT_DRIVE_ID, targetFile.id);
            context.log(`📊 Enhanced metadata retrieved for: ${enhancedMetadata.name}`);

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

                // ✨ NEW: Parse VTT timestamps and extract meeting metadata
                context.log('🕐 Parsing VTT timestamps...');
                const timestampBlocks = parseVttTimestamps(vttContent);
                context.log(`✅ Extracted ${timestampBlocks.length} timestamp blocks`);

                context.log('📋 Extracting meeting metadata...');
                const meetingMetadata = extractMeetingMetadata(vttContent, enhancedMetadata);
                context.log(`✅ Meeting title: "${meetingMetadata.title}"`);
                context.log(`🔗 Video URL: ${meetingMetadata.videoUrl}`);

                // RATE LIMITING OPTIMIZATION: Truncate large VTT files
                const MAX_TOKENS = 8000; // Conservative limit for your tier
                const CHARS_PER_TOKEN = 4; // Approximate ratio
                const maxChars = MAX_TOKENS * CHARS_PER_TOKEN;

                if (vttContent.length > maxChars) {
                    context.log(`⚠️ Large file detected (${vttContent.length} chars). Truncating to ${maxChars} chars to avoid rate limits.`);
                    vttContent = vttContent.substring(0, maxChars);
                    context.log(`✂️ Truncated content to ${vttContent.length} characters`);
                }

                // ✨ NEW: Process with enhanced training-focused Azure OpenAI
                context.log('🤖 Processing with Azure OpenAI - Training Analysis...');
                
                const openaiClient = new OpenAIClient(
                    process.env.OPENAI_ENDPOINT,
                    new AzureKeyCredential(process.env.OPENAI_KEY)
                );

                // ✨ NEW: Use training-specific prompt
                const trainingPrompt = createTrainingAnalysisPrompt(vttContent);

                const messages = [
                    {
                        role: "system",
                        content: "You are a Dynamics 365 CRM training expert who creates detailed, actionable summaries of training sessions for easy reference and review."
                    },
                    {
                        role: "user",
                        content: trainingPrompt
                    }
                ];

                context.log(`🤖 Calling OpenAI deployment: ${process.env.OPENAI_DEPLOYMENT}`);
                context.log(`📊 Estimated tokens: ~${Math.ceil(vttContent.length / CHARS_PER_TOKEN)}`);
                
                const result = await openaiClient.getChatCompletions(
                    process.env.OPENAI_DEPLOYMENT,
                    messages,
                    {
                        maxTokens: 2000, // Increased for more detailed training analysis
                        temperature: 0.3
                    }
                );

                const summary = result.choices[0].message.content;
                context.log(`✅ Training summary generated (${summary.length} characters)`);

                // ✨ OPTIMIZED: Extract key points from summary and match with timestamps
                context.log('🔍 Extracting optimized key training points...');
                const keyPoints = extractKeyPointsFromSummary(summary, timestampBlocks);
                context.log(`✅ Extracted ${keyPoints.length} optimized key training points`);

                // ✨ NEW: Format enhanced output matching requirements
                const enhancedResponse = formatTrainingOutput(
                    summary, 
                    timestampBlocks, 
                    {
                        ...meetingMetadata,
                        fileSize: targetFile.size
                    }, 
                    keyPoints
                );

                context.log(`🎯 Enhanced training analysis complete!`);
                context.log(`📊 Results: ${keyPoints.length} key points (optimized), ${timestampBlocks.length} timestamps`);

                return {
                    status: 200,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(enhancedResponse)
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
                    deployment: process.env.OPENAI_DEPLOYMENT,
                    trainingEnhanced: true,
                    optimized: true
                })
            };
        }
    }
});