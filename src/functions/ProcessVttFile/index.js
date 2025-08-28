const fetch = require('node-fetch');
const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { OpenAI } = require('openai');
//const vttUtils = require('./vtt-utils');
// require helpers as a module object and always call via vttUtils.* to avoid destructure timing issues
const vttUtils = require('./vtt-utils');

process.on('unhandledRejection', (reason) => {
    console.error('Unhandled Rejection:', reason);
});

// Logging shim
function setupLogging(context) {
    try {
        if (!context) return;
        if (typeof context.error !== 'function') context.error = (...args) => console.error(...args);
        if (typeof context.warn !== 'function')  context.warn  = (...args) => console.warn(...args);
        if (typeof context.info !== 'function')  context.info  = (...args) => console.info(...args);
        if (typeof context.log !== 'function') {
            const base = (...args) => console.log(...args);
            base.error = (...args) => context.error(...args);
            base.warn  = (...args) => context.warn(...args);
            base.info  = (...args) => context.info(...args);
            context.log = base;
        }
    } catch { }
}

// HTTP Trigger
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
                if (!body || body.trim() === '') throw new Error('Request body is empty');
                const requestData = JSON.parse(body);
                fileName = requestData.name;
                batchMode = requestData.batchMode || false;
                fileNames = requestData.fileNames || [];
                outputFormat = requestData.outputFormat || 'json';
            }

            if (batchMode && fileNames.length > 1) {
                return await processBatchFiles(context, fileNames, outputFormat);
            } else {
                const singleFile = fileName || (fileNames.length > 0 ? fileNames[0] : null);
                if (!singleFile) throw new Error('File name is required');
                return await processSingleFile(context, singleFile, outputFormat);
            }
        } catch (error) {
            context.log.error('❌ Function execution failed:', error?.message || error);
            return {
                status: 500,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ success: false, error: error?.message || String(error), timestamp: new Date().toISOString(), processingTimeMs: Date.now() - startTime })
            };
        }
    }
});

// processBatchFiles uses vttUtils.chunkArray
async function processBatchFiles(context, fileNames, outputFormat = 'json') {
    const concurrencyLimit = Number(process.env.BATCH_CONCURRENCY || 3);
    const batches = vttUtils.chunkArray(fileNames, concurrencyLimit);
    const results = [];
    for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
        const batch = batches[batchIndex];
        const batchResults = await Promise.allSettled(batch.map(f => processSingleFile(context, f, outputFormat)));
        results.push(...batchResults.map(r => r.value || r.reason));
    }
    return { status: 200, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(results, null, 2) };
}

// processSingleFile delegates to processSingleVttFile and wraps response
async function processSingleFile(context, fileName, outputFormat = 'json') {
    const result = await processSingleVttFile(context, fileName, outputFormat);

    if (
        result &&
        typeof result.status === 'number' &&
        result.headers &&
        (
            typeof result.body === 'string' ||
            (typeof Buffer !== 'undefined' && Buffer.isBuffer(result.body)) ||
            (result.body instanceof Uint8Array)
        )
    ) {
        return result;
    }

    return {
        status: result.status || (result?.success ? 200 : 500),
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(result)
    };
}

// Minimal helpers (kept in index.js where used)
function getVideoUrlFromMetadata(fileMetadata) {
    return (
        fileMetadata?.VideoURL ||
        fileMetadata?.videoUrl ||
        fileMetadata?.video_url ||
        fileMetadata?.webUrl ||
        ""
    );
}

function createVideoLink(timestamp, videoUrl) {
    if (!timestamp || !videoUrl) return null;
    const normalize = (ts) => {
        const noMs = String(ts).split('.')[0].trim();
        const parts = noMs.split(':').map(p => p.trim()).filter(Boolean);
        let h = 0, m = 0, s = 0;
        if (parts.length === 1) s = Number(parts[0]) || 0;
        else if (parts.length === 2) { m = Number(parts[0]) || 0; s = Number(parts[1]) || 0; }
        else { h = Number(parts[0]) || 0; m = Number(parts[1]) || 0; s = Number(parts[2]) || 0; }
        return { h, m, s, totalSeconds: h * 3600 + m * 60 + s };
    };
    const { h, m, s, totalSeconds } = normalize(timestamp);
    const pad = n => String(Math.max(0, Math.floor(n))).padStart(2, '0');
    const sep = videoUrl.includes('?') ? '&' : '?';
    const fragmentCandidate = `${videoUrl}#t=${pad(h)}h${pad(m)}m${pad(s)}s`;
    if (/sharepoint\.com|stream\.microsoft\.com|microsoftstream|onedrive\.live\.com|my\.sharepoint\.com/i.test(videoUrl)) {
        return `${videoUrl}${sep}startTime=${totalSeconds}`;
    }
    return fragmentCandidate;
}

// processSingleVttFile: main work
// ...existing code...
// ...existing code...
async function processSingleVttFile(context, fileName, outputFormat = 'json') {
    const processingStartTime = Date.now();
    try {
        context.log(`🎬 Starting VTT processing for: ${fileName}`);

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
        const missing = ['tenantId','clientId','clientSecret','openaiKey','openaiEndpoint','sharepointDriveId'].filter(k => !config[k]);
        if (missing.length) throw new Error(`Missing required config: ${missing.join(',')}`);
        context.log('🔧 Loaded configuration (sensitive values redacted)');

        const openaiClient = new OpenAI({ apiKey: config.openaiKey, baseURL: `${config.openaiEndpoint}/openai` });

        const credential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
        const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ['https://graph.microsoft.com/.default'] });
        const graphClient = Client.initWithMiddleware({ authProvider });

        // locate VTT files in drive (root + folders one level)
        let driveItems = await graphClient.api(`/drives/${config.sharepointDriveId}/root/children`).get();
        const vttFiles = [];
        for (const item of driveItems.value || []) {
            if (item.file && item.name && item.name.toLowerCase().endsWith('.vtt')) vttFiles.push(item);
            else if (item.folder) {
                const folderItems = await graphClient.api(`/drives/${config.sharepointDriveId}/items/${item.id}/children`).get().catch(()=>({ value: [] }));
                for (const sub of folderItems.value || []) if (sub.file && sub.name && sub.name.toLowerCase().endsWith('.vtt')) vttFiles.push(sub);
            }
        }

        let targetFile = vttFiles.find(f => f.name.toLowerCase() === fileName.toLowerCase())
            || vttFiles.find(f => f.name.toLowerCase().includes((fileName || '').replace('.vtt','').toLowerCase()));
        if (!targetFile) return { success: false, status: 404, error: `File not found: ${fileName}`, processedAt: new Date().toISOString(), processingTimeMs: Date.now() - processingStartTime };

        // try Graph video details (chapters/webUrl) then fallback to metadata
        let videoChapters = [];
        let finalVideoUrl = '';
        try {
            const videoDetails = await graphClient.api(`/drives/${config.sharepointDriveId}/items/${targetFile.id}/video?$select=chapters,webUrl`).get().catch(()=>null);
            if (videoDetails && videoDetails.webUrl) {
                finalVideoUrl = videoDetails.webUrl;
                videoChapters = (videoDetails.chapters || []).map(c => ({ title: c.title || '', startMs: (typeof c.start === 'number') ? c.start : (Number(c.start) || 0) }));
            } else {
                finalVideoUrl = getVideoUrlFromMetadata(targetFile);
            }
        } catch {
            finalVideoUrl = getVideoUrlFromMetadata(targetFile);
        }

        // fetch VTT content via download URL
        const fileDetails = await graphClient.api(`/drives/${config.sharepointDriveId}/items/${targetFile.id}`).select('@microsoft.graph.downloadUrl,name,size,id').get();
        const downloadUrl = fileDetails && fileDetails['@microsoft.graph.downloadUrl'] ? fileDetails['@microsoft.graph.downloadUrl'] : targetFile['@microsoft.graph.downloadUrl'];
        if (!downloadUrl) throw new Error('Download URL not available for VTT file');
        const httpFetch = globalThis.fetch ? globalThis.fetch.bind(globalThis) : fetch;
        const response = await httpFetch(downloadUrl);
        if (!response.ok) throw new Error(`Failed to download VTT: HTTP ${response.status}`);
        const raw = await response.text();
        const MAX_CHARS = process.env.MAX_VTT_CHARS ? Number(process.env.MAX_VTT_CHARS) : 0;
        const vttContent = (MAX_CHARS > 0) ? raw.slice(0, MAX_CHARS) : raw;

        // parse timestamp blocks
        if (typeof vttUtils.parseVttTimestamps !== 'function') throw new Error('vttUtils.parseVttTimestamps missing');
        const timestampBlocks = vttUtils.parseVttTimestamps(vttContent || '');

        //
        // NEW: fetch listItem fields (SharePoint column values) and prefer viewer URL in those fields
        //
        let fileFields = {};
        try {
            fileFields = await graphClient.api(`/drives/${config.sharepointDriveId}/items/${targetFile.id}/listItem/fields`).get().catch(()=> ({}));
            context.log('Retrieved listItem fields keys:', Object.keys(fileFields || {}));
        } catch (fieldsErr) {
            context.log('No listItem fields available or failed to fetch fields:', fieldsErr?.message || fieldsErr);
            fileFields = {};
        }

        // helper: find a likely "viewer" URL in fields
        function findViewerUrlFromFields(fieldsObj) {
            if (!fieldsObj || typeof fieldsObj !== 'object') return null;
            const vals = Object.values(fieldsObj);
            for (const v of vals) {
                if (!v || typeof v !== 'string') continue;
                const s = v.trim();
                if (!s.startsWith('http')) continue;
                if (s.includes('/:v:') || s.includes('/%3Av%3A') || /my\.sharepoint\.com\/:v:|sharepoint\.com\/:v:/.test(s)) return s;
                if (s.includes('sharepoint.com') && (s.includes('?e=') || s.includes('nav='))) return s;
            }
            for (const v of vals) {
                if (typeof v === 'string' && v.trim().startsWith('http')) return v.trim();
            }
            return null;
        }

        const viewerFieldUrl = findViewerUrlFromFields(fileFields);
        const preferredViewerUrl = viewerFieldUrl || finalVideoUrl || targetFile.webUrl || '';

        // add ?web=1 for SharePoint/OneDrive viewer pages where helpful
        function addWebParamIfSharepoint(url) {
            if (!url) return url;
            try {
                const u = new URL(url);
                if (/sharepoint|my\.sharepoint|onedrive/i.test(u.hostname)) {
                    if (!u.searchParams.has('web')) u.searchParams.set('web', '1');
                    return u.toString();
                }
            } catch { }
            return url;
        }
        const finalViewerUrl = addWebParamIfSharepoint(preferredViewerUrl);

        // generate transcript text for AI
        const transcriptText = (timestampBlocks || []).map(b => `${b.timestamp || ''} ${b.content || b.text || ''}`).join('\n').trim();
        let summary = '';
        let keyPoints = [];

        try {
            const aiResponse = await openaiClient.chat.completions.create({
                model: config.deployment,
                messages: [{ role: 'system', content: 'You output only strict JSON.' }, { role: 'user', content: transcriptText || '' }],
                temperature: 0.2,
                max_tokens: 512
            });
            const aiContent = aiResponse?.choices?.[0]?.message?.content ?? '';
            const parsed = vttUtils.safeParseModelJson(aiContent);
            summary = parsed.summary || vttUtils.generateFallbackSummary(transcriptText);
            keyPoints = Array.isArray(parsed.keyPoints) ? parsed.keyPoints : [];
        } catch (e) {
            summary = vttUtils.generateFallbackSummary(transcriptText);
            keyPoints = vttUtils.deriveKeyPointsFallbackFromText(transcriptText);
        }

        // Enrich keyPoints with timestamps and viewer links where possible
        keyPoints = (keyPoints || []).map((kp, idx) => {
            const ts = kp && kp.timestamp ? kp.timestamp : timestampBlocks[idx]?.timestamp;
            const norm = vttUtils.normalizeTimestampToSeconds(ts || '00:00:00');
            const seconds = Math.floor(norm.totalSeconds || 0);
            const videoLink = finalViewerUrl ? `${finalViewerUrl}${finalViewerUrl.includes('?') ? '&' : '?'}t=${seconds}` : '';
            return {
                title: kp.title || (typeof kp === 'string' ? kp : (timestampBlocks[idx]?.text || timestampBlocks[idx]?.content || '').slice(0, 120)),
                timestamp: ts || timestampBlocks[idx]?.timestamp || '',
                speaker: kp.speaker || timestampBlocks[idx]?.speaker || '',
                videoLink
            };
        });

        if (!summary || summary.trim().length < 10) summary = vttUtils.generateFallbackSummary(transcriptText);
        if (!Array.isArray(keyPoints) || keyPoints.length === 0) {
            keyPoints = vttUtils.deriveKeyPointsFallbackFromText(transcriptText).slice(0,8).map((t, i) => {
                const ts = timestampBlocks[i]?.timestamp || '';
                const norm = vttUtils.normalizeTimestampToSeconds(ts || '00:00:00');
                const seconds = Math.floor(norm.totalSeconds || 0);
                const videoLink = finalViewerUrl ? `${finalViewerUrl}${finalViewerUrl.includes('?') ? '&' : '?'}t=${seconds}` : '';
                return { title: t, timestamp: ts, speaker: timestampBlocks[i]?.speaker || '', videoLink };
            });
        }

        // Build result object (fixes previous undefined result error)
        const result = {
            success: true,
            meetingTitle: (fileFields && (fileFields.Title || fileFields.FileLeafRef)) || targetFile.name || '',
            date: (fileFields && (fileFields.EventDate || fileFields.Modified || targetFile.lastModifiedDateTime)) || new Date().toISOString(),
            videoUrl: finalViewerUrl || '',
            file: fileName,
            actualFile: targetFile.name,
            summary,
            keyPoints,
            timestampBlocks,
            metadata: { processingTimeMs: Date.now() - processingStartTime, originalContentLength: raw.length, fieldsFound: Object.keys(fileFields || {}).length }
        };

        const formatted = await vttUtils.applyOutputFormat(context, result, outputFormat);
        context.log('✅ Output formatted');
        return formatted;
    } catch (error) {
        context.log.error('❌ processSingleVttFile error:', error?.message || error);
        return { success: false, status: 500, error: error?.message || String(error), stack: error?.stack || 'No stack trace', processedAt: new Date().toISOString(), processingTimeMs: Date.now() - processingStartTime };
    }
}
// ...existing code...