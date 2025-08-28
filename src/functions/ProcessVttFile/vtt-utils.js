// ...existing code...
const { Document, Packer, Paragraph, TextRun } = require('docx');

/**
 * Format and return output in requested format.
 * Supports: json, html, markdown, summary, word (.docx)
 */
async function applyOutputFormat(context, result, outputFormat = 'json') {
    const fmt = (outputFormat || 'json').toLowerCase();

    if (fmt === 'html') {
        const html = (result.htmlContent) || (result.summary ? `<h1>${result.meetingTitle}</h1><p>${result.summary}</p>` : JSON.stringify(result, null, 2));
        return {
            status: 200,
            headers: { 'Content-Type': 'text/html' },
            body: html
        };
    } else if (fmt === 'markdown') {
        const md = `# ${result.meetingTitle}\n\n${result.summary}\n\n` + (result.keyPoints || []).map(k => `- ${k.title || k}`).join('\n');
        return { status: 200, headers: { 'Content-Type': 'text/markdown' }, body: md };
    } else if (fmt === 'word') {
        try {
            // Defensive creator fallback
            let creatorName = 'Transcript Processor';
            try {
                if (result && typeof result === 'object') {
                    const cand =
                        (result.file && (result.file.creator?.name || result.file.creator?.displayName)) ||
                        (result.actualFile && result.actualFile.creator?.name) ||
                        (result.metadata && result.metadata.creator) ||
                        (result.createdBy && (result.createdBy.user?.displayName || result.createdBy.user?.email));
                    if (cand) creatorName = String(cand);
                }
            } catch { /* ignore */ }

            const children = [];

            // Title line - "Meeting Analysis: <Title>"
            const titlePart = String(result.meetingTitle || result.actualFile || result.file || 'Meeting').replace(/\s+/g, ' ').trim();
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: `Meeting Analysis: ${titlePart}`, bold: true, size: 56 })
                ]
            }));

            // Generated / Processing header lines
            const generatedDate = (() => {
                try {
                    const d = result.date ? new Date(result.date) : new Date();
                    const yyyy = d.getFullYear();
                    const mm = String(d.getMonth() + 1).padStart(2, '0');
                    const dd = String(d.getDate()).padStart(2, '0');
                    let t = d.toLocaleTimeString('en-US', { hour12: true });
                    t = t.replace(/\s?AM$/, ' a.m.').replace(/\s?PM$/, ' p.m.');
                    return `${yyyy}-${mm}-${dd} at ${t}`;
                } catch {
                    return String(result.date || new Date().toISOString());
                }
            })();

            const procMs = result.metadata && typeof result.metadata.processingTimeMs === 'number' ? result.metadata.processingTimeMs : (result.metadata && result.metadata.processingTime ? result.metadata.processingTime : 0);
            const kpCount = Array.isArray(result.keyPoints) ? result.keyPoints.length : 0;

            children.push(new Paragraph({ children: [ new TextRun({ text: `Generated: ${generatedDate}`, italics: true, size: 22 }) ] }));
            children.push(new Paragraph({ children: [ new TextRun({ text: `Processing Time: ${procMs}ms | Key Points: ${kpCount}`, italics: true, size: 22 }) ] }));
            children.push(new Paragraph({})); // spacer

            // Executive Summary
            if (result.summary) {
                children.push(new Paragraph({ children: [ new TextRun({ text: 'Executive Summary', bold: true, size: 28 }) ] }));
                const summaryLines = String(result.summary || '').split(/\n/).map(s => s.trim()).filter(Boolean);
                for (const line of summaryLines) children.push(new Paragraph({ children: [ new TextRun({ text: line, size: 24 }) ] }));
                children.push(new Paragraph({})); // spacer
            }

            // Key Discussion Points header
            const kpList = Array.isArray(result.keyPoints) ? result.keyPoints : [];
            children.push(new Paragraph({ children: [ new TextRun({ text: `Key Discussion Points (${kpList.length || 0} items)`, bold: true, size: 28 }) ] }));
            children.push(new Paragraph({})); // small spacer

            // helper: sanitize and shorten titles
            function extractShortTitle(raw) {
                if (!raw) return 'Point';
                let s = String(raw).trim();
                // remove leading timestamps like "00:00:03.719" or "00:00:03"
                s = s.replace(/^[\s]*\d{1,2}:\d{2}:\d{2}(?:[.,]\d+)?\s*/,'');
                // remove leading mm:ss pattern
                s = s.replace(/^[\s]*\d{1,2}:\d{2}(?:[.,]\d+)?\s*/,'');
                // remove inline speaker "Name: "
                s = s.replace(/^[A-Za-z.\- ]{1,60}:\s*/,'');
                // collapse whitespace
                s = s.replace(/\s+/g,' ').trim();
                // pick first sentence or 8 words
                const sent = s.split(/(?<=[.!?])\s+/)[0];
                const words = sent.split(/\s+/).filter(Boolean);
                if (words.length <= 8) return sent;
                return words.slice(0,8).join(' ') + '...';
            }

            // Use a safe viewer base URL (result.videoUrl) if available
            const baseViewer = (result.videoUrl || '').toString().trim();

            // For each key point: numbered line with timestamp, short title, and [Video Link]
            for (let i = 0; i < kpList.length; i++) {
                const kp = kpList[i] || {};
                const idx = i + 1;
                const rawTs = kp.timestamp || kp.time || kp.start || kp.startSeconds || '';
                const ts = (typeof rawTs === 'number') ? formatSecondsAsHms(rawTs) : (String(rawTs || '').split(/[.,]/)[0] || (kp.startSeconds ? formatSecondsAsHms(kp.startSeconds) : '00:00:00'));
                const candidateTitle = kp.title || kp.name || kp.summary || kp.text || '';
                let title = extractShortTitle(candidateTitle);
                if (!title || title === 'Point') {
                    // fallback to nearest timestamp block content (if provided in result.timestampBlocks)
                    if (Array.isArray(result.timestampBlocks) && result.timestampBlocks[i] && (result.timestampBlocks[i].text || result.timestampBlocks[i].content)) {
                        title = extractShortTitle(result.timestampBlocks[i].text || result.timestampBlocks[i].content);
                    }
                }

                // determine viewer url for this kp: prefer kp.videoLink/kp.videoUrl then result.videoUrl
                const baseForThis = (kp.videoLink || kp.videoUrl || baseViewer || '').toString().trim();
                const seconds = (typeof rawTs === 'number') ? Math.floor(rawTs) : (() => {
                    try {
                        const parts = String(rawTs).split(/[.,]/)[0].split(':').map(n => Number(n) || 0);
                        if (parts.length === 1) return parts[0];
                        if (parts.length === 2) return parts[0] * 60 + parts[1];
                        return parts[0] * 3600 + parts[1] * 60 + parts[2];
                    } catch { return 0; }
                })();

                const paraChildren = [
                    new TextRun({ text: `${idx}. `, size: 24 }),
                    new TextRun({ text: `${ts} `, size: 24 }),
                    new TextRun({ text: `${title} `, size: 24 })
                ];

                // Append [Video Link] marker (no ExternalHyperlink used to avoid docx API incompatibilities)
                if (baseForThis) {
                    // create viewer URL with time parameter
                    const viewerUrl = baseForThis.includes('?') ? `${baseForThis}&t=${seconds}` : `${baseForThis}?t=${seconds}`;
                    paraChildren.push(new TextRun({ text: '[Video Link]', color: '0000FF', underline: {}, size: 24 }));
                    // append short URL in small font (keeps clickable URL visible if needed)
                    paraChildren.push(new TextRun({ text: ` ${viewerUrl}`, size: 18 }));
                } else {
                    paraChildren.push(new TextRun({ text: '[Video Link]', color: '0000FF', underline: {}, size: 24 }));
                }

                children.push(new Paragraph({ children: paraChildren }));
            }

            children.push(new Paragraph({})); // spacer

            // Processing Information
            children.push(new Paragraph({ children: [ new TextRun({ text: 'Processing Information', bold: true, size: 26 }) ] }));
            const fileSize = (result.metadata && result.metadata.fileSize) ? result.metadata.fileSize : (result.fileSize || 'N/A');
            const timestampsCount = Array.isArray(result.timestampBlocks) ? result.timestampBlocks.length : 0;
            const tokenInfo = (result.metadata && result.metadata.tokens) ? result.metadata.tokens : (result.metadata && result.metadata.tokenDetails ? result.metadata.tokenDetails : null);
            const tokensLine = tokenInfo ? `Tokens: prompt ${tokenInfo.prompt || ''}, completion ${tokenInfo.completion || ''}, total ${tokenInfo.total || ''}` : '';
            const processingLine = `File Size: ${fileSize} | Timestamps: ${timestampsCount} | Processing Time: ${procMs}ms`;
            children.push(new Paragraph({ children: [ new TextRun({ text: processingLine, size: 22 }) ] }));
            if (tokensLine) children.push(new Paragraph({ children: [ new TextRun({ text: tokensLine, size: 22 }) ] }));

            children.push(new Paragraph({})); // spacer

            // Footer
            const processedAt = result.metadata && result.metadata.processedAt ? result.metadata.processedAt : (new Date()).toISOString();
            children.push(new Paragraph({ children: [ new TextRun({ text: `File: ${result.file || ''} | Processed: ${processedAt}`, size: 20 }) ] }));
            children.push(new Paragraph({ children: [ new TextRun({ text: 'Generated by Azure Functions VTT Meeting Transcript Processor', italics: true, size: 18 }) ] }));

            const doc = new Document({
                creator: creatorName,
                title: result.meetingTitle || 'Meeting Transcript',
                description: result.summary || '',
                sections: [{ children }]
            });

            const buffer = await Packer.toBuffer(doc);
            const safeName = String(result.meetingTitle || (result.file || 'Meeting')).replace(/[^a-z0-9_.-]/gi, '_').slice(0, 200) || 'Meeting';
            const filename = `${safeName}.docx`;

            return {
                status: 200,
                headers: {
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'Content-Disposition': `attachment; filename="${filename}"`,
                    'Content-Length': String(buffer.length)
                },
                isRaw: true,
                body: buffer
            };
        } catch (err) {
            if (context && typeof context.log === 'function') context.log('Error generating .docx:', err);
            return {
                status: 500,
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ success: false, error: 'Failed to generate Word document', details: err && err.message ? err.message : String(err) })
            };
        }
    } else if (fmt === 'summary') {
        return {
            status: 200,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ meetingTitle: result.meetingTitle, summary: result.summary, keyPoints: result.keyPoints }, null, 2)
        };
    }

    // default JSON
    return {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(result, null, 2)
    };
}

// ...existing helper functions unchanged...
function normalizeTimestampToSeconds(ts) {
    if (ts === undefined || ts === null) return { h: 0, m: 0, s: 0, totalSeconds: 0 };
    const noMs = String(ts).split(/[.,]/)[0].trim();
    const parts = noMs.split(':').map(p => Number(p) || 0);
    let h = 0, m = 0, s = 0;
    if (parts.length === 1) { s = parts[0]; }
    else if (parts.length === 2) { m = parts[0]; s = parts[1]; }
    else { h = parts[0]; m = parts[1]; s = parts[2]; }
    const totalSeconds = h * 3600 + m * 60 + s;
    return { h, m, s, totalSeconds };
}

function formatSecondsAsHms(sec) {
    const s = Math.max(0, Math.floor(Number(sec) || 0));
    const h = Math.floor(s / 3600);
    const m = Math.floor((s % 3600) / 60);
    const ss = s % 60;
    return [h, m, ss].map(n => String(n).padStart(2, '0')).join(':');
}

function formatTimestampForUrl(milliseconds) {
    if (typeof milliseconds !== 'number' || isNaN(milliseconds)) return '0m0s';
    const totalSeconds = Math.floor(milliseconds / 1000);
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;
    return `${minutes}m${seconds}s`;
}

/**
 * Minimal meeting metadata extraction from VTT and file info.
 */
function extractMeetingMetadata(vttContent = '', targetFile = {}, siteUrl = '', fileFields = {}) {
    const title = (fileFields.Title || targetFile.name || 'Meeting');
    const date = fileFields.EventDate || targetFile.createdDateTime || new Date().toISOString();
    const videoUrl = fileFields.VideoURL || fileFields.videoUrl || fileFields.URL || targetFile.webUrl || (siteUrl ? siteUrl : '');
    return { title, date, videoUrl };
}

/**
 * Derive simple keypoints by splitting transcript text into sentences and returning phrases.
 */
function deriveKeyPointsFallbackFromText(text = '') {
    if (!text) return [];
    const sentences = text.split(/(?<=[.!?])\s+/).map(s => s.trim()).filter(Boolean);
    const picks = [];
    for (const s of sentences) {
        const trimmed = s.replace(/\s+/g, ' ');
        if (trimmed.length > 20) picks.push(trimmed);
        if (picks.length >= 12) break;
    }
    if (picks.length === 0) {
        return text.split('\n').map(l => l.trim()).filter(Boolean).slice(0, 8);
    }
    return picks;
}

/**
 * Safely parse model JSON that might include stray text.
 */
function safeParseModelJson(input) {
    if (!input || typeof input !== 'string') return {};
    const trimmed = input.trim();
    try {
        return JSON.parse(trimmed);
    } catch {
        const m = trimmed.match(/\{[\s\S]*\}/);
        if (m) {
            try {
                return JSON.parse(m[0]);
            } catch {
                return {};
            }
        }
    }
    return {};
}

/**
 * Simple fallback summary (first 1-2 sentences)
 */
function generateFallbackSummary(text = '') {
    if (!text) return '';
    const sentences = text.split(/(?<=[.!?])\s+/).map(s => s.trim()).filter(Boolean);
    if (sentences.length === 0) return text.slice(0, 300);
    return sentences.slice(0, 2).join(' ');
}

/**
 * chunkArray helper
 */
function chunkArray(arr = [], size = 1) {
    if (!Array.isArray(arr) || size < 1) return [];
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
}

/**
 * Parse VTT cues and extract timestamps + speaker + text.
 */
function parseVttTimestamps(vttContent = '') {
    const text = String(vttContent || '').replace(/\r/g, '');
    const lines = text.split('\n');
    const cues = [];
    let i = 0;

    const tsRegex = /(\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d{1,3})?)\s*-->\s*(\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d{1,3})?)/;

    while (i < lines.length) {
        const line = lines[i].trim();
        if (!line) { i++; continue; }

        // skip numeric cue indexes
        if (/^\d+$/.test(line)) { i++; continue; }

        const m = line.match(tsRegex);
        if (m) {
            const start = m[1].replace(',', '.');
            const end = m[2].replace(',', '.');
            i++;
            const textLines = [];
            while (i < lines.length && lines[i].trim() !== '') {
                textLines.push(lines[i]);
                i++;
            }
            const rawText = textLines.join('\n').trim();

            // extract speaker tag like <v Name>...</v>
            let speaker = null;
            let cleaned = rawText.replace(/<v\s+([^>]+)>/gi, (_match, p1) => {
                speaker = (p1 || speaker || '').trim();
                return '';
            }).replace(/<\/v>/gi, '').trim();

            // fallback: try inline speaker pattern "Name: text"
            if (!speaker) {
                const inline = cleaned.match(/^([^:]{1,60}):\s*(.+)$/s);
                if (inline) {
                    speaker = inline[1].trim();
                    cleaned = (inline[2] || '').trim();
                }
            }

            const startSeconds = normalizeTimestampToSeconds(start).totalSeconds;
            const endSeconds = normalizeTimestampToSeconds(end).totalSeconds;

            cues.push({
                start,
                end,
                startSeconds,
                endSeconds,
                timestamp: start,
                text: cleaned,
                content: cleaned,
                speaker
            });
            continue;
        }

        i++;
    }

    return cues;
}

module.exports = {
    parseVttTimestamps,
    extractMeetingMetadata,
    deriveKeyPointsFallbackFromText,
    chunkArray,
    safeParseModelJson,
    generateFallbackSummary,
    applyOutputFormat,
    normalizeTimestampToSeconds,
    formatTimestampForUrl,
    formatSecondsAsHms
};
// ...existing code...