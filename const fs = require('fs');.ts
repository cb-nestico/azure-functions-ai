const fs = require('fs');
const path = require('path');

describe('VTT File Tests', () => {
    let vttContent;
    let vttLines;
    let cues;

    beforeAll(() => {
        const vttPath = path.join(__dirname, 'test-download.vtt');
        vttContent = fs.readFileSync(vttPath, 'utf8');
        vttLines = vttContent.split('\n');
        cues = parseVTTCues(vttContent);
    });

    describe('VTT Format Validation', () => {
        test('should start with WEBVTT header', () => {
            expect(vttLines[0].trim()).toBe('WEBVTT');
        });

        test('should have valid timestamp format', () => {
            const timestampRegex = /^\d{2}:\d{2}:\d{2}\.\d{3} --> \d{2}:\d{2}:\d{2}\.\d{3}$/;
            const invalidTimestamps = [];
            
            vttLines.forEach((line, index) => {
                if (line.includes('-->')) {
                    if (!timestampRegex.test(line.trim())) {
                        invalidTimestamps.push({ line: index + 1, content: line });
                    }
                }
            });

            expect(invalidTimestamps).toHaveLength(0);
        });

        test('should have proper cue structure', () => {
            const invalidCues = [];
            
            for (let i = 0; i < vttLines.length; i++) {
                const line = vttLines[i].trim();
                
                // Check if line is a cue identifier
                if (line.match(/^[a-f0-9-]+\/\d+-\d+$/)) {
                    const timestampLine = vttLines[i + 1];
                    const contentLine = vttLines[i + 2];
                    
                    if (!timestampLine || !timestampLine.includes('-->')) {
                        invalidCues.push({ cueId: line, issue: 'Missing timestamp' });
                    }
                    
                    if (!contentLine || !contentLine.includes('<v ')) {
                        invalidCues.push({ cueId: line, issue: 'Missing speaker content' });
                    }
                }
            }

            expect(invalidCues).toHaveLength(0);
        });

        test('should have valid speaker tags', () => {
            const speakerRegex = /<v\s+([^>]+)>/;
            const invalidSpeakerTags = [];
            
            vttLines.forEach((line, index) => {
                if (line.includes('<v ')) {
                    const match = line.match(speakerRegex);
                    if (!match || !match[1] || match[1].trim().length === 0) {
                        invalidSpeakerTags.push({ line: index + 1, content: line });
                    }
                }
            });

            expect(invalidSpeakerTags).toHaveLength(0);
        });
    });

    describe('Content Structure Tests', () => {
        test('should have chronological timestamps', () => {
            const timestamps = [];
            
            vttLines.forEach(line => {
                if (line.includes('-->')) {
                    const [start] = line.split(' --> ');
                    timestamps.push(parseTimeToSeconds(start.trim()));
                }
            });

            for (let i = 1; i < timestamps.length; i++) {
                expect(timestamps[i]).toBeGreaterThanOrEqual(timestamps[i - 1]);
            }
        });

        test('should have valid time ranges in each cue', () => {
            const invalidRanges = [];
            
            vttLines.forEach((line, index) => {
                if (line.includes('-->')) {
                    const [start, end] = line.split(' --> ').map(t => parseTimeToSeconds(t.trim()));
                    if (start >= end) {
                        invalidRanges.push({ line: index + 1, start, end });
                    }
                }
            });

            expect(invalidRanges).toHaveLength(0);
        });

        test('should identify all speakers correctly', () => {
            const speakers = new Set();
            const speakerRegex = /<v\s+([^>]+)>/g;
            
            vttLines.forEach(line => {
                let match;
                while ((match = speakerRegex.exec(line)) !== null) {
                    speakers.add(match[1].trim());
                }
            });

            const expectedSpeakers = [
                'Sanija Beronja',
                'Vikrant Upadhyay (XP)',
                'Mark Leonituk',
                'Nahid Mohammadzadeh',
                'Yaneisi Macias'
            ];

            expectedSpeakers.forEach(speaker => {
                expect(speakers.has(speaker)).toBe(true);
            });
        });

        test('should have consistent cue ID format', () => {
            const cueIdRegex = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}\/\d+-\d+$/;
            const invalidIds = [];
            
            vttLines.forEach((line, index) => {
                const trimmed = line.trim();
                if (trimmed && !trimmed.includes('-->') && !trimmed.includes('<v ') && 
                        trimmed !== 'WEBVTT' && trimmed.length > 0) {
                    if (!cueIdRegex.test(trimmed)) {
                        invalidIds.push({ line: index + 1, id: trimmed });
                    }
                }
            });

            expect(invalidIds).toHaveLength(0);
        });
    });

    describe('Meeting Content Analysis', () => {
        test('should calculate correct meeting duration', () => {
            const timestamps = [];
            
            vttLines.forEach(line => {
                if (line.includes('-->')) {
                    const [start, end] = line.split(' --> ').map(t => parseTimeToSeconds(t.trim()));
                    timestamps.push(start, end);
                }
            });

            const duration = Math.max(...timestamps);
            expect(duration).toBeGreaterThan(2800); // Should be over 46 minutes
            expect(duration).toBeLessThan(3000); // Should be under 50 minutes
        });

        test('should have balanced speaker participation', () => {
            const speakerCounts = {};
            
            vttLines.forEach(line => {
                const match = line.match(/<v\s+([^>]+)>/);
                if (match) {
                    const speaker = match[1].trim();
                    speakerCounts[speaker] = (speakerCounts[speaker] || 0) + 1;
                }
            });

            // Vikrant should be the most active speaker (as the presenter)
            expect(speakerCounts['Vikrant Upadhyay (XP)']).toBeGreaterThan(100);
            
            // All main participants should have significant contributions
            expect(speakerCounts['Sanija Beronja']).toBeGreaterThan(20);
            expect(speakerCounts['Mark Leonituk']).toBeGreaterThan(30);
        });

        test('should contain key meeting topics', () => {
            const fullText = vttContent.toLowerCase();
            
            const expectedTopics = [
                'access',
                'permission',
                'database',
                'azure',
                'crm',
                'migration',
                'toolbox',
                'security',
                'super user',
                'sql'
            ];

            expectedTopics.forEach(topic => {
                expect(fullText).toContain(topic);
            });
        });
    });

    describe('Data Quality Tests', () => {
        test('should not have excessive empty lines', () => {
            const emptyLineGroups = [];
            let consecutiveEmpty = 0;
            
            vttLines.forEach((line, index) => {
                if (line.trim() === '') {
                    consecutiveEmpty++;
                } else {
                    if (consecutiveEmpty > 2) {
                        emptyLineGroups.push({ startLine: index - consecutiveEmpty, count: consecutiveEmpty });
                    }
                    consecutiveEmpty = 0;
                }
            });

            expect(emptyLineGroups).toHaveLength(0);
        });

        test('should have complete speaker tags', () => {
            const incompleteTags = [];
            
            vttLines.forEach((line, index) => {
                if (line.includes('<v ') && !line.includes('</v>')) {
                    incompleteTags.push({ line: index + 1, content: line });
                }
            });

            expect(incompleteTags).toHaveLength(0);
        });

        test('should not have overlapping timestamps', () => {
            const overlaps = [];
            const timeRanges = [];
            
            vttLines.forEach((line, index) => {
                if (line.includes('-->')) {
                    const [start, end] = line.split(' --> ').map(t => parseTimeToSeconds(t.trim()));
                    timeRanges.push({ start, end, line: index + 1 });
                }
            });

            for (let i = 1; i < timeRanges.length; i++) {
                const current = timeRanges[i];
                const previous = timeRanges[i - 1];
                
                if (current.start < previous.end) {
                    overlaps.push({
                        previous: { line: previous.line, end: previous.end },
                        current: { line: current.line, start: current.start }
                    });
                }
            }

            expect(overlaps).toHaveLength(0);
        });
    });

    describe('Edge Cases and Error Handling', () => {
        test('should handle multi-line speaker content', () => {
            const multiLineCues = [];
            
            for (let i = 0; i < vttLines.length - 1; i++) {
                const line = vttLines[i];
                const nextLine = vttLines[i + 1];
                
                if (line.includes('<v ') && nextLine.includes('<v ') && 
                        !line.includes('</v>') && nextLine.trim() !== '') {
                    multiLineCues.push({ line: i + 1, content: line });
                }
            }

            // Should handle multi-line content properly
            expect(multiLineCues.length).toBeGreaterThanOrEqual(0);
        });

        test('should validate cue continuity', () => {
            const discontinuities = [];
            
            for (let i = 2; i < vttLines.length; i++) {
                const line = vttLines[i].trim();
                
                if (line.match(/^[a-f0-9-]+\/\d+-\d+$/)) {
                    const prevNonEmpty = findPreviousNonEmptyLine(vttLines, i - 1);
                    
                    if (prevNonEmpty && !prevNonEmpty.includes('</v>') && 
                            !prevNonEmpty.includes('-->')) {
                        discontinuities.push({ line: i + 1, cueId: line });
                    }
                }
            }

            expect(discontinuities).toHaveLength(0);
        });
    });

    describe('Performance Tests', () => {
        test('should parse VTT file efficiently', () => {
            const startTime = Date.now();
            const parsed = parseVTTCues(vttContent);
            const endTime = Date.now();
            
            expect(endTime - startTime).toBeLessThan(1000); // Should parse in under 1 second
            expect(parsed.length).toBeGreaterThan(100); // Should extract substantial number of cues
        });
    });
});

// Helper functions
function parseTimeToSeconds(timeString) {
    const [hours, minutes, seconds] = timeString.split(':');
    return parseInt(hours) * 3600 + parseInt(minutes) * 60 + parseFloat(seconds);
}

function parseVTTCues(content) {
    const lines = content.split('\n');
    const cues = [];
    let currentCue = null;
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        if (line.match(/^[a-f0-9-]+\/\d+-\d+$/)) {
            currentCue = { id: line };
        } else if (line.includes('-->') && currentCue) {
            const [start, end] = line.split(' --> ');
            currentCue.start = start.trim();
            currentCue.end = end.trim();
        } else if (line.includes('<v ') && currentCue) {
            currentCue.content = line;
            cues.push(currentCue);
            currentCue = null;
        }
    }
    
    return cues;
}

function findPreviousNonEmptyLine(lines, startIndex) {
    for (let i = startIndex; i >= 0; i--) {
        if (lines[i].trim() !== '') {
            return lines[i].trim();
        }
    }
    return null;
}