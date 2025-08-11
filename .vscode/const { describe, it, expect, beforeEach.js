const { describe, it, expect, beforeEach, afterEach, jest } = require('@jest/globals');
const { ClientSecretCredential } = require('@azure/identity');
const { Client, TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client');
const { OpenAI } = require('openai');

// filepath: src/functions/ProcessVttFile/index.test.js

// Mock Azure Functions context
const mockContext = {
    log: jest.fn(),
    'log.error': jest.fn(),
    invocationId: 'test-invocation-id',
    executionContext: {
        invocationId: 'test-invocation-id',
        functionName: 'ProcessVttFile'
    }
};

// Mock environment variables
const mockEnv = {
    TENANT_ID: 'test-tenant-id',
    CLIENT_ID: 'test-client-id',
    CLIENT_SECRET: 'test-client-secret',
    OPENAI_ENDPOINT: 'https://ai-teams-eastus2.openai.azure.com/',
    OPENAI_KEY: 'test-openai-key',
    OPENAI_DEPLOYMENT: 'gpt-4o-text',
    SHAREPOINT_DRIVE_ID: 'test-drive-id',
    SHAREPOINT_SITE_URL: 'https://test.sharepoint.com/sites/test'
};

// Mock modules
jest.mock('@azure/identity');
jest.mock('@microsoft/microsoft-graph-client');
jest.mock('openai');


// Mock implementations
const mockGraphClient = {
    api: jest.fn().mockReturnThis(),
    get: jest.fn(),
    select: jest.fn().mockReturnThis()
};

const mockOpenAIClient = {
    chat: {
        completions: {
            create: jest.fn()
        }
    }
};

Client.initWithMiddleware = jest.fn().mockReturnValue(mockGraphClient);
OpenAI.mockImplementation(() => mockOpenAIClient);

// Import the functions to test
const { 
    parseVttTimestamps, 
    extractMeetingMetadata, 
    extractKeyPoints, 
    createVideoLink 
} = require('./index.js');

describe('ProcessVttFile Azure Function', () => {
    beforeEach(() => {
        // Reset mocks
        jest.clearAllMocks();
        
        // Set environment variables
        Object.keys(mockEnv).forEach(key => {
            process.env[key] = mockEnv[key];
        });
        
        // Reset context
        mockContext.log.mockClear();
    });

    afterEach(() => {
        // Clean up environment variables
        Object.keys(mockEnv).forEach(key => {
            delete process.env[key];
        });
    });

    describe('parseVttTimestamps', () => {
        it('should parse VTT timestamps correctly', () => {
            const vttContent = `WEBVTT

00:00:05.000 --> 00:00:10.000
<v John Doe>Hello everyone, welcome to today's meeting.

00:00:10.000 --> 00:00:15.000
<v Jane Smith>Thank you John. Let's start with the agenda.

00:00:15.000 --> 00:00:20.000
We'll be discussing Dynamics 365 CRM features.`;

            const result = parseVttTimestamps(vttContent);

            expect(result).toHaveLength(3);
            expect(result[0]).toEqual({
                timestamp: '00:00:05',
                content: 'Hello everyone, welcome to today\'s meeting. ',
                speaker: 'John Doe'
            });
            expect(result[1]).toEqual({
                timestamp: '00:00:10',
                content: 'Thank you John. Let\'s start with the agenda. ',
                speaker: 'Jane Smith'
            });
            expect(result[2]).toEqual({
                timestamp: '00:00:15',
                content: 'We\'ll be discussing Dynamics 365 CRM features. ',
                speaker: null
            });
        });

        it('should handle empty VTT content', () => {
            const result = parseVttTimestamps('');
            expect(result).toEqual([]);
        });

        it('should handle VTT content without speakers', () => {
            const vttContent = `WEBVTT

00:00:05.000 --> 00:00:10.000
This is a meeting transcript without speaker tags.`;

            const result = parseVttTimestamps(vttContent);

            expect(result).toHaveLength(1);
            expect(result[0]).toEqual({
                timestamp: '00:00:05',
                content: 'This is a meeting transcript without speaker tags. ',
                speaker: null
            });
        });
    });

    describe('extractMeetingMetadata', () => {
        it('should extract title from NOTE line', () => {
            const vttContent = `WEBVTT
NOTE Dynamics 365 CRM Training Session - Advanced Features

00:00:05.000 --> 00:00:10.000
Meeting content here.`;

            const fileMetadata = { name: 'test-meeting.vtt' };
            const sharepointSiteUrl = 'https://test.sharepoint.com/sites/test';

            const result = extractMeetingMetadata(vttContent, fileMetadata, sharepointSiteUrl);

            expect(result.title).toBe('Dynamics 365 CRM Training Session - Advanced Features');
            expect(result.videoUrl).toBe('https://test.sharepoint.com/sites/test/Shared%20Documents/test-meeting');
            expect(result.date).toMatch(/^\d{4}-\d{2}-\d{2}$/); // YYYY-MM-DD format
            expect(result.filename).toBe('test-meeting.vtt');
        });

        it('should fallback to filename when no NOTE line exists', () => {
            const vttContent = `WEBVTT

00:00:05.000 --> 00:00:10.000
Meeting content without NOTE line.`;

            const fileMetadata = { name: 'dynamics-crm-training.vtt' };
            const sharepointSiteUrl = 'https://test.sharepoint.com/sites/test';

            const result = extractMeetingMetadata(vttContent, fileMetadata, sharepointSiteUrl);

            expect(result.title).toBe('Dynamics Crm Training');
            expect(result.videoUrl).toBe('https://test.sharepoint.com/sites/test/Shared%20Documents/dynamics-crm-training');
        });

        it('should handle missing SharePoint site URL', () => {
            const vttContent = 'WEBVTT\n\nSome content';
            const fileMetadata = { name: 'test.vtt' };
            const sharepointSiteUrl = null;

            const result = extractMeetingMetadata(vttContent, fileMetadata, sharepointSiteUrl);

            expect(result.videoUrl).toBe('https://yourtenant.sharepoint.com/video-placeholder');
        });
    });

    describe('createVideoLink', () => {
        it('should create proper video link with timestamp', () => {
            const timestamp = '00:15:30';
            const videoUrl = 'https://test.sharepoint.com/video';

            const result = createVideoLink(timestamp, videoUrl);

            expect(result).toBe('https://test.sharepoint.com/video#t=00h15m30s');
        });

        it('should handle single digit values', () => {
            const timestamp = '01:05:03';
            const videoUrl = 'https://test.sharepoint.com/video';

            const result = createVideoLink(timestamp, videoUrl);

            expect(result).toBe('https://test.sharepoint.com/video#t=01h05m03s');
        });
    });

    describe('extractKeyPoints', () => {
        it('should extract key points from summary with timestamps', () => {
            const summary = `### Meeting Summary

**XRM Toolbox Configuration**
How to configure and use XRM Toolbox for data management.

**Environment Access Management**
Setting up proper access controls for development environments.

**Custom Field Creation**
Best practices for creating custom fields in Dynamics 365.`;

            const timestampBlocks = [
                { timestamp: '00:05:30', speaker: 'John Doe', content: 'XRM discussion' },
                { timestamp: '00:10:15', speaker: 'Jane Smith', content: 'Environment talk' },
                { timestamp: '00:15:45', speaker: 'Bob Wilson', content: 'Custom fields' }
            ];

            const videoUrl = 'https://test.sharepoint.com/video';

            const result = extractKeyPoints(summary, timestampBlocks, videoUrl);

            expect(result).toHaveLength(3);
            expect(result[0]).toEqual({
                title: 'XRM Toolbox Configuration',
                description: 'Key discussion point from John Doe',
                timestamp: '00:05:30',
                videoLink: 'https://test.sharepoint.com/video#t=00h05m30s',
                speaker: 'John Doe'
            });
            expect(result[1]).toEqual({
                title: 'Environment Access Management',
                description: 'Key discussion point from Jane Smith',
                timestamp: '00:10:15',
                videoLink: 'https://test.sharepoint.com/video#t=00h10m15s',
                speaker: 'Jane Smith'
            });
        });

        it('should handle summary without markdown formatting', () => {
            const summary = 'Regular text without special formatting';
            const timestampBlocks = [];
            const videoUrl = 'https://test.sharepoint.com/video';

            const result = extractKeyPoints(summary, timestampBlocks, videoUrl);

            expect(result).toEqual([]);
        });

        it('should filter out short titles', () => {
            const summary = `**Hi**
**This is a longer title that should be included**`;

            const timestampBlocks = [
                { timestamp: '00:05:30', speaker: 'John', content: 'content' },
                { timestamp: '00:10:15', speaker: 'Jane', content: 'content' }
            ];

            const videoUrl = 'https://test.sharepoint.com/video';

            const result = extractKeyPoints(summary, timestampBlocks, videoUrl);

            expect(result).toHaveLength(1);
            expect(result[0].title).toBe('This is a longer title that should be included');
        });
    });

    describe('Integration Tests', () => {
        it('should process complete VTT workflow', () => {
            const vttContent = `WEBVTT
NOTE Dynamics 365 CRM Training - XRM Toolbox Session

00:00:05.000 --> 00:00:10.000
<v Trainer>Welcome to today's training on XRM Toolbox.

00:00:10.000 --> 00:00:25.000
<v Trainer>We'll be covering data import and export features.

00:00:25.000 --> 00:00:40.000
<v Student>Can you show us how to connect to the environment?`;

            const fileMetadata = { name: 'xrm-toolbox-training.vtt' };
            const sharepointSiteUrl = 'https://test.sharepoint.com/sites/training';

            // Test timestamp parsing
            const timestamps = parseVttTimestamps(vttContent);
            expect(timestamps).toHaveLength(3);

            // Test metadata extraction
            const metadata = extractMeetingMetadata(vttContent, fileMetadata, sharepointSiteUrl);
            expect(metadata.title).toBe('Dynamics 365 CRM Training - XRM Toolbox Session');

            // Test key points extraction (with mock summary)
            const mockSummary = `**XRM Toolbox Introduction**
**Data Import/Export Features**
**Environment Connection Setup**`;

            const keyPoints = extractKeyPoints(mockSummary, timestamps, metadata.videoUrl);
            expect(keyPoints).toHaveLength(3);
            expect(keyPoints[0].title).toBe('XRM Toolbox Introduction');
            expect(keyPoints[0].speaker).toBe('Trainer');
        });
    });

    describe('Error Handling', () => {
        it('should handle malformed VTT content gracefully', () => {
            const malformedVtt = 'Not a valid VTT file content';
            
            const result = parseVttTimestamps(malformedVtt);
            expect(result).toEqual([]);
        });

        it('should handle null/undefined inputs', () => {
            expect(parseVttTimestamps(null)).toEqual([]);
            expect(parseVttTimestamps(undefined)).toEqual([]);
            
            const metadata = extractMeetingMetadata('', { name: 'test.vtt' }, null);
            expect(metadata.title).toBe('Test');
            expect(metadata.videoUrl).toBe('https://yourtenant.sharepoint.com/video-placeholder');
        });
    });

    describe('Performance Tests', () => {
        it('should handle large VTT files efficiently', () => {
            // Generate large VTT content
            let largeVttContent = 'WEBVTT\n\n';
            for (let i = 0; i < 1000; i++) {
                const hours = Math.floor(i / 3600).toString().padStart(2, '0');
                const minutes = Math.floor((i % 3600) / 60).toString().padStart(2, '0');
                const seconds = (i % 60).toString().padStart(2, '0');
                
                largeVttContent += `${hours}:${minutes}:${seconds}.000 --> ${hours}:${minutes}:${(parseInt(seconds) + 5) % 60}.000\n`;
                largeVttContent += `<v Speaker${i % 10}>This is content block ${i}.\n\n`;
            }

            const startTime = Date.now();
            const result = parseVttTimestamps(largeVttContent);
            const processingTime = Date.now() - startTime;

            expect(result).toHaveLength(1000);
            expect(processingTime).toBeLessThan(1000); // Should complete within 1 second
        });
    });
});

// Export functions for testing (add this to your main index.js)
module.exports = {
    parseVttTimestamps,
    extractMeetingMetadata,
    extractKeyPoints,
    createVideoLink
};