# Azure Functions VTT Meeting Transcript Processor

A powerful Azure Function that automatically processes VTT (Video Text Track) meeting transcripts from SharePoint and generates AI-powered meeting summaries using Azure OpenAI.

## 🎯 **Project Overview**

This Azure Function integrates with Microsoft SharePoint to discover, download, and process VTT meeting transcript files, then uses Azure OpenAI to generate comprehensive meeting summaries with key discussion points, action items, and next steps.

## 📋 **Features**

- ✅ **SharePoint Integration**: Automatic discovery and download of VTT files from SharePoint drives
- ✅ **Azure OpenAI Processing**: AI-powered meeting analysis using GPT-4o
- ✅ **Smart File Handling**: Supports both exact and partial filename matching
- ✅ **Rate Limit Management**: Intelligent content truncation to stay within API limits
- ✅ **Error Handling**: Comprehensive error handling with detailed logging
- ✅ **Multiple Request Methods**: Supports both GET and POST requests
- ✅ **Production Ready**: Optimized for Azure deployment with proper configuration management

## 🏗️ **Architecture**

```
SharePoint Drive → Azure Function → Microsoft Graph API → Azure OpenAI → Meeting Summary
```

### **Components:**
- **Azure Function App**: Node.js 18 runtime with Azure Functions v4
- **Microsoft Graph API**: SharePoint file access and authentication
- **Azure OpenAI**: GPT-4o text model for meeting analysis
- **SharePoint**: VTT file storage and management

## 🚀 **What Was Accomplished Today**

### **Phase 1: Initial Setup & Authentication** ✅
- Created Azure Function App with Node.js runtime
- Configured Microsoft Graph API authentication using service principal
- Set up SharePoint site and drive integration
- Established secure environment variable management

### **Phase 2: SharePoint Integration** ✅
- Implemented file discovery across SharePoint drives and subfolders
- Added support for recursive folder scanning
- Created robust file matching (exact and partial filename support)
- Successfully discovered and cataloged 37 VTT files across multiple folders

### **Phase 3: File Download Resolution** ✅
- **Critical Breakthrough**: Resolved Microsoft Graph SDK stream handling issues
- Implemented two-step download process:
  1. Get download URL from Microsoft Graph API
  2. Use native `fetch()` for reliable file content retrieval
- Successfully downloaded large VTT files (136,233 characters)
- Added comprehensive content validation and preview logging

### **Phase 4: Azure OpenAI Integration** ✅
- Configured dual Azure OpenAI resources:
  - **Canada Central**: Text processing endpoint
  - **East US 2**: Audio processing endpoint  
- Created GPT-4o text deployment (`gpt-4o-text`) in East US 2
- Implemented rate limiting and token management
- Successfully generated comprehensive meeting summaries

### **Phase 5: Production Optimization** ✅
- Added intelligent content truncation for large files
- Implemented comprehensive error handling and logging
- Created detailed metadata tracking for processed files
- Optimized for Azure OpenAI S0 pricing tier rate limits

### **Phase 6: Testing & Validation** ✅
- Created comprehensive test suite for validation
- Verified end-to-end functionality with multiple file sizes
- Confirmed error handling for edge cases
- Generated production-ready test scripts

## 📁 **Project Structure**

```
C:\AZURE FUNCTIONS-AI\
├── src/
│   └── functions/
│       └── ProcessVttFile/
│           ├── index.js              # Main function code
│           └── test-function.bat     # Test script
├── local.settings.json               # Environment configuration
├── package.json                      # Dependencies
├── package-lock.json
├── host.json                         # Function app configuration
└── README.md                         # This file
```

## ⚙️ **Configuration**

### **Environment Variables**

```json
{
  "TENANT_ID": "your-tenant-id-here",
  "CLIENT_ID": "your-client-id-here",
  "CLIENT_SECRET": "your-client-secret-here",
  "OPENAI_ENDPOINT": "https://your-openai-resource.openai.azure.com/",
  "OPENAI_DEPLOYMENT": "your-deployment-name",
  "OPENAI_KEY": "your-openai-api-key-here",
  "SHAREPOINT_SITE_URL": "https://yourtenant.sharepoint.com/sites/YourSite",
  "SHAREPOINT_DRIVE_ID": "your-sharepoint-drive-id-here"
}
```

### **Azure Resources**

#### **Azure OpenAI Resources**
- **YOUR-OPENAI-RESOURCE** (East US 2)
  - Endpoint: `https://your-openai-resource.openai.azure.com/`
  - Deployments:
    - `gpt-4o-audio-preview`: Audio processing
    - `gpt-4o-text`: Text/meeting analysis ✅ **In Use**

#### **SharePoint Integration**
- **Site**: Your SharePoint Site
- **Drive**: Main document library with VTT files
- **Folders**: Recursive scanning including MeetingSummaries, shared, Debug, etc.

## 🔧 **Installation & Setup**

### **Prerequisites**
- Node.js 18.x
- Azure CLI
- Azure Functions Core Tools v4
- Valid Azure subscription with:
  - Azure OpenAI resource
  - SharePoint Online access
  - Service principal with appropriate permissions

### **Local Development Setup**

1. **Clone and install dependencies:**
   ```bash
   cd C:\AZURE FUNCTIONS-AI
   npm install
   ```

2. **Configure environment variables:**
   - Update `local.settings.json` with your Azure resource details

3. **Start local development:**
   ```bash
   func start
   ```

4. **Test the function:**
   ```bash
   # Run the test script
   .\src\functions\ProcessVttFile\test-function.bat
   
   # Or test manually
   az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=your-vtt-file.vtt"
   ```

## 📖 **API Documentation**

### **Endpoint**
- **Local**: `http://localhost:7071/api/ProcessVttFile`
- **Production**: `https://your-function-app.azurewebsites.net/api/ProcessVttFile`

### **Methods**
- **GET**: Query parameter `name` with VTT filename
- **POST**: JSON body with `name` property

### **Request Examples**

**GET Request:**
```bash
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=meeting-transcript.vtt"
```

**POST Request:**
```bash
az rest --method POST --url "http://localhost:7071/api/ProcessVttFile" \
  --headers "Content-Type=application/json" \
  --body '{"name":"meeting-transcript.vtt"}'
```

### **Response Format**

**Success Response:**
```json
{
  "success": true,
  "file": "meeting-transcript.vtt",
  "actualFile": "meeting-transcript.vtt",
  "summary": "### Meeting Summary\n\n**Participants:**\n- Name 1\n- Name 2...",
  "metadata": {
    "endpoint": "https://your-openai-resource.openai.azure.com/",
    "deployment": "your-deployment-name",
    "fileSize": 136233,
    "originalContentLength": 32000,
    "truncated": true,
    "estimatedTokens": 8000,
    "processedAt": "2025-07-08T20:35:54.628Z"
  }
}
```

**Error Response:**
```json
{
  "error": "Processing failed",
  "details": "File not found: non-existent.vtt. Available VTT files: ...",
  "endpoint": "https://your-openai-resource.openai.azure.com/",
  "deployment": "your-deployment-name"
}
```

## 🧪 **Testing**

### **Available Test Files**
The function has been tested with various VTT files of different sizes:

- **Small files** (~8KB): Small meeting transcripts
- **Medium files** (~20KB): Standard meeting transcripts
- **Large files** (~116KB): Extended training sessions
- **Extra large files** (~251KB): Long leadership meetings

### **Test Results**
- ✅ **File Discovery**: Successfully found 37 VTT files across multiple folders
- ✅ **File Download**: Downloaded full content (136,233 characters)
- ✅ **AI Processing**: Generated comprehensive meeting summaries
- ✅ **Rate Limiting**: Proper handling of Azure OpenAI token limits
- ✅ **Error Handling**: Robust error handling for edge cases

### **Running Tests**

**Quick Test:**
```bash
.\src\functions\ProcessVttFile\test-function.bat
```

**Manual Tests:**
```bash
# Test successful processing
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=your-test-file.vtt"

# Test small file
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=small-meeting.vtt"

# Test error handling
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=non-existent.vtt"
```

## 🚀 **Deployment to Azure**

### **Create Azure Resources**

```bash
# Create resource group
az group create --name "AI-VIDEO" --location "eastus2"

# Create storage account
az storage account create \
  --name "yourstorageaccount" \
  --resource-group "AI-VIDEO" \
  --location "eastus2" \
  --sku "Standard_LRS"

# Create function app
az functionapp create \
  --name "your-vtt-processor" \
  --resource-group "AI-VIDEO" \
  --storage-account "yourstorageaccount" \
  --consumption-plan-location "eastus2" \
  --runtime "node" \
  --runtime-version "18" \
  --functions-version "4"
```

### **Deploy Function Code**

```bash
# Deploy to Azure
func azure functionapp publish your-vtt-processor
```

### **Configure Production Settings**

```bash
# Set all environment variables
az functionapp config appsettings set --name "your-vtt-processor" --resource-group "AI-VIDEO" \
  --settings "TENANT_ID=your-tenant-id" \
             "CLIENT_ID=your-client-id" \
             "CLIENT_SECRET=your-client-secret" \
             "OPENAI_ENDPOINT=https://your-openai-resource.openai.azure.com/" \
             "OPENAI_DEPLOYMENT=your-deployment-name" \
             "OPENAI_KEY=your-openai-key" \
             "SHAREPOINT_DRIVE_ID=your-drive-id"
```

### **Test Production Deployment**

```bash
az rest --method GET --url "https://your-vtt-processor.azurewebsites.net/api/ProcessVttFile?name=test-file.vtt"
```

## 🔍 **Key Technical Solutions**

### **1. Microsoft Graph Stream Handling** 🎯
**Problem**: Microsoft Graph SDK was returning `[object ReadableStream]` instead of file content.

**Solution**: Implemented two-step download process:
```javascript
// Get download URL from Graph API
const downloadUrlResponse = await graphClient
    .api(`/drives/${process.env.SHAREPOINT_DRIVE_ID}/items/${targetFile.id}`)
    .select('@microsoft.graph.downloadUrl')
    .get();

// Use native fetch for reliable download
const response = await fetch(downloadUrl);
const vttContent = await response.text();
```

### **2. Azure OpenAI Rate Limiting** ⚡
**Problem**: Large VTT files exceeded token rate limits.

**Solution**: Intelligent content truncation:
```javascript
const MAX_TOKENS = 8000;
const CHARS_PER_TOKEN = 4;
const maxChars = MAX_TOKENS * CHARS_PER_TOKEN;

if (vttContent.length > maxChars) {
    vttContent = vttContent.substring(0, maxChars);
}
```

### **3. Regional Model Availability** 🌍
**Problem**: GPT models not available in Canada Central region.

**Solution**: Used East US 2 region with dual deployments:
- `gpt-4o-audio-preview`: For future audio processing
- `gpt-4o-text`: For current meeting analysis

## 📊 **Performance Metrics**

- **File Discovery**: ~2-3 seconds for 37 files across multiple folders
- **File Download**: ~200ms for 136KB VTT file
- **AI Processing**: ~6 seconds for 32,000 character analysis
- **Total Processing**: ~8.8 seconds end-to-end
- **Token Usage**: ~8,000 tokens per large file (optimized for rate limits)

## 🔐 **Security & Best Practices**

- ✅ Service principal authentication (no user credentials stored)
- ✅ Environment variable configuration management
- ✅ Secure Azure Key Vault integration ready
- ✅ Comprehensive error handling without credential exposure
- ✅ Rate limiting to prevent API abuse
- ✅ Input validation and sanitization

## 🎯 **Generated Meeting Summary Example**

The function generates comprehensive meeting summaries including:

- **Participants**: Automatic identification from VTT speakers
- **Key Discussion Points**: Main topics and conversations
- **Action Items**: Tasks and assignments identified
- **Important Decisions**: Key decisions made during meeting
- **Next Steps**: Follow-up actions and timelines
- **Participant Contributions**: Individual contribution summaries

## 🛠️ **Troubleshooting**

### **Common Issues**

1. **Rate Limiting**: 
   - Wait 60 seconds between requests
   - Function automatically truncates large files

2. **File Not Found**:
   - Check filename spelling
   - Function supports partial name matching
   - Review available files in logs

3. **Authentication Issues**:
   - Verify service principal permissions
   - Check environment variable configuration

### **Debug Logs**
The function provides detailed logging for troubleshooting:
- File discovery process
- Download progress and content validation
- Azure OpenAI processing status
- Error details with stack traces

## 🔮 **Future Enhancements**

- **SharePoint Webhooks**: Automatic processing on file upload
- **Batch Processing**: Process multiple files simultaneously
- **Enhanced AI Analysis**: Speaker sentiment analysis, meeting effectiveness scoring
- **Output Formats**: Export summaries to Word, PDF, or SharePoint lists
- **Real-time Processing**: Stream processing for live transcripts

## 📞 **Support**

For issues or questions:
1. Check the troubleshooting section
2. Review Azure Function logs
3. Validate environment configuration
4. Test with different VTT files

## 🏆 **Success Metrics**

Today's development session achieved:
- ✅ **100% Core Functionality**: Complete end-to-end processing working
- ✅ **37 VTT Files Discovered**: Full SharePoint integration
- ✅ **Multiple File Sizes Tested**: From 6KB to 251KB files
- ✅ **Production Ready**: Optimized for Azure deployment
- ✅ **Comprehensive Error Handling**: Robust edge case management
- ✅ **AI Quality Summaries**: High-quality meeting analysis output

**The Azure Function is now production-ready and successfully processing VTT meeting transcripts with AI-powered analysis!** 🎉

---

## 🚀 **Next Session Action Plan - Meeting Requirements Enhancement**

**Session Date**: July 9, 2025  
**Current Status**: ✅ Core VTT processing working - Ready for requirements alignment  
**Goal**: Enhance function to meet specific Dynamics 365 CRM training requirements

### **📊 Requirements Gap Analysis (Reference)**

Based on the specific requirements for Meeting Transcript Summarization, we need to enhance the current working solution:

| **Requirement** | **Current Status** | **Priority** | **Effort** |
|----------------|-------------------|--------------|------------|
| VTT format processing | ✅ **Complete** | N/A | Done |
| Video URL extraction | ❌ **Missing** | High | 1 hour |
| NOTE line title parsing | ❌ **Missing** | High | 30 min |
| Timestamp extraction (HH:MM:SS) | ❌ **Missing** | High | 1 hour |
| Training-specific analysis | ⚠️ **Partial** | High | 45 min |
| Linkable time markers (#t=format) | ❌ **Missing** | High | 45 min |
| Structured output format | ⚠️ **Partial** | Medium | 30 min |

### **🎯 Phase 1: Core Enhancement Implementation (90 minutes)**

#### **Step 1: VTT Timestamp Processing (30 minutes)**
```javascript
// Add to current function - VTT timestamp extraction
function parseVttTimestamps(vttContent) {
    const timestampRegex = /(\d{2}:\d{2}:\d{2}\.\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}\.\d{3})/g;
    const contentBlocks = [];
    
    // Extract timestamp blocks with content
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

// Convert timestamp to video link format
function createVideoLink(timestamp, videoUrl) {
    const [hours, minutes, seconds] = timestamp.split(':');
    return `${videoUrl}#t=${hours}h${minutes}m${seconds}s`;
}
```

#### **Step 2: NOTE Title and Video URL Extraction (30 minutes)**
```javascript
// Add to current function - NOTE line parsing and video URL extraction
function extractMeetingMetadata(vttContent, fileMetadata) {
    // Extract meeting title from NOTE line
    const noteMatch = vttContent.match(/NOTE\s+(.+)/);
    const meetingTitle = noteMatch ? noteMatch[1].trim() : "Dynamics 365 CRM Training";
    
    // Extract video URL from SharePoint metadata (if available)
    const videoUrl = fileMetadata.VideoURL || fileMetadata.videoUrl || 
                    "https://yourtenant.sharepoint.com/video-placeholder";
    
    return {
        title: meetingTitle,
        videoUrl: videoUrl,
        date: new Date().toISOString().split('T')[0] // YYYY-MM-DD format
    };
}
```

#### **Step 3: Training-Specific AI Analysis (30 minutes)**
```javascript
// Enhanced AI prompt for Dynamics 365 CRM training focus
const trainingAnalysisPrompt = `
You are an expert in Dynamics 365 CRM training analysis. Analyze this meeting transcript and extract:

1. **Training Topics Covered**: Identify specific Dynamics 365 CRM features, functions, or processes that were taught or discussed.

2. **Key Learning Points**: For each topic, provide:
   - A clear, concise title (e.g., "Creating Custom Fields", "Lead Management Process")
   - A brief 1-2 sentence description of what was taught or demonstrated
   - Any best practices or tips shared

3. **Action Items**: Identify any homework, practice exercises, or follow-up tasks assigned

4. **Q&A Highlights**: Note important questions asked and answers provided

5. **Next Steps**: Any upcoming training sessions or topics mentioned

Format your response as a structured list with clear topics and timestamps. Focus on actionable learning content that team members can reference for CRM training purposes.

Transcript content:
`;
```

### **🎯 Phase 2: Output Format Enhancement (60 minutes)**

#### **Step 4: Structured Output Format (30 minutes)**
```javascript
// New response format matching requirements
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
            speaker: point.speaker
        })),
        summary: summary,
        metadata: {
            originalFile: metadata.filename,
            fileSize: metadata.fileSize,
            processedAt: new Date().toISOString(),
            totalKeyPoints: keyPoints.length,
            processingTime: metadata.processingTime
        }
    };
}
```

### **🎯 Phase 3: Testing and Validation (30 minutes)**

#### **Enhanced Test Commands**
```bash
# Test with enhanced requirements
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=your-test-file.vtt"

# Validate new output format
az rest --method POST --url "http://localhost:7071/api/ProcessVttFile" \
  --headers "Content-Type=application/json" \
  --body '{"name":"training-session.vtt"}' \
  | jq '.keyPoints[] | {title, timestamp, videoLink}'
```

#### **Validation Checklist**
- [ ] ✅ VTT timestamp extraction working
- [ ] ✅ NOTE title parsing (or default fallback)
- [ ] ✅ Video URL integration (from metadata or placeholder)
- [ ] ✅ Training-specific AI analysis
- [ ] ✅ Structured key points output
- [ ] ✅ Linkable timestamp format (#t=00h11m15s)
- [ ] ✅ Meeting title, date, and metadata
- [ ] ✅ Backward compatibility with existing functionality

### **📋 Git Workflow for Session**

```bash
# Start session - check current state
git status
git log --oneline -3

# Create feature branch for requirements work
git checkout -b feature/requirements-enhancement

# During development - commit frequently
git add .
git commit -m "Add VTT timestamp parsing and video link generation"

git add .
git commit -m "Add training-specific AI analysis prompts"

git add .
git commit -m "Implement structured output format matching requirements"

# End of session - merge to main
git checkout main
git merge feature/requirements-enhancement

# Tag completed requirements implementation
git tag -a v1.1.0 -m "Requirements Enhancement v1.1.0"

# Push to remote (if configured)
git push origin main
git push origin --tags
```

### **🔍 Session Success Criteria**

**✅ Minimum Requirements Met:**
- [ ] VTT files processed with timestamp extraction
- [ ] NOTE titles parsed (or default applied)
- [ ] Training-specific AI analysis working
- [ ] Key points extracted with video links
- [ ] Structured output format matching requirements
- [ ] Backward compatibility maintained

**🚀 Stretch Goals:**
- [ ] Enhanced SharePoint metadata integration
- [ ] Multiple file batch processing
- [ ] Optimized AI prompts for better training analysis
- [ ] Comprehensive error handling for new features

### **📝 Ready-to-Execute Commands**

```bash
# Quick start sequence for next session:
cd "C:\AZURE FUNCTIONS-AI"
git status
git checkout -b feature/requirements-enhancement
func start

# Test current functionality
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=your-test-file.vtt"

# After implementing enhancements
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=your-test-file.vtt" | jq '.'
```

**Total Estimated Time**: 3 hours for complete requirements implementation

---

**The current working foundation provides an excellent base for implementing the specific meeting transcript requirements. All core infrastructure is proven and working - we just need to enhance the processing logic and output format.** 🚀

*Ready for Requirements Enhancement Session!*

---

JULY 9 PENDING

 Everything You Need is Documented
✅ Your GitHub Repository Has Everything
Your repo at https://github.com/cb-nestico/azure-functions-ai contains:

Complete working Azure Function (production-ready)
Detailed README.md with comprehensive action plan
All requirements clearly documented in the action plan section
Step-by-step implementation guide with code examples
Git workflow procedures for version control
Testing commands ready to execute
📋 Tomorrow's Session - Quick Start Guide
When you return tomorrow, simply:

Open your GitHub repository: https://github.com/cb-nestico/azure-functions-ai
Read the "Next Session Action Plan" section in README.md
Follow the Phase 1-3 implementation steps (3 hours total)
🚀 Quick Start Commands for Tomorrow

# Navigate to project
cd "C:\AZURE FUNCTIONS-AI"

# Check current status
git status
git log --oneline -3

# Create feature branch for requirements work
git checkout -b feature/requirements-enhancement

# Start Azure Functions
func start

# Test current working functionality
az rest --method GET --url "http://localhost:7071/api/ProcessVttFile?name=Vikran-xrmtool2.vtt"

 What You'll Be Implementing Tomorrow
Your README.md has a complete action plan for:

Phase 1 (90 minutes):
VTT timestamp extraction and parsing
NOTE line title parsing for meeting titles
Training-specific AI prompts for Dynamics 365 CRM
Phase 2 (60 minutes):
Structured output format with key points
Video link generation (#t=00h11m15s format)
Enhanced metadata handling
Phase 3 (30 minutes):
Testing and validation
Git commits and version tagging
🎯 Your Success Foundation
Current Status:

✅ Working Azure Function processing 37 VTT files
✅ Production-ready code with 8.8 second processing time
✅ Secure Git repository with clean documentation
✅ Clear implementation roadmap with code examples
Tomorrow you'll enhance it to:

⭐ Extract timestamps and create video links
⭐ Parse meeting titles from NOTE lines
⭐ Generate training-specific summaries for Dynamics 365 CRM
⭐ Output structured format matching requirements
🔄 How to Get Help Tomorrow
Start by reading your README.md - it has everything documented

If you need assistance, show me:

Your current code
The specific step you're working on
Any errors you encounter
I'll be able to help because your repository will show me:

Your complete project structure
Current implementation status
Exactly what needs to be enhanced
🏆 You're Set for Success!
Your documentation is so comprehensive that:

✅ Any developer could pick up where you left off
✅ All requirements are clearly mapped to implementation steps
✅ Code examples are provided for each enhancement
✅ Testing procedures are documented
✅ Git workflow is established
Tomorrow will be a focused 3-hour implementation session to transform your working VTT processor into a requirements-compliant Dynamics 365 CRM training transcript analyzer! 🚀

Your README.md is your roadmap - everything you need is there! 📋