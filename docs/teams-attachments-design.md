# Teams Attachments MCP Tools - Design Document

## Overview

This document outlines the design for MCP tools to handle Microsoft Teams message attachments (download and upload).

## Implementation Status

### SharePoint Tools - COMPLETED (2025-12-19)

| Tool | Status | Notes |
|------|--------|-------|
| `download_sharepoint_file` | Done | Downloads files from SharePoint URLs |
| `upload_to_sharepoint` | Done | Uploads + creates sharing link |
| `send_chat_message` attachments | Done | Simple string array of URLs |

**Note:** Images can use the same SharePoint workflow - no separate tools needed!

### Hosted Content Tools - COMPLETED (2025-12-19)

| Tool | Status | Notes |
|------|--------|-------|
| `download_hosted_content` | Done | Downloads inline images from message body |

**Note:** For sending images, use the SharePoint workflow (upload + attach). Hosted content download is for receiving inline images from other users.

---

## Two Attachment Systems

Teams uses two completely different storage systems for attachments:

### 1. Hosted Content (Inline Images)

- **Storage**: Embedded directly in Teams messages
- **Use case**: Screenshots, inline images (<4MB)
- **Download API**: `GET /chats/{chatId}/messages/{messageId}/hostedContents/{hostedContentId}/$value`
- **Upload**: Embedded in message POST request (no separate upload step)

### 2. SharePoint Files

- **Storage**: User's OneDrive/SharePoint
- **Use case**: Documents, large files
- **Download API**: `GET /shares/{u!base64url(contentUrl)}/driveItem/content`
- **Upload**: Upload to OneDrive, create sharing link, reference in message

---

## Implemented Tools

### 1. download_sharepoint_file

Download a file attached via SharePoint.

**Parameters:**
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| contentUrl | string | Yes | SharePoint URL from attachment.contentUrl |
| targetDir | string | No | Directory to save (defaults to cwd) |

**Workflow:**
1. Get driveItem metadata: `GET /shares/{encoded}/driveItem` (for canonical filename)
2. Download content: `GET /shares/{encoded}/driveItem/content`

**Returns:**
```json
{
  "path": "./document.pdf",
  "filename": "document.pdf",
  "size": 7228
}
```

**Note:** Filename comes from Graph API metadata (`driveItem.name`), not URL parsing.

---

### 2. upload_to_sharepoint

Upload a file to OneDrive and create a sharing link.

**Parameters:**
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| filePath | string | Yes | Local file path to upload |

**Workflow:**
1. Upload file: `PUT /me/drive/root:/{filename}:/content`
2. Create sharing link: `POST /me/drive/items/{id}/createLink` with `scope: 'organization'`

**Returns:** Just the sharing link URL as a string.
```
https://docvault-my.sharepoint.com/:t:/g/personal/user/EncodedFileId
```

**Key Discovery:** Direct `webUrl` (e.g., `/personal/user/Documents/file.txt`) doesn't work - users get "no access" error. Must use sharing links created via `createLink` API.

---

### 3. send_chat_message with attachments

Extended to support file attachments.

**New Parameter:**
```typescript
attachments?: string[]  // Array of SharePoint sharing links
```

**Example:**
```typescript
await send_chat_message({
  chatId: '19:...',
  message: 'Here are the files',
  attachments: [
    'https://company.sharepoint.com/:t:/g/personal/user/ABC123',
    'https://company.sharepoint.com/:t:/g/personal/user/DEF456'
  ]
});
```

**Key Discovery:** No `<a>` tag or `<attachment>` marker needed in message body. Just pass the URLs in the `attachments` array - Teams renders the attachment card automatically.

**Internal Implementation:**
```typescript
attachments = urls.map((contentUrl) => ({
  id: randomUUID(),
  contentType: 'reference',
  contentUrl,
}));
```

Uses proper Graph types: `ChatMessageAttachment`, `ItemBody`, `Partial<ChatMessage>`.

---

## Complete Workflow Example

```typescript
// 1. Upload a file
const sharingUrl = await upload_to_sharepoint({
  filePath: '/path/to/report.pdf'
});
// Returns: "https://company.sharepoint.com/:t:/g/personal/user/ABC123"

// 2. Send message with attachment
await send_chat_message({
  chatId: '19:abc...@thread.v2',
  message: 'Here is the report',
  attachments: [sharingUrl]
});
```

---

## Images via SharePoint (Recommended)

Images can use the exact same SharePoint workflow as files. Teams displays a nice preview card.

**Example:**
```typescript
// Upload image
const url = await upload_to_sharepoint({ filePath: '/path/to/screenshot.png' });
// Returns: "https://company.sharepoint.com/:i:/g/personal/user/ABC123"
// Note: Images use /:i:/g/ prefix, files use /:t:/g/

// Send with attachment
await send_chat_message({
  chatId: '...',
  message: 'Here is the screenshot',
  attachments: [url]
});
```

**Advantages:**
- Same workflow for files and images
- No base64 encoding needed
- Teams shows preview card with thumbnail
- Simpler for agents to use

---

## Alternative: Hosted Content (Inline Images)

For embedding images directly inside message text (not as attachments), use hosted contents.

### Sending inline images

Base64 data is embedded in the message POST request.

**API Format:**
```typescript
await client.api('/chats/{chatId}/messages').post({
  body: {
    contentType: 'html',
    content: '<p>Text before</p><img src="../hostedContents/1/$value" /><p>Text after</p>'
  },
  hostedContents: [{
    '@microsoft.graph.temporaryId': '1',
    contentBytes: 'base64EncodedImageData...',
    contentType: 'image/png'
  }]
});
```

**Key points:**
- `@microsoft.graph.temporaryId` matches the number in `../hostedContents/1/$value`
- `contentBytes` is base64-encoded image data
- `contentType` is the MIME type (image/png, image/jpeg, etc.)
- Size limit: ~4MB

### 4. download_hosted_content

Download inline images embedded in message body.

**Parameters:**
| Param | Type | Required | Description |
|-------|------|----------|-------------|
| chatId | string | Yes | Chat ID |
| messageId | string | Yes | Message ID containing the image |
| hostedContentId | string | Yes | ID from `<img src>` URL (between `/hostedContents/` and `/$value`) |
| targetDir | string | No | Directory to save file (defaults to cwd) |

**Workflow:**
1. Download content: `GET /chats/{chatId}/messages/{messageId}/hostedContents/{hostedContentId}/$value`
2. Detect file type from magic bytes (PNG, JPEG, GIF)
3. Save to file with appropriate extension

**Returns:**
```json
{
  "path": "./image-abc123def456.png",
  "size": 15234
}
```

**Note:** Filename is generated from first 12 chars of hostedContentId + detected extension.

---

### When to use Hosted Content vs SharePoint

| Use Case | Approach |
|----------|----------|
| Sharing files/images as attachments | SharePoint (recommended) |
| Inline images within message text | Hosted Content |
| Large files (>4MB) | SharePoint |
| Screenshots, diagrams as attachments | SharePoint |

---

## Implementation Files

### Created
- `src/utils/sharepoint.ts` - SharePoint URL encoding + filename extraction
- `src/tools/attachments/download_sharepoint_file.ts`
- `src/tools/attachments/upload_to_sharepoint.ts`
- `src/tools/attachments/index.ts`

### Modified
- `src/client/graph.types.ts` - Added SharePoint types
- `src/client/graph.client.ts` - Added downloadSharePointFile, uploadToSharePoint, attachment handling
- `src/tools/messages/send_chat_message.ts` - Added attachments parameter (string array)
- `src/tools/index.ts` - Registered new tools

---

## Key Learnings

1. **Sharing links required**: Direct OneDrive `webUrl` paths don't grant access. Must use `createLink` API to generate sharing links.

2. **No HTML markers needed**: Initially thought `<a>` tags and `<attachment>` markers were required in message body. Testing proved only the `attachments` array is needed.

3. **Minimal API surface**: Final implementation uses minimal parameters:
   - Upload: just `filePath`
   - Send: just `attachments: string[]`

4. **Name field optional**: The `name` field in attachments is optional - Teams extracts it from the SharePoint file metadata.

5. **Images work via SharePoint**: Same workflow for files and images. No need for hosted content complexity. Sharing links use `/:i:/g/` for images, `/:t:/g/` for files.

---

## Agent Integration Guide

This section explains how an AI agent can determine which download tool to use based on the Teams message object structure.

### Detecting Attachment Types

Teams messages can contain two types of downloadable content:

#### 1. SharePoint Attachments (Files/Images)

**Detection**: Check the `attachments` array for items with `contentType: "reference"`

```json
{
  "attachments": [
    {
      "id": "abc123",
      "contentType": "reference",
      "contentUrl": "https://company.sharepoint.com/:t:/g/personal/user/EncodedId",
      "name": "document.pdf"
    }
  ]
}
```

**Tool to use**: `download_sharepoint_file`
**Parameters to extract**:
- `contentUrl` → from `attachment.contentUrl`

#### 2. Hosted Content (Inline Images)

**Detection**: Parse `body.content` for `<img>` tags containing `/hostedContents/`

```json
{
  "body": {
    "contentType": "html",
    "content": "<p>Check this:</p><img src=\"https://graph.microsoft.com/v1.0/chats/19:abc@thread.v2/messages/123/hostedContents/aWQ9eC1.../$value\">"
  }
}
```

**Tool to use**: `download_hosted_content`
**Parameters to extract**:
- `chatId` → from message context or URL path
- `messageId` → from message `id` field or URL path
- `hostedContentId` → the string between `/hostedContents/` and `/$value` in the img src

### Example Detection Logic

```typescript
function detectDownloadableContent(message: ChatMessage, chatId: string) {
  const downloads = [];

  // Check for SharePoint attachments
  if (message.attachments) {
    for (const att of message.attachments) {
      if (att.contentType === 'reference' && att.contentUrl) {
        downloads.push({
          type: 'sharepoint',
          tool: 'download_sharepoint_file',
          params: { contentUrl: att.contentUrl },
          displayName: att.name || 'Unknown file'
        });
      }
    }
  }

  // Check for inline images in body
  if (message.body?.content) {
    const imgRegex = /hostedContents\/([^/]+)\/\$value/g;
    let match;
    while ((match = imgRegex.exec(message.body.content)) !== null) {
      downloads.push({
        type: 'hosted',
        tool: 'download_hosted_content',
        params: {
          chatId,
          messageId: message.id,
          hostedContentId: match[1]
        },
        displayName: 'Inline image'
      });
    }
  }

  return downloads;
}
```

### Prompt Template for Agents

When presenting downloadable content to an agent, include:

```
This message contains attachments:

**SharePoint Files** (use download_sharepoint_file):
- document.pdf: contentUrl="https://company.sharepoint.com/:t:/g/..."

**Inline Images** (use download_hosted_content):
- Image 1: chatId="19:abc@thread.v2", messageId="123", hostedContentId="aWQ9eC1..."

To download any file, call the appropriate tool with the parameters shown.
```

### Sending Attachments

To send files/images with a message:

1. Upload the file: `upload_to_sharepoint({ filePath: "/path/to/file" })`
2. Get the sharing URL from the response
3. Send message with attachment: `send_chat_message({ chatId, message, attachments: [sharingUrl] })`

Both files and images use the same SharePoint workflow.

---

## Future: Worker-teams Integration

After MCP tools are ready, worker-teams needs updates:
1. Extract attachment metadata from messages (already done)
2. Pass metadata to Claude instead of pre-downloading
3. Update prompts to show download params
4. Remove auto-download code
