# File Operations Use Cases

This document analyzes real-world file operation scenarios and evaluates whether our MCP tools adequately support them.

## URL Formats Users Provide

Users share files with agents using various URL formats:

### 1. Direct SharePoint URL (Doc.aspx)
```
https://docvault-my.sharepoint.com/:x:/r/personal/evrim_a_first_bet/_layouts/15/Doc.aspx?sourcedoc=%7BEA4D4B0C-757D-4814-998E-5F8DE82C887F%7D&file=Greg%20Test.xlsx&action=default&mobileredirect=true
```
**Key info**: Contains `sourcedoc` GUID in URL-encoded format (`%7B...%7D` = `{...}`)

### 2. Sharing Link
```
https://docvault-my.sharepoint.com/:x:/g/personal/evrim_a_first_bet/IQAMS03qfXUUSJmOX43oLIh_AeTsTYH-h0Fpx8kvQIZQ2pc?e=IW1YEq
```
**Key info**: Base64-encoded share token after the path

### 3. Teams File Link (from chat)
```
https://docvault.sharepoint.com/sites/TeamName/Shared%20Documents/file.xlsx
```

---

## Use Cases & Tool Coverage

### UC1: Download a Shared File

**Scenario**: User shares a file link, agent needs to download and view it.

**Current Tool**: `download_sharepoint_file`
- Accepts: Array of SharePoint/OneDrive URLs
- Returns: Downloaded file path, filename, size

**Status**: FULLY SUPPORTED (tested 2024-12-22)

| URL Type | Supported? | Notes |
|----------|------------|-------|
| Sharing link (`:x:/g/`) | YES | Works via `/shares/{encoded}/driveItem/content` |
| Doc.aspx with sourcedoc | YES | Same encoding works - tested and confirmed |
| Direct file path | ? | Needs testing |

Both sharing links and Doc.aspx URLs work with the same base64 encoding approach.

---

### UC2: Upload a New File

**Scenario**: Agent creates a file and shares it with user.

**Current Tool**: `upload_to_sharepoint`
- Accepts: Local file path
- Returns: Share URL (view-only link)

**Status**: WORKS, BUT LIMITED

**Issues**:
1. Returns only `contentUrl` - doesn't return `itemId` for later operations
2. Creates view-only link (`type: 'view'`) - user can't edit
3. No option to set permissions during upload

**Proposed Changes**:
- Return both `itemId` and `contentUrl`
- Add optional `role` parameter ('view' | 'edit')

---

### UC3: Update Agent's Own File

**Scenario**: Agent previously uploaded a file, now needs to update it.

**Current Tool**: `upload_to_sharepoint` (same name = overwrite)

**Status**: WORKS (with same filename)

But agent loses track of files between sessions. No way to list agent's uploaded files.

**Gap**: Consider adding `list_my_files` tool?

---

### UC4: Update a User's Shared File

**Scenario**: User shares a file with edit permissions, agent modifies and saves back.

**Current Tool**: NONE (needs implementation)

**Required API**: `PUT /drives/{driveId}/items/{itemId}/content`

**Status**: API VERIFIED WORKING (tested 2024-12-22)

**Prerequisites** (now complete):
1. OAuth scope: `Files.ReadWrite.All` - ADDED
2. User must grant edit permissions on the file
3. File must not be locked (not open in browser/Excel)

**Workflow verified**:
```javascript
// 1. Get driveId and itemId from share URL
const driveItem = await client.api(`/shares/${encodedUrl}/driveItem`).get();
const driveId = driveItem.parentReference.driveId;
const itemId = driveItem.id;

// 2. PUT new content
await client.api(`/drives/${driveId}/items/${itemId}/content`).put(fileContent);
```

**Still needed**: New tool `update_sharepoint_file` to expose this to agents

---

### UC5: Create a Copy of Shared File

**Scenario**: User shares a file, agent creates a copy in agent's OneDrive.

**Current Tools**: `download_sharepoint_file` + `upload_to_sharepoint`

**Status**: WORKS

Workflow:
1. Download shared file to local
2. Upload to agent's OneDrive (possibly with new name)
3. Share link with user

---

### UC6: Grant Permissions on Uploaded File

**Scenario**: Agent uploads a file, user needs edit access.

**Current Tool**: NONE (planned)

**Required API**: `POST /me/drive/items/{itemId}/invite`

**Status**: NOT SUPPORTED

**Required**: New tool `grant_file_permission`
- Input: itemId, emails[], role ('read' | 'write'), sendInvitation?
- Grants permissions to specified users

---

### UC7: Get File Metadata Without Downloading

**Scenario**: Agent needs file info (size, modified date, permissions) without downloading.

**Current Tool**: NONE

**Status**: NOT SUPPORTED

**Consider**: `get_file_info` tool
- Input: shareUrl
- Returns: name, size, lastModified, permissions, driveId, itemId

---

## Tool Inventory

### Current Tools

| Tool | Purpose | Gaps |
|------|---------|------|
| `download_sharepoint_file` | Download files from share URLs | May not handle all URL formats |
| `upload_to_sharepoint` | Upload to agent's OneDrive | No itemId return, view-only links |

### Proposed New Tools

| Tool | Purpose | Priority |
|------|---------|----------|
| `update_sharepoint_file` | Update content of existing file | HIGH |
| `grant_file_permission` | Grant read/write access to users | HIGH |
| `get_file_info` | Get metadata without downloading | MEDIUM |
| `list_my_files` | List agent's uploaded files | LOW |

---

## Required OAuth Scopes

~~Current: `Files.ReadWrite` - only agent's own OneDrive~~

Updated: `Files.ReadWrite.All` - all files user can access (including shared)

**Status**: COMPLETED (2024-12-22) - Scope added and approved by IT

---

## Implementation Priority

1. ~~**Add `Files.ReadWrite.All` scope**~~ - DONE
2. ~~**Verify URL format handling**~~ - DONE (both sharing links and Doc.aspx work)
3. **Add `update_sharepoint_file`** - enable UC4 (API verified working)
4. **Enhance `upload_to_sharepoint`** - return itemId for follow-up operations
5. **Add `grant_file_permission`** - enable UC6
6. **Add `get_file_info`** - convenience tool (lower priority)

---

## Open Questions

1. Should `download_sharepoint_file` auto-detect URL format and normalize?
2. Should we combine upload/update into one smart tool?
3. Do we need to handle Teams channel files differently?
4. Should permissions be granted automatically on upload, or kept separate?
