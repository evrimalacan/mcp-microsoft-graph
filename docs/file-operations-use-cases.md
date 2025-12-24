# File Operations Use Cases

This document analyzes real-world file operation scenarios and our MCP tool coverage.

**Last Updated**: 2025-12-23 10:09

---

## Architecture: ID-Based Operations

### Why IDs over URLs?
- **Stability**: Microsoft may change URL formats; IDs are stable
- **Simplicity**: IDs are shorter and less error-prone than long URLs
- **Consistency**: All tools use the same identifier format

### The Pattern
```
User shares URL → download_sharepoint_file → { itemId, driveId, path }
                                                    ↓
                        Use IDs for all subsequent operations
```

**Download is the only URL entry point.** All other tools use `itemId` + `driveId`.

---

## URL Formats (Input to Download)

All formats work with `download_sharepoint_file`:

### 1. Browser URL (Doc.aspx)
```
https://docvault-my.sharepoint.com/:x:/r/personal/user/_layouts/15/Doc.aspx?sourcedoc=%7BGUID%7D&file=Name.xlsx
```

### 2. Sharing Link
```
https://docvault-my.sharepoint.com/:x:/g/personal/user/BASE64TOKEN?e=CODE
```

### 3. Teams File Link
```
https://docvault.sharepoint.com/sites/TeamName/Shared%20Documents/file.xlsx
```

Internally, all URLs are resolved via `/shares/{encoded}/driveItem` to get stable IDs.

---

## Tool Inventory

| Tool | Input | Output | Purpose |
|------|-------|--------|---------|
| `download_sharepoint_file` | URL | path, **itemId, driveId** | Download + get IDs |
| `upload_to_sharepoint` | filePath | **itemId, driveId**, shareUrl | Upload new file |
| `update_sharepoint_file` | **itemId, driveId**, filePath | success | Update existing file |
| `grant_file_permission` | **itemId, driveId**, emails, role | success | Grant access |

### Tool Details

#### `download_sharepoint_file`
- **Input**: URL (any format - browser, sharing link, Teams)
- **Output**: `{ path, size, itemId, driveId }`
- **Purpose**: Entry point - converts URL to stable IDs

#### `upload_to_sharepoint`
- **Input**: `filePath` (local file)
- **Output**: `{ itemId, driveId, shareUrl }`
- **Note**: Uploads to agent's OneDrive, creates view-only link

#### `update_sharepoint_file`
- **Input**: `{ itemId, driveId, filePath }`
- **Output**: `{ success, filename, size }`
- **Requires**: Edit permissions on file

#### `grant_file_permission`
- **Input**: `{ itemId, driveId, emails[], role }` (role is mandatory: 'read' or 'write')
- **Output**: `{ success, message }`
- **Note**: Accepts multiple emails, works on any drive where you have permission to share

---

## Use Cases

### UC1: Download and Analyze
```
User: "Analyze this file: [URL]"
Agent: download_sharepoint_file(url) → { path, itemId, driveId }
       Read and analyze local file
```

### UC2: Upload New File
```
User: "Create a report for me"
Agent: Create file locally
       upload_to_sharepoint(filePath) → { itemId, driveId, shareUrl }
       Share URL with user
```

### UC3: Update Shared File
```
User: "Update this spreadsheet: [URL]"
Agent: download_sharepoint_file(url) → { path, itemId, driveId }
       Modify file locally
       update_sharepoint_file(itemId, driveId, filePath) → success
```

### UC4: Grant Permissions
```
User: "Share this file with john@company.com"
Agent: (already has itemId, driveId from previous download)
       grant_file_permission(itemId, driveId, email, 'write') → success
```

### UC5: Create Copy with New Name
```
Agent: download_sharepoint_file(url) → local file
       Rename locally
       upload_to_sharepoint(newPath) → new file in agent's drive
```

---

## Error Handling

### Access Denied (403)
```json
{
  "success": false,
  "error": "Access denied. You don't have permission to access this file. Ask the file owner to share it with you."
}
```

### File Locked (423 / notAllowed)
```json
{
  "success": false,
  "error": "File is locked. It may be open in a browser or desktop app. Ask the owner to close it and try again."
}
```

### User Not Found
```json
{
  "success": false,
  "error": "User not found. The email address could not be resolved in this organization."
}
```

**Note**: `bypass-shared-lock` header does NOT work when file is actively open.

---

## Test Scenarios

### TS1: Download via Browser URL
```
User: "Can you analyze this file? https://...Doc.aspx?sourcedoc=..."
Agent: download_sharepoint_file(url) → { path, itemId, driveId }
Expected: File downloaded, IDs returned for later use
```

### TS2: Download via Sharing Link
```
User: "Check this out: https://.../:x:/g/.../TOKEN?e=CODE"
Agent: download_sharepoint_file(url) → { path, itemId, driveId }
Expected: Same result as browser URL
```

### TS3: Update User's File
```
User: "Add a row to this spreadsheet: [URL]"
Agent: download_sharepoint_file(url) → { path, itemId, driveId }
       Modify file locally
       update_sharepoint_file(itemId, driveId, path) → success
Expected: File updated at original location
```

### TS4: Update Fails - No Permission
```
User: "Update this file: [URL]" (user didn't grant write access)
Agent: download_sharepoint_file(url) → success
       update_sharepoint_file(itemId, driveId, path) → error
Expected: "Access denied. You don't have edit permissions..."
```

### TS5: Update Fails - File Locked
```
User: "Update this file: [URL]" (file open in browser)
Agent: update_sharepoint_file(itemId, driveId, path) → error
Expected: "File is locked. It may be open in a browser..."
```

### TS6: Create Copy of User's File
```
User: "Make a copy of this for me: [URL]"
Agent: download_sharepoint_file(url) → { path }
       Rename file locally
       upload_to_sharepoint(newPath) → { itemId, driveId, shareUrl }
Expected: New file in agent's drive, share URL returned
```

### TS7: Share File with Colleague
```
User: "Share this with john@company.com: [URL]"
Agent: download_sharepoint_file(url) → { itemId, driveId }
       grant_file_permission(itemId, driveId, "john@company.com", "read")
Expected: Permission granted
```

### TS8: Grant Write Access to User
```
User: "Let me edit the file you created"
Agent: (has itemId, driveId from upload)
       grant_file_permission(itemId, driveId, user_email, "write")
Expected: User can now edit
```

### TS9: Share Fails - Invalid Email
```
User: "Share with fake@nonexistent.com"
Agent: grant_file_permission(itemId, driveId, "fake@nonexistent.com", "read")
Expected: "User not found. The email address could not be resolved..."
```

---

## Unit Test Results (2025-12-23)

| # | Test | Result |
|---|------|--------|
| 1 | Download with browser URL | PASSED |
| 2 | Download with sharing link | PASSED |
| 3 | Update file (file closed) | PASSED |
| 4 | Update file (file open) | FAILED (expected - locked) |
| 5 | Grant read permission | PASSED |
| 6 | Grant write permission | PASSED |
| 7 | Grant permission on other's drive | PASSED |
| 8 | Error handling - access denied | PASSED |
| 9 | Error handling - file locked | PASSED |
| 10 | Error handling - user not found | PASSED |

---

## API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /shares/{encoded}/driveItem` | Resolve URL to itemId + driveId |
| `GET /shares/{encoded}/driveItem/content` | Download file content |
| `PUT /me/drive/root:/{filename}:/content` | Upload new file |
| `POST /me/drive/items/{id}/createLink` | Create share link |
| `PUT /drives/{driveId}/items/{itemId}/content` | Update existing file |
| `POST /drives/{driveId}/items/{itemId}/invite` | Grant permissions |

---

## OAuth Scopes

| Scope | Purpose | Status |
|-------|---------|--------|
| `Files.ReadWrite.All` | Read/write all accessible files | ACTIVE |

---

## Implementation Status

### Completed
- [x] Error handling with clear messages
- [x] Support for all URL formats
- [x] Download from any SharePoint/OneDrive URL
- [x] Update files with edit permission
- [x] Grant permissions on any accessible file

### To Implement
- [x] Return `itemId` + `driveId` from download
- [x] Return `driveId` from upload
- [x] Change update tool to use IDs instead of URL
- [x] Add `driveId` parameter to grant_file_permission

---

## Agent Prompt Guidance

| User Says | Tool | Notes |
|-----------|------|-------|
| "Analyze this file: [URL]" | `download_sharepoint_file` | Returns IDs for later |
| "Create a report for me" | `upload_to_sharepoint` | Returns IDs + share URL |
| "Update this spreadsheet" | `update_sharepoint_file` | Use IDs from download |
| "Share with john@..." | `grant_file_permission` | Use IDs, specify role |
| "Let me edit the file" | `grant_file_permission` | role='write' |
