# File Upload

## Add-GraphLargeFile

Uploads a local file to Microsoft Graph (OneDrive, SharePoint, Teams, etc.) using the **resumable upload session** protocol. The file is split into 5 MB chunks and uploaded sequentially.

> **Tip:** Use `-Verbose` to monitor upload progress chunk by chunk.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `LocalFilePath` | `String` | **Yes** | Full local path to the file to upload. The file must exist and be readable. |
| `GraphFilePath` | `String` | **Yes** | Microsoft Graph API path for the destination, **without** the `:/createUploadSession` suffix. |
| `ConflictBehavior` | `String` | No | What to do when a file already exists: `replace` (default), `rename`, or `fail`. |

### Outputs

None. The function streams chunk requests silently. Use `-Verbose` for progress.

### Examples

```powershell
# Upload to the current user's OneDrive
Add-GraphLargeFile `
    -LocalFilePath 'C:\Reports\annual-report.xlsx' `
    -GraphFilePath '/me/drive/root:/Reports/annual-report.xlsx'
```

```powershell
# Upload with verbose progress
Add-GraphLargeFile `
    -LocalFilePath 'C:\Videos\training.mp4' `
    -GraphFilePath 'https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/Videos/training.mp4' `
    -Verbose
```

```powershell
# Rename instead of replace on conflict
Add-GraphLargeFile `
    -LocalFilePath 'C:\Exports\data.csv' `
    -GraphFilePath '/me/drive/root:/Exports/data.csv' `
    -ConflictBehavior rename
```

```powershell
# Fail on conflict (useful to prevent accidental overwrites)
Add-GraphLargeFile `
    -LocalFilePath 'C:\Archive\backup.zip' `
    -GraphFilePath '/me/drive/root:/Archive/backup.zip' `
    -ConflictBehavior fail
```

### Upload Session Process

1. The command calls `POST {GraphFilePath}:/createUploadSession` to obtain an upload URL.
2. The file is read in 5 MB chunks (320 KB × 16).
3. Each chunk is sent with a `Content-Range` header (`bytes start-end/total`).
4. The upload URL is valid for a limited time (typically 24 hours). If the upload fails midway, the session may expire before a retry.

### Graph Path Format

The `GraphFilePath` must point to a **driveItem path** with the file name as the last segment:

| Destination | Example path |
|---|---|
| Current user's OneDrive | `/me/drive/root:/Folder/file.txt` |
| Another user's OneDrive | `/users/{userId}/drive/root:/Folder/file.txt` |
| SharePoint site drive | `/sites/{siteId}/drive/root:/Folder/file.txt` |
| SharePoint document library by path | `/drives/{driveId}/root:/path/to/file.txt` |

You can also use absolute URLs:

```powershell
Add-GraphLargeFile `
    -LocalFilePath 'C:\file.pdf' `
    -GraphFilePath 'https://graph.microsoft.com/v1.0/me/drive/root:/Docs/file.pdf'
```

### Notes

- Uses `Invoke-GraphWithRetry` internally for session creation and each chunk PUT, so throttling is handled automatically.
- The file stream is always closed in a `finally` block, even if upload fails.
- For files under ~4 MB, a direct `Invoke-GraphWithRetry -Method Put` call to the driveItem content endpoint may be simpler.
- Reference: [Microsoft Graph — Create upload session](https://learn.microsoft.com/graph/api/driveitem-createuploadsession)
