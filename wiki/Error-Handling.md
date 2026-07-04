# Error Handling

## ConvertFrom-GraphErrorRecord

Extracts and deserializes the Microsoft Graph error payload from a PowerShell `ErrorRecord`.

When a Graph API call fails, the HTTP response body contains a JSON object with an `error` key holding `code`, `message`, and sometimes `innerError` details. PowerShell stores this raw JSON in `$_.ErrorDetails`. `ConvertFrom-GraphErrorRecord` parses it so you can inspect and act on the specific Graph error.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `ErrorRecord` | `ErrorRecord` | **Yes** | A PowerShell `ErrorRecord`. Accepts pipeline input (`$_` in a `catch` block). |

### Outputs

`System.Object` — the deserialized Graph error object when the error details contain a valid Graph JSON error payload. Returns nothing when the payload is absent or not parseable.

### Examples

```powershell
# Basic catch and parse
try {
    Invoke-GraphWithRetry -RequestUri '/users/does-not-exist@contoso.com' -ErrorAction Stop
}
catch {
    $graphError = $_ | ConvertFrom-GraphErrorRecord
    Write-Host "Graph error code:    $($graphError.error.code)"
    Write-Host "Graph error message: $($graphError.error.message)"
}
```

```powershell
# Parse the most recent error from $Error
$details = $Error[0] | ConvertFrom-GraphErrorRecord
if ($details) {
    Write-Warning "Last Graph error: $($details.error.message)"
}
```

```powershell
# Branch on error code
try {
    Get-GraphData -RequestUri "/groups/$groupId" -ErrorAction Stop
}
catch {
    $graphError = $_ | ConvertFrom-GraphErrorRecord
    switch ($graphError.error.code) {
        'Request_ResourceNotFound' { Write-Warning "Group $groupId not found." }
        'Authorization_RequestDenied' { Write-Error "Insufficient permissions." }
        default { Write-Error -ErrorRecord $_ }
    }
}
```

```powershell
# Used internally by Add-GraphReference for PermissiveModify
try {
    Invoke-GraphWithRetry -Method Post -RequestUri $refUri -Body $body -ErrorAction Stop
}
catch {
    $details = $_ | ConvertFrom-GraphErrorRecord
    if ($details.error.message -match 'object references already exist') {
        Write-Verbose "Reference already exists, skipping."
    } else {
        Write-Error -ErrorRecord $_
    }
}
```

### Graph Error Object Structure

```json
{
  "error": {
    "code": "Request_ResourceNotFound",
    "message": "Resource 'user@contoso.com' does not exist or one of its queried reference-property objects are not present.",
    "innerError": {
      "request-id": "...",
      "date": "...",
      "client-request-id": "..."
    }
  }
}
```

### Notes

- Returns `$null` silently when the error record has no JSON details or the JSON does not contain `error.message`.
- Does not modify or re-throw the error — the caller decides what to do.
- All write operations in this module use `-ErrorAction Stop` internally when they need to inspect the error type, so you may need `-ErrorAction Stop` in your own calls to catch terminating errors.
