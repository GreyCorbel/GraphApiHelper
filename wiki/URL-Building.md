# URL Building

## New-GraphUri

Builds a Microsoft Graph request URL from a base path and OData query options, without making any HTTP call.

Use this command to:
- Construct absolute Graph URLs with query options for `Invoke-GraphWithRetry` or `Get-GraphData`
- Build **relative paths** for batch subrequests (`Invoke-GraphBatch`)
- Compose URLs programmatically without string concatenation

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Uri` | `String` | **Yes** | Base Graph URL or relative path (e.g. `/users`). |
| `WithSelect` | `String[]` | No | `$select` query option. Multiple values are joined with a comma. |
| `WithFilter` | `String` | No | `$filter` query option. |
| `WithCount` | `Switch` | No | Appends `$count=true`. |
| `WithExpand` | `String` | No | `$expand` query option. |
| `WithSearch` | `String` | No | `$search` query option. Values not already quoted or wrapped in `()` are automatically wrapped in `"`. |
| `Top` | `Int` | No | `$top` query option. |
| `Skip` | `Int` | No | `$skip` query option. |
| `Relative` | `Switch` | No | Return a relative path instead of an absolute URL. Strips scheme and host from absolute URIs. |

### Outputs

`System.String` — the fully constructed URL or relative path.

### Examples

```powershell
# Absolute URL from a relative path (uses configured BaseUri)
$url = New-GraphUri -Uri '/users' -Top 25
# → https://graph.microsoft.com/v1.0/users?$top=25
```

```powershell
# Add select and filter
$url = New-GraphUri -Uri '/users' `
    -WithSelect 'id,displayName,userPrincipalName' `
    -WithFilter "accountEnabled eq true"
# → https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName&$filter=accountEnabled eq true
```

```powershell
# Search with automatic quoting
$url = New-GraphUri -Uri '/users' -WithSearch 'displayName:Alex' -WithCount
# $search value is wrapped in quotes automatically
# → https://graph.microsoft.com/v1.0/users?$search="displayName:Alex"&$count=true
```

```powershell
# Preserve existing query string parameters
$url = New-GraphUri -Uri '/users?$orderby=displayName' -Top 10
# → https://graph.microsoft.com/v1.0/users?$orderby=displayName&$top=10
```

```powershell
# Relative path for a batch subrequest
$batchUrl = New-GraphUri -Uri '/users' -WithSelect 'id,displayName' -Top 5 -Relative
# → /users?$select=id,displayName&$top=5
```

```powershell
# Strip host/scheme from an absolute URL
$rel = New-GraphUri -Uri 'https://graph.microsoft.com/v1.0/users' -Relative
# → /v1.0/users
```

### URL Construction Rules

| Input | `-Relative` | Result |
|---|---|---|
| Relative path (`/users`) | `$false` | BaseUri prepended → absolute URL |
| Relative path (`/users`) | `$true` | Returned as-is |
| Absolute URL (`https://...`) | `$false` | Returned as-is |
| Absolute URL (`https://...`) | `$true` | Scheme and host stripped → relative path |

### Notes

- Query parameters are appended with `?` when no query string exists, or with `&` when one already does.
- Values are **not** URL-encoded by this function — pass pre-encoded values when required.
- `$search` values that already start with `"` or `(` are passed through unchanged.
- Call `Set-GraphBaseUri` before using relative paths without `-Relative`.
