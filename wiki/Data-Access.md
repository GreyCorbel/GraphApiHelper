# Data Access

These commands cover retrieving data from Microsoft Graph and making individual API calls with retry protection.

---

## Get-GraphData

Executes a Microsoft Graph GET request and **automatically follows `@odata.nextLink` pagination**, returning the complete dataset as a flat collection.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `RequestUri` | `String` | **Yes** | Full or relative Graph URI. Alias: `-Uri`. |
| `WithSelect` | `String[]` | No | `$select` query option. Example: `'id,displayName'`. |
| `WithFilter` | `String` | No | `$filter` query option. Example: `"accountEnabled eq true"`. |
| `WithCount` | `Switch` | No | Appends `$count=true` to the request. |
| `WithExpand` | `String` | No | `$expand` query option. |
| `WithSearch` | `String` | No | `$search` query option. |
| `Top` | `Int` | No | `$top` page size hint. |
| `Skip` | `Int` | No | `$skip` offset. |
| `RetryableErrorCodes` | `Int[]` | No | HTTP status codes to retry. Default: `429`. |
| `OperationName` | `String` | No | Name logged to Application Insights. Default: `Get-GraphData`. |
| `AdditionalHeaders` | `Hashtable` | No | Extra headers added to every request in the chain. |
| `NoContinue` | `Switch` | No | Stop after the first page; do not follow `@odata.nextLink`. |
| `ResponseMetadataVariable` | `String` | No | Variable name to store metadata (`NextLink`, `DeltaLink`, `@odata.count`, etc.) from the last page. Also populated when the command is interrupted (e.g. Ctrl+C) — `meta.NextLink` will contain the URL to resume from. |
| `AuthorizationHeader` | `Hashtable` | No | Pre-obtained authorization header (e.g. from `Get-GraphAuthorizationHeader`). When provided, token acquisition is skipped. Useful when need to use access token retrieved by other means than from configured AadAuthenticationFactory instance. |

### Examples

```powershell
# Retrieve all users (automatic pagination)
$users = Get-GraphData -RequestUri '/users'
```

```powershell
# Select specific fields and filter
$enabled = Get-GraphData -RequestUri '/users' `
    -WithSelect 'id,displayName,userPrincipalName' `
    -WithFilter "accountEnabled eq true"
```

```powershell
# Advanced query requiring ConsistencyLevel header
$users = Get-GraphData -RequestUri '/users' `
    -WithFilter "startswith(displayName,'A')" `
    -WithCount `
    -AdditionalHeaders @{ ConsistencyLevel = 'eventual' }
```

```powershell
# First page only (no pagination)
$firstPage = Get-GraphData -RequestUri '/users' -Top 10 -NoContinue
```

```powershell
# Capture total count from response metadata
$members = Get-GraphData -RequestUri "/groups/$groupId/members" `
    -WithCount `
    -AdditionalHeaders @{ ConsistencyLevel = 'eventual' } `
    -ResponseMetadataVariable 'meta'
Write-Host "Total members: $($meta.Count)"
```

```powershell
# Retry on throttling AND transient service errors
$data = Get-GraphData -RequestUri '/users' -RetryableErrorCodes 429, 503
```

```powershell
# Resume an interrupted load
# Start a large load — if Ctrl+C is pressed, $meta.NextLink holds the resume URL
$users = Get-GraphData -RequestUri '/users' -ResponseMetadataVariable 'meta'

# If interrupted, $meta.NextLink is populated; resume from that URL
if ($meta.NextLink) {
    $remaining = Get-GraphData -RequestUri $meta.NextLink -ResponseMetadataVariable 'meta'
}
```

```powershell
# Preview the request without calling Graph
Get-GraphData -RequestUri '/users' -WhatIf
```

### Notes

- Uses `Invoke-GraphWithRetry` internally, so all retry behaviour applies.
- Returns individual items from `.value` when Graph returns a collection, or the raw object for single-resource responses.
- For write operations (POST/PATCH/DELETE) use `Invoke-GraphWithRetry` directly.
- Delta queries are fully supported — see [Delta Queries](#delta-queries) below.

### Delta Queries

Microsoft Graph delta queries let you track changes to a resource over time without fetching the entire collection on every run. The pattern is:

1. **Initial sync** — call the `/delta` endpoint to get all items plus a `@odata.deltaLink` at the end.
2. **Incremental syncs** — pass the saved deltaLink as `-RequestUri` to get only the items that changed (added, updated, or deleted) since the previous sync.

`Get-GraphData` handles pagination for both phases automatically. Use `-ResponseMetadataVariable` to capture the `DeltaLink` from the final page.

#### Initial sync

```powershell
# Full sync — retrieve all users and save the deltaLink for future incremental syncs
$users = Get-GraphData -RequestUri '/users/delta' `
    -WithSelect 'id,displayName,userPrincipalName,accountEnabled' `
    -ResponseMetadataVariable 'meta'

# Persist the deltaLink (e.g. to a file or database) for the next run
$meta.DeltaLink | Set-Content -Path '.\users-deltalink.txt'

Write-Host "Synced $($users.Count) users. DeltaLink saved."
```

#### Incremental sync

```powershell
# Load the deltaLink saved from the previous run
$deltaLink = Get-Content -Path '.\users-deltalink.txt' -Raw

# Fetch only changes since the last sync
$changes = Get-GraphData -RequestUri $deltaLink -ResponseMetadataVariable 'meta'

# Items with '@removed' were deleted; all others were added or updated
$deleted  = $changes | Where-Object { $_.'@removed' }
$modified = $changes | Where-Object { -not $_.'@removed' }

Write-Host "$($modified.Count) added/updated, $($deleted.Count) deleted."

# Always overwrite the saved deltaLink with the latest one
$meta.DeltaLink | Set-Content -Path '.\users-deltalink.txt'
```

#### Notes on delta queries

- The deltaLink is returned in `$meta.DeltaLink` only on the **final page** (when `@odata.nextLink` is absent). If pagination is still in progress, `DeltaLink` is `$null` and `NextLink` holds the URL of the next page.
- Delta query URLs already encode the selected properties and filter from the initial call — do not add query options when using a deltaLink as `-RequestUri`.
- Deleted items include an `@removed` property with a `reason` field (`changed` or `deleted`).
- Delta queries are available for most Entra ID resources (users, groups, directoryObjects, etc.). Refer to the [Microsoft Graph delta query documentation](https://learn.microsoft.com/graph/delta-query-overview) for the full list.
- The deltaLink token typically remains valid for up to 30 days. If it expires, start a new initial sync.

---

## Invoke-GraphWithRetry

Sends a single Microsoft Graph request with **built-in retry logic** for throttling and other transient HTTP errors.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `RequestUri` | `String` | **Yes** | Full or relative Graph URI. Alias: `-Uri`. |
| `Method` | `String` | No | HTTP method: `Get`, `Post`, `Put`, `Patch`, `Delete`. Default: `Get`. |
| `Body` | `Object` | No | Request body for `Post`, `Put`, `Patch`. |
| `ContentType` | `String` | No | Body content type. Default: `application/json`. |
| `Headers` | `Hashtable` | No | Additional HTTP headers. Authorization is added automatically. |
| `RetryableErrorCodes` | `Int[]` | No | HTTP status codes to retry. Default: `429`. |
| `MaxRetries` | `Int` | No | Maximum retry attempts before giving up. Default: `100`. |
| `DefaultBackOffSeconds` | `Int` | No | Fallback wait time (seconds) when `Retry-After` is absent. Default: `1`. |
| `OperationName` | `String` | No | Name logged to Application Insights. Default: `Invoke-GraphWithRetry`. |
| `AuthorizationHeader` | `Hashtable` | No | Pre-obtained authorization header (e.g. from `Get-GraphAuthorizationHeader`). When provided, token acquisition is skipped. Useful when need to use access token retrieved by other means than from configured AadAuthenticationFactory instance. |

### Examples

```powershell
# Simple GET
$me = Invoke-GraphWithRetry -RequestUri '/me'
```

```powershell
# Create a group
$body = @{
    displayName    = 'Finance Team'
    mailEnabled    = $false
    mailNickname   = 'finance-team'
    securityEnabled = $true
} | ConvertTo-Json

Invoke-GraphWithRetry -RequestUri '/groups' -Method Post -Body $body
```

```powershell
# Update a user's job title
$patch = @{ jobTitle = 'Senior Engineer' } | ConvertTo-Json
Invoke-GraphWithRetry -RequestUri '/users/john.doe@contoso.com' -Method Patch -Body $patch
```

```powershell
# Delete a user
Invoke-GraphWithRetry -RequestUri "/users/$userId" -Method Delete
```

```powershell
# Preview without calling Graph
Invoke-GraphWithRetry -RequestUri "/users/$userId" -Method Delete -WhatIf
```

```powershell
# Extend retries to cover 503 Service Unavailable
Invoke-GraphWithRetry -RequestUri '/users' -RetryableErrorCodes 429, 503
```

### Retry Behaviour

| Situation | Wait time |
|---|---|
| HTTP 429 with `Retry-After` header | Value of the `Retry-After` header (seconds) |
| HTTP 429 without `Retry-After` | `DefaultBackOffSeconds × retryAttempt` |
| Other retryable status code | `DefaultBackOffSeconds × retryAttempt` |

---

## Get-GraphAuthorizationHeader

Returns a hashtable containing the `Authorization: Bearer <token>` header for use with `Invoke-RestMethod` or other HTTP clients.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `FactoryName` | `String` | No | Factory name override for token acquisition. When omitted, uses the factory configured by `Set-GraphAadFactory`. |

### Examples

```powershell
$headers = Get-GraphAuthorizationHeader
Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/me' -Headers $headers
```

```powershell
# Combine with additional headers
$headers = Get-GraphAuthorizationHeader
$headers['ConsistencyLevel'] = 'eventual'
Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/users?$count=true' -Headers $headers
```

```powershell
# Override the factory for a single call
$headers = Get-GraphAuthorizationHeader -FactoryName 'ManagedIdentityFactory'
```

### Notes

- Uses the factory configured via `Set-GraphAadFactory` and scopes from `Set-GraphScopes`.
- The token is acquired fresh or from the factory's cache on each call — no manual token management is needed.
- All other GraphApiHelper commands call this internally; you only need it when making raw `Invoke-RestMethod` calls.
