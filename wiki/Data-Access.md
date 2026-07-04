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
| `ResponseMetadataVariable` | `String` | No | Variable name to store metadata (`@odata.count`, etc.) from the last page. |

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
# Preview the request without calling Graph
Get-GraphData -RequestUri '/users' -WhatIf
```

### Notes

- Uses `Invoke-GraphWithRetry` internally, so all retry behaviour applies.
- Returns individual items from `.value` when Graph returns a collection, or the raw object for single-resource responses.
- For write operations (POST/PATCH/DELETE) use `Invoke-GraphWithRetry` directly.

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

None.

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

### Notes

- Uses the factory configured via `Set-GraphAadFactory` and scopes from `Set-GraphScopes`.
- The token is acquired fresh or from the factory's cache on each call — no manual token management is needed.
- All other GraphApiHelper commands call this internally; you only need it when making raw `Invoke-RestMethod` calls.
