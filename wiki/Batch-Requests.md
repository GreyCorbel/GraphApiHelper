# Batch Requests

Microsoft Graph supports combining up to **20 subrequests** into a single HTTP call via the `/$batch` endpoint. This reduces round-trips and can significantly improve performance for scenarios that require many independent operations.

---

## New-GraphBatchRequest

Creates a single subrequest object for use in a Graph batch payload.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Id` | `String` | **Yes** | Unique identifier for this subrequest within the batch. Used for response correlation and `DependsOn` ordering. |
| `Method` | `String` | **Yes** | HTTP method: `GET`, `POST`, `PUT`, `PATCH`, `DELETE`. |
| `Url` | `String` | **Yes** | **Relative** Graph path (e.g. `/me` or `/users?$top=5`). Absolute URLs are rejected. |
| `Body` | `Object` | No | Request body for `POST`, `PUT`, `PATCH`. When provided and `Content-Type` is not set in `Headers`, it defaults to `application/json`. |
| `Headers` | `Hashtable` | No | Per-subrequest headers (e.g. `ConsistencyLevel`). |
| `DependsOn` | `String[]` | No | List of subrequest IDs that must complete before this one is executed. |

### Outputs

`PSCustomObject` (typed as `GraphApiHelper.GraphBatchRequest`) suitable for passing to `Invoke-GraphBatch`.

### Examples

```powershell
# Simple GET subrequest
$req = New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
```

```powershell
# PATCH subrequest with a body
$req = New-GraphBatchRequest -Id '2' -Method PATCH `
    -Url '/users/john.doe@contoso.com' `
    -Body @{ jobTitle = 'Principal Engineer' }
```

```powershell
# POST to create a group
$req = New-GraphBatchRequest -Id '3' -Method POST -Url '/groups' -Body @{
    displayName    = 'Batch Group'
    mailEnabled    = $false
    mailNickname   = 'batch-group'
    securityEnabled = $true
}
```

```powershell
# Subrequest that depends on another
$createGroup  = New-GraphBatchRequest -Id 'create' -Method POST -Url '/groups' -Body $groupBody
$addMember    = New-GraphBatchRequest -Id 'addMember' -Method POST `
    -Url '/groups/{id}/members/$ref' `
    -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$userId" } `
    -DependsOn @('create')
```

```powershell
# Use New-GraphUri to build a query URL for a subrequest
$url = New-GraphUri -Uri '/users' -WithSelect 'id,displayName' -Top 5 -Relative
$req = New-GraphBatchRequest -Id '4' -Method GET -Url $url
```

---

## Invoke-GraphBatch

Collects batch subrequests, sends them to the `/$batch` endpoint, and returns the array of response objects.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `BatchRequest` | `PSCustomObject[]` | **Yes** | One or more request objects from `New-GraphBatchRequest`. Accepts pipeline input. Alias: `-Requests`. |
| `RetryableErrorCodes` | `Int[]` | No | HTTP status codes on the outer batch call to retry. Default: `429`. |
| `RequestHeaders` | `Hashtable` | No | Headers for the outer `/$batch` HTTP request. |
| `OperationName` | `String` | No | Name logged to Application Insights. Default: `Invoke-GraphBatch`. |
| `AuthorizationHeader` | `Hashtable` | No | Pre-obtained authorization header (e.g. from `Get-GraphAuthorizationHeader`). When provided, token acquisition is skipped. Useful for reusing a token across many calls. |

### Outputs

`System.Object[]` — the `responses` array from the Graph batch response. Each item has `id`, `status`, `headers`, and `body` properties.

### Examples

```powershell
# Build and send a batch of two GET requests
$requests = @(
    New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
    New-GraphBatchRequest -Id '2' -Method GET -Url '/organization'
)

$responses = Invoke-GraphBatch -BatchRequest $requests
$responses | Select-Object id, status
```

```powershell
# Pipeline input
@(
    New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
    New-GraphBatchRequest -Id '2' -Method GET -Url (New-GraphUri -Uri '/users' -Top 5 -Relative)
    New-GraphBatchRequest -Id '3' -Method POST -Url '/groups' -Body @{
        displayName    = 'Pipeline Group'
        mailEnabled    = $false
        mailNickname   = 'pipeline-group'
        securityEnabled = $true
    }
) | Invoke-GraphBatch
```

```powershell
# Check individual subrequest results
$responses = Invoke-GraphBatch -BatchRequest $requests
foreach ($r in $responses) {
    if ($r.status -ne 200 -and $r.status -ne 201) {
        Write-Warning "Request $($r.id) failed with status $($r.status): $($r.body.error.message)"
    }
}
```

```powershell
# Retry outer batch call on both throttling and transient errors
$responses = Invoke-GraphBatch -BatchRequest $requests -RetryableErrorCodes 429, 503
```

```powershell
# Preview the batch call without sending it
Invoke-GraphBatch -BatchRequest $requests -WhatIf
```

### Constraints

| Constraint | Value |
|---|---|
| Maximum subrequests per batch | 20 |
| Minimum subrequests | 1 |
| Duplicate `Id` values | Not allowed (throws) |
| Absolute URLs in subrequests | Not allowed (use relative paths) |

### Notes

- The `status` in each response item is the HTTP status of that subrequest, not the outer batch call.
- A `200 OK` outer response does not mean every subrequest succeeded — always inspect each item's `status`.
- Use `DependsOn` in subrequests to sequence dependent operations (e.g. create then update).
- For more than 20 operations, split into multiple `Invoke-GraphBatch` calls.
