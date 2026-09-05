# Search

Commands for querying the Microsoft Graph Search API (`/search/query`), which covers content search across
mail, drive items, sites, lists, people, Teams messages, and Search-connector external items.

---

## Invoke-GraphSearchQuery

Executes a Microsoft Graph Search query for a **single entity type** and automatically requests subsequent
result pages until Graph reports no more results are available.

Unlike `Get-GraphData`, the Search API does not use `@odata.nextLink`. Instead, each request specifies a
`from`/`size` offset and the response reports `moreResultsAvailable`. `Invoke-GraphSearchQuery` handles this
offset-based paging model for you.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `EntityType` | `String` | **Yes** | One of: `list`, `listItem`, `site`, `drive`, `driveItem`, `message`, `event`, `person`, `chatMessage`, `externalItem`, `acronym`, `bookmark`, `qna`. |
| `QueryString` | `String` | **Yes** | The search query (KQL-like syntax supported by Microsoft Search). |
| `Fields` | `String[]` | No | Properties to return per hit. Defaults to the entity type's default field set when omitted. |
| `SortProperties` | `Hashtable[]` | No | Sort specification, e.g. `@{ Name = 'lastModifiedDateTime'; IsDescending = $true }`. |
| `ContentSources` | `String[]` | No | External connection content source identifiers. Only applies to `-EntityType externalItem`. |
| `Region` | `String` | No | Geographic region required by application (non-delegated) permission searches for some entity types. |
| `EnableTopResults` | `Switch` | No | Requests relevance-ranked ordering when supported by the entity type. |
| `Size` | `Int` | No | Hits requested per page. Default: `25`. Graph enforces an entity-type-specific maximum. |
| `From` | `Int` | No | Zero-based offset of the first hit to return. Use to resume an interrupted search. Default: `0`. |
| `NoContinue` | `Switch` | No | Retrieve only the first page; do not request subsequent pages. |
| `RetryableErrorCodes` | `Int[]` | No | HTTP status codes to retry. Default: `429`. |
| `OperationName` | `String` | No | Name logged to Application Insights. Default: `Invoke-GraphSearchQuery`. |
| `AuthorizationHeader` | `Hashtable` | No | Pre-obtained authorization header (e.g. from `Get-GraphAuthorizationHeader`). When provided, token acquisition is skipped. |
| `ResponseMetadataVariable` | `String` | No | Variable name to store metadata (`Total`, `MoreResultsAvailable`, `NextFrom`, `Size`) from the last page. Also populated when the command is interrupted or stops early — `$meta.NextFrom` can be passed back in as `-From` to resume. |

### Examples

```powershell
# Search mail messages for the current user context
$hits = Invoke-GraphSearchQuery -EntityType message -QueryString 'subject:invoice'
```

```powershell
# Search drive items, requesting only specific fields, 50 hits per page
$hits = Invoke-GraphSearchQuery -EntityType driveItem -QueryString 'quarterly report' -Fields 'name','webUrl' -Size 50
```

```powershell
# Retrieve only the first page and capture the offset needed to resume
$hits = Invoke-GraphSearchQuery -EntityType driveItem -QueryString 'roadmap' -ResponseMetadataVariable 'meta' -NoContinue

if ($meta.MoreResultsAvailable) {
    $more = Invoke-GraphSearchQuery -EntityType driveItem -QueryString 'roadmap' -From $meta.NextFrom
}
```

```powershell
# Search an external connection's indexed content (application permissions)
$hits = Invoke-GraphSearchQuery -EntityType externalItem -QueryString 'contoso' `
    -ContentSources 'connections/contoso-kb' -Region 'NAM'
```

```powershell
# Sort by most recently modified first
$hits = Invoke-GraphSearchQuery -EntityType driveItem -QueryString 'budget' `
    -SortProperties @{ Name = 'lastModifiedDateTime'; IsDescending = $true }
```

```powershell
# Preview the request without calling Graph
Invoke-GraphSearchQuery -EntityType message -QueryString 'subject:invoice' -WhatIf
```

### Notes

- Queries a single entity type per call; call the command again for additional entity types.
- Uses `Invoke-GraphWithRetry` internally, so all retry behaviour applies.
- Uses the authentication factory configured via `Set-GraphAadFactory` and scopes from `Set-GraphScopes`.
- Some entity types and application (non-delegated) permission searches require `-Region`. See the
  [Microsoft Graph Search documentation](https://learn.microsoft.com/graph/search-concept-overview) for details.
