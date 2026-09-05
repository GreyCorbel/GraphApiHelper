# Command Reference

Quick-reference table of all commands exported by the `GraphApiHelper` module.

## All Exported Commands

| Command | Category | Description |
|---|---|---|
| [`Set-GraphAadFactory`](Configuration#set-graphaadfactory) | Configuration | Sets the AadAuthenticationFactory instance used for token acquisition. |
| [`Set-GraphScopes`](Configuration#set-graphscopes) | Configuration | Sets the OAuth2 scopes requested when obtaining access tokens. |
| [`Set-GraphBaseUri`](Configuration#set-graphbaseuri) | Configuration | Sets the base URI prepended to relative Graph paths (must include version segment). |
| [`Set-GraphAiLogger`](Configuration#set-graphailogger) | Configuration | Attaches an AiLogger instance for Application Insights telemetry. |
| [`Get-GraphAuthorizationHeader`](Data-Access#get-graphauthorizationheader) | Authentication | Returns a hashtable with the `Authorization: Bearer` header. |
| [`Get-GraphData`](Data-Access#get-graphdata) | Data Retrieval | GET with automatic `@odata.nextLink` pagination. Returns the full dataset. |
| [`Invoke-GraphWithRetry`](Data-Access#invoke-graphwithretry) | Data Retrieval | Single Graph request (GET/POST/PUT/PATCH/DELETE) with retry logic. |
| [`New-GraphUri`](URL-Building#new-graphuri) | URL Building | Builds an absolute or relative Graph URL with OData query options. |
| [`New-GraphBatchRequest`](Batch-Requests#new-graphbatchrequest) | Batch | Creates a subrequest object for a Graph `/$batch` payload. |
| [`Invoke-GraphBatch`](Batch-Requests#invoke-graphbatch) | Batch | Sends up to 20 subrequests in a single `/$batch` call. |
| [`Invoke-GraphSearchQuery`](Search#invoke-graphsearchquery) | Search | Queries `/search/query` for a single entity type with automatic from/size result paging. |
| [`Add-GraphReference`](Directory-References#add-graphreference) | Directory | Adds a member or owner `$ref` link to a group, application, or service principal. |
| [`Remove-GraphReference`](Directory-References#remove-graphreference) | Directory | Removes a member or owner `$ref` link. |
| [`Get-GraphReferenceUri`](Directory-References#get-graphreferenceuri) | Directory | Builds a `directoryObjects` reference URI from an object ID. |
| [`Add-GraphLargeFile`](File-Upload#add-graphlargefile) | File Upload | Uploads a local file using the Graph resumable upload session protocol. |
| [`ConvertFrom-GraphErrorRecord`](Error-Handling#convertfrom-grapherrorrecord) | Error Handling | Extracts the Graph JSON error payload from a PowerShell `ErrorRecord`. |

## Common Aliases

| Alias | Resolves To |
|---|---|
| `-Uri` | `-RequestUri` (on `Get-GraphData`, `Invoke-GraphWithRetry`) |
| `-Requests` | `-BatchRequest` (on `Invoke-GraphBatch`) |

## Module Defaults

| Setting | Default |
|---|---|
| Authentication factory name | `graph` |
| Base URI | `https://graph.microsoft.com/v1.0` |
| Scope | `https://graph.microsoft.com/.default` |
| AI logger | `$null` (disabled) |
| Retry codes | `429` |
| Max retries | `100` |
| Default back-off | `1` second |

## Installation

```powershell
Install-Module AadAuthenticationFactory   # required
Install-Module AiLogger                   # optional telemetry
Install-Module GraphApiHelper
```

## Module Information

| Property | Value |
|---|---|
| Module version | 1.0.10 |
| PowerShell compatibility | Core |
| Author | Jiri Formacek |
| Company | GreyCorbel Solutions |
| PSGallery | [GraphApiHelper](https://www.powershellgallery.com/packages/GraphApiHelper) |
| Project | [github.com/GreyCorbel/GraphApiHelper](https://github.com/GreyCorbel/GraphApiHelper) |
