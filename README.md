# GraphApiHelper

PowerShell module that provides a robust, production-ready wrapper around the Microsoft Graph API. It handles authentication token acquisition, automatic pagination, throttling retries, large file uploads, and optional Application Insights telemetry — so callers can focus on business logic rather than HTTP plumbing.

## Dependencies

| Module | Purpose |
|---|---|
| [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory) | Obtains Azure AD access tokens (supports managed identity, service principal, interactive, etc.) |
| [AiLogger](https://github.com/GreyCorbel/AiLogger) | Optional Application Insights telemetry logging |

Install dependencies from the PowerShell Gallery before using this module:

```powershell
Install-Module -Name AadAuthenticationFactory
Install-Module -Name AiLogger          # optional, only needed for telemetry
Install-Module -Name GraphApiHelper
```

## Quick Start

### 1. Configure authentication

Use `AadAuthenticationFactory` to create and register an authentication factory, then point `GraphApiHelper` at it:

```powershell
# Create a factory that uses a managed identity (e.g. in Azure Automation / Azure Functions)
New-AadAuthenticationFactory -Name 'ManagedIdentity' -UseManagedIdentity

# Tell GraphApiHelper which factory to use
Set-GraphAadFactory -Name 'ManagedIdentity'
```

### 2. Read data

```powershell
# Retrieve a single user — returns the user object directly
Get-GraphData -RequestUri '/users/john.doe@contoso.com'

# Retrieve all users — automatically pages through all result pages
$users = Get-GraphData -Uri '/users' -WithSelect 'displayName,userPrincipalName,mail'
$users | Select-Object displayName, mail
```

### 3. Write data

```powershell
# Create a new security group
$body = @{
    displayName     = 'Finance Team'
    mailEnabled     = $false
    mailNickname    = 'finance-team'
    securityEnabled = $true
} | ConvertTo-Json

Invoke-GraphWithRetry -RequestUri '/groups' -Method Post -Body $body

# Update a user's job title
$patch = @{ jobTitle = 'Senior Engineer' } | ConvertTo-Json
Invoke-GraphWithRetry -RequestUri '/users/john.doe@contoso.com' -Method Patch -Body $patch

# Delete a group
Invoke-GraphWithRetry -RequestUri "/groups/$groupId" -Method Delete
```

## Commands

### `Set-GraphAadFactory`
Registers which `AadAuthenticationFactory` factory name to use for token acquisition.

```powershell
Set-GraphAadFactory -Name 'ManagedIdentity'
```

### `Set-GraphScopes`
Overrides the OAuth2 scope used when requesting tokens. The default (`https://graph.microsoft.com/.default`) covers all permissions granted to the application.

```powershell
# Use default application permissions (recommended for service accounts / managed identity)
Set-GraphScopes -Scopes 'https://graph.microsoft.com/.default'
```

### `Set-GraphBaseUri`
Sets the base URI prepended to relative request paths. Defaults to `https://graph.microsoft.com/v1.0`. Change this to target sovereign clouds or the beta endpoint.

```powershell
# US Government cloud
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'

# Beta endpoint
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/beta'
```

### `Get-GraphData`
Issues a GET request and **automatically follows all `@odata.nextLink` pages**, returning the complete dataset.

```powershell
# All members of a group (handles pages transparently)
$members = Get-GraphData -RequestUri "/groups/$groupId/members"

# Query options via Get-GraphData parameters
$guests = Get-GraphData `
    -RequestUri '/users' `
    -WithSelect 'id,displayName,userPrincipalName' `
    -WithFilter "userType eq 'Guest'" `
    -WithCount `
    -WithExpand 'manager($select=id,displayName)' `
    -WithSearch '"displayName:alex"' `
    -Top 25

# Retrieve only the first page even when @odata.nextLink is present
$firstPage = Get-GraphData -RequestUri '/users' -Top 10 -NoContinue
```

### `New-GrapUri`
Builds a request URL using the same query option parameters as `Get-GraphData`, without sending a request. Use `-Relative` when you need a batch subrequest URL.

```powershell
$url = New-GrapUri `
    -Uri '/users' `
    -WithSelect 'id,displayName,userPrincipalName' `
    -WithFilter "userType eq 'Guest'" `
    -WithCount `
    -Top 25

$url

$batchUrl = New-GrapUri `
    -Uri '/users' `
    -WithSelect 'id,displayName' `
    -Top 5 `
    -Relative

# Absolute URI with -Relative extracts just the path (useful if you have full URLs from logs)
$batchUrl2 = New-GrapUri `
    -Uri 'https://graph.microsoft.com/v1.0/users' `
    -WithSelect 'id,displayName' `
    -Top 5 `
    -Relative

# $batchUrl and $batchUrl2 are equivalent: /users?$select=id,displayName&$top=5
```

### `Invoke-GraphWithRetry`
Issues any HTTP method against Graph. Automatically retries on **HTTP 429 (throttling)** using the `Retry-After` header, up to 100 times.

```powershell
# Send a Teams chat message
$body = @{
    body = @{ content = 'Hello from GraphApiHelper!' }
} | ConvertTo-Json -Depth 5

Invoke-GraphWithRetry `
    -RequestUri "/chats/$chatId/messages" `
    -Method Post `
    -Body $body
```

### `Invoke-GraphBatch`
Collects multiple Graph request definitions, sends a single `/$batch` call via `Invoke-GraphWithRetry`, and returns individual batch response items.

Microsoft Graph limits each batch to 20 subrequests.

```powershell
$requests = @(
    New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
    New-GraphBatchRequest -Id '2' -Method GET -Url (New-GrapUri -Uri '/users' -Top 5 -Relative)
    New-GraphBatchRequest -Id '3' -Method PATCH -Url '/users/john.doe@contoso.com' -Body @{ jobTitle = 'Principal Engineer' }
)

$responses = Invoke-GraphBatch -BatchRequest $requests
$responses | Select-Object id, status
```

### `New-GraphBatchRequest`
Creates a structured Graph batch subrequest object to pass to `Invoke-GraphBatch`. The `-Id` parameter is required.

```powershell
$req1 = New-GraphBatchRequest -Id '1' -Method GET -Url '/users?$top=5'
$req2 = New-GraphBatchRequest -Id '2' -Method GET -Url '/groups?$top=5' -DependsOn '1'

Invoke-GraphBatch -BatchRequest @($req1, $req2)
```

### `Add-GraphLargeFile`
Uploads a local file of any size to OneDrive or SharePoint using Graph's **resumable upload session** protocol (5 MB chunks). Existing files are replaced automatically.

```powershell
# Upload a report to the current user's OneDrive
Add-GraphLargeFile `
    -LocalFilePath 'C:\Reports\annual-report.xlsx' `
    -GraphFilePath '/me/drive/root:/Reports/annual-report.xlsx' `
    -Verbose

# Upload to a SharePoint document library
Add-GraphLargeFile `
    -LocalFilePath 'C:\Videos\onboarding.mp4' `
    -GraphFilePath "/sites/$siteId/drive/root:/Training/onboarding.mp4"
```

### `Get-GraphAuthorizationHeader`
Returns a hashtable containing the `Authorization: Bearer <token>` header. Primarily used internally, but useful when you need to call a Graph endpoint with your own `Invoke-RestMethod`.

```powershell
$headers = Get-GraphAuthorizationHeader
Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/me' -Headers $headers
```

### `Set-GraphAiLogger`
Attaches an `AiLogger` instance for Application Insights telemetry. All Graph calls will emit dependency telemetry under the configured operation name.

```powershell
$logger = Connect-AiLogger -ConnectionString '<your-connection-string>'
Set-GraphAiLogger -Logger $logger
```

## Notes

- **Relative URIs** are supported everywhere — the configured `BaseUri` is automatically prepended to any path that does not start with `http`.
- **Pagination** is handled transparently by `Get-GraphData`. Use `Get-GraphData -NoContinue`, or `Invoke-GraphWithRetry` when you need only a single page.
- **Throttle protection** — `Invoke-GraphWithRetry` honours the `Retry-After` response header and backs off accordingly.
- Only **PowerShell Core** (`CompatiblePSEditions = Core`) is supported.
