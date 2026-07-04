# Directory References

These commands manage `$ref` membership and ownership links for groups, applications, and service principals in Microsoft Entra ID (Azure AD).

---

## Add-GraphReference

Adds a member or owner reference to a Graph directory object.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `ObjectId` | `String` | **Yes** | The ID of the target group, application, or service principal. |
| `MemberId` | `String` | **Yes** | The object ID to add. Accepts pipeline input. |
| `objectType` | `String` | No | The Graph object type: `groups`, `applications`, `servicePrincipals`. Default: `groups`. |
| `ReferenceType` | `String` | No | The collection to add to: `members` or `owners`. Default: `members`. |
| `PermissiveModify` | `Switch` | No | Suppress the error when the reference already exists. |

### Outputs

None. Write-Verbose messages report the result.

### Examples

```powershell
# Add a user to a group
Add-GraphReference -ObjectId $groupId -MemberId $userId
```

```powershell
# Add multiple users from pipeline
$userIds | Add-GraphReference -ObjectId $groupId
```

```powershell
# Add a user as a group owner
Add-GraphReference -ObjectId $groupId -ReferenceType owners -MemberId $userId
```

```powershell
# Add a service principal to an application, ignore if already present
Add-GraphReference -ObjectId $appId -objectType applications `
    -ReferenceType members -MemberId $spId -PermissiveModify
```

---

## Remove-GraphReference

Removes a member or owner reference from a Graph directory object.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `ObjectId` | `String` | **Yes** | The ID of the target group, application, or service principal. |
| `MemberId` | `String` | **Yes** | The object ID to remove. Accepts pipeline input. |
| `objectType` | `String` | No | The Graph object type: `groups`, `applications`, `servicePrincipals`. Default: `groups`. |
| `ReferenceType` | `String` | No | The collection to remove from: `members` or `owners`. Default: `members`. |
| `PermissiveModify` | `Switch` | No | Suppress the error when the reference does not exist (HTTP 404). |

### Outputs

None. Write-Verbose messages report the result.

### Examples

```powershell
# Remove a user from a group
Remove-GraphReference -ObjectId $groupId -MemberId $userId
```

```powershell
# Remove multiple users from pipeline
$userIds | Remove-GraphReference -ObjectId $groupId
```

```powershell
# Remove an owner; ignore if not found
Remove-GraphReference -ObjectId $groupId -ReferenceType owners `
    -MemberId $ownerId -PermissiveModify
```

```powershell
# Remove a service principal from an application's members
Remove-GraphReference -ObjectId $appId -objectType applications `
    -ReferenceType members -MemberId $spId
```

---

## Get-GraphReferenceUri

Builds the fully-qualified `directoryObjects` reference URI for a given object ID. Used internally by `Add-GraphReference` and useful when constructing batch or custom payloads.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `ObjectId` | `String` | **Yes** | The Azure AD / Entra ID object identifier. |

### Outputs

`System.String` — a URI in the form `https://graph.microsoft.com/v1.0/directoryObjects/<ObjectId>`.

### Examples

```powershell
# Get the reference URI for a user
Get-GraphReferenceUri -ObjectId '11111111-2222-3333-4444-555555555555'
# → https://graph.microsoft.com/v1.0/directoryObjects/11111111-2222-3333-4444-555555555555
```

```powershell
# Use in a custom $ref body
$body = @{
    '@odata.id' = Get-GraphReferenceUri -ObjectId $userId
} | ConvertTo-Json

Invoke-GraphWithRetry -RequestUri "/groups/$groupId/members/`$ref" -Method Post -Body $body
```

### Notes

- The URI always uses the `v1.0` path segment regardless of the configured `BaseUri`, matching the Graph `$ref` link format.
- The scheme and host are derived from the currently configured `BaseUri`.
