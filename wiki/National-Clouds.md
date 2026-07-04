# National Clouds

Microsoft Graph is available in several sovereign cloud instances in addition to the global commercial cloud. Each instance has its own endpoint URL and OAuth2 scope.

Configure both `BaseUri` and `Scopes` to match your target cloud before making any API calls.

---

## Global Commercial Cloud (default)

```powershell
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/v1.0'
Set-GraphScopes  -Scopes 'https://graph.microsoft.com/.default'
```

This is the default configuration. No changes are needed unless you are targeting a different cloud or API version.

---

## Global Cloud — Beta Endpoint

```powershell
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/beta'
Set-GraphScopes  -Scopes 'https://graph.microsoft.com/.default'
```

Use this to access preview API features. Note that beta endpoints are not recommended for production use.

---

## US Government — GCC

```powershell
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/v1.0'
Set-GraphScopes  -Scopes 'https://graph.microsoft.com/.default'
```

GCC tenants use the same endpoint as the commercial cloud.

---

## US Government — GCC High (L4)

```powershell
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'
Set-GraphScopes  -Scopes 'https://graph.microsoft.us/.default'
```

---

## US Government — DoD (L5)

```powershell
Set-GraphBaseUri -BaseUri 'https://dod-graph.microsoft.us/v1.0'
Set-GraphScopes  -Scopes 'https://dod-graph.microsoft.us/.default'
```

---

## China (21Vianet)

```powershell
Set-GraphBaseUri -BaseUri 'https://microsoftgraph.chinacloudapi.cn/v1.0'
Set-GraphScopes  -Scopes 'https://microsoftgraph.chinacloudapi.cn/.default'
```

---

## Quick Reference Table

| Cloud | BaseUri | Scope |
|---|---|---|
| Global (v1.0) | `https://graph.microsoft.com/v1.0` | `https://graph.microsoft.com/.default` |
| Global (beta) | `https://graph.microsoft.com/beta` | `https://graph.microsoft.com/.default` |
| US Gov GCC | `https://graph.microsoft.com/v1.0` | `https://graph.microsoft.com/.default` |
| US Gov GCC High | `https://graph.microsoft.us/v1.0` | `https://graph.microsoft.us/.default` |
| US Gov DoD | `https://dod-graph.microsoft.us/v1.0` | `https://dod-graph.microsoft.us/.default` |
| China (21Vianet) | `https://microsoftgraph.chinacloudapi.cn/v1.0` | `https://microsoftgraph.chinacloudapi.cn/.default` |

---

## Authentication Factory Considerations

Your `AadAuthenticationFactory` instance must also target the correct authority URL and cloud for the token to be accepted by the Graph endpoint. Ensure the factory is created with the matching tenant authority:

```powershell
# GCC High example — authority must match the cloud
New-AadAuthenticationFactory `
    -Name 'GccHighFactory' `
    -TenantId 'your-tenant-id' `
    -AuthorityUrl 'https://login.microsoftonline.us' `
    -ClientId 'your-app-id' `
    -ClientSecret $secret

Set-GraphAadFactory -Name 'GccHighFactory'
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'
Set-GraphScopes  -Scopes 'https://graph.microsoft.us/.default'
```

---

## References

- [Microsoft Graph national cloud deployments](https://learn.microsoft.com/graph/deployments)
- [Microsoft identity platform — national clouds](https://learn.microsoft.com/entra/identity-platform/authentication-national-cloud)
- [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory)
