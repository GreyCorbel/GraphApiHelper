# Configuration

These commands configure module-level settings that are shared by all GraphApiHelper commands. Call them once at the start of your script before making any Graph API calls.

---

## Set-GraphAadFactory

Sets the [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory) instance to use for obtaining access tokens.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Name` | `String` | **Yes** | The factory name registered with `New-AadAuthenticationFactory`. |
| `Force` | `Switch` | No | Skip validation that the factory currently exists. Useful when the factory is registered later in the script. |

### Examples

```powershell
# Use a managed identity factory
Set-GraphAadFactory -Name 'ManagedIdentityFactory'
```

```powershell
# Use a custom factory name
Set-GraphAadFactory -Name 'MyAppFactory'
```

```powershell
# Set name before the factory is registered
Set-GraphAadFactory -Name 'FactoryRegisteredLater' -Force
```

### Notes

- When `-Force` is not specified the command throws if the factory cannot be found in the current session.
- The configured factory name is stored in module scope and used by all subsequent Graph commands.
- Default factory name is `graph`.

---

## Set-GraphScopes

Sets the OAuth2 scopes used when requesting access tokens.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Scopes` | `String[]` | **Yes** | One or more OAuth2 scopes. |

### Examples

```powershell
# Application permissions (default)
Set-GraphScopes -Scopes 'https://graph.microsoft.com/.default'
```

```powershell
# Delegated permissions
Set-GraphScopes -Scopes @(
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Mail.Read'
)
```

```powershell
# Sovereign cloud scope
Set-GraphScopes -Scopes 'https://graph.microsoft.us/.default'
```

### Notes

- The default scope is `https://graph.microsoft.com/.default` which requests all application permissions granted to the identity.
- For delegated scenarios ensure the authentication factory is configured for interactive or device-code flow.

---

## Set-GraphBaseUri

Sets the base URI prepended to relative Graph paths. The URI must include the API version segment.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `BaseUri` | `String` | **Yes** | Full base URI including API version (e.g. `https://graph.microsoft.com/v1.0`). |

### Examples

```powershell
# Global cloud, stable version (default)
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/v1.0'
```

```powershell
# Global cloud, beta endpoint
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/beta'
```

```powershell
# US Government (GCC High)
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'
```

### Notes

- The base URI is used by `New-GraphUri` when resolving relative paths like `/users`.
- See [National Clouds](National-Clouds) for the complete list of supported endpoints.

---

## Set-GraphAiLogger

Attaches an Application Insights logger for dependency and exception telemetry.

### Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `Logger` | `Object` | **Yes** | An AiLogger connection object returned by `Connect-AiLogger`. |

### Examples

```powershell
# Enable telemetry
$logger = Connect-AiLogger -ConnectionString 'InstrumentationKey=...'
Set-GraphAiLogger -Logger $logger
```

```powershell
# Disable telemetry
Set-GraphAiLogger -Logger $null
```

### Notes

- Requires the optional [AiLogger](https://github.com/GreyCorbel/AiLogger) module.
- When configured, every Graph request is logged as a dependency event with duration, result code, and success flag.
- Exceptions are logged as Application Insights exception events.
- Set to `$null` to disable telemetry without reloading the module.
