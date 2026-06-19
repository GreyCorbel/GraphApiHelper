<#
.SYNOPSIS
Initializes module-level Graph connection state.

.DESCRIPTION
Creates the default GraphConnection instance used by GraphApiHelper commands.
The default connection uses:
- Base URI: https://graph.microsoft.com/v1.0
- Scope: https://graph.microsoft.com/.default
- No Application Insights logger

This script runs during module import and prepares shared configuration consumed
by commands such as Invoke-GraphWithRetry, Get-GraphData, and New-GraphUri.

.INPUTS
None
This script does not accept pipeline input.

.OUTPUTS
None
Initializes module state and does not emit output.

.NOTES
Internal initialization script. Not intended to be called directly.
#>
$script:graphConnection = new-object GraphConnection('https://graph.microsoft.com/v1.0', @('https://graph.microsoft.com/.default'), $null)
