function ConvertFrom-GraphErrorRecord
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    process
    {
        $details = $ErrorRecord.ErrorDetails | ConvertFrom-Json -ErrorAction SilentlyContinue
        if($null -ne $details.error.message)
        {
            $details
        }
    }
}
