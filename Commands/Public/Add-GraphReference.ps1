function Add-GraphReference
{
    <#
    .SYNOPSIS
    Adds a reference to a Microsoft Graph object.

    .DESCRIPTION
    Adds a reference to a Microsoft Graph group, application, or service principal.
    This is typically used to add members or owners by creating the corresponding $ref link.

    .PARAMETER ObjectId
    The identifier of the Microsoft Graph object that will receive the reference.

    .PARAMETER objectType
    The Microsoft Graph object type. Valid values are groups, applications, and servicePrincipals.

    .PARAMETER ReferenceType
    The reference collection to update. Valid values are members and owners.

    .PARAMETER MemberId
    The identifier of the object being referenced, such as a user, group, or service principal.

    .PARAMETER PermissiveModify
    Suppresses errors when the reference already exists.

    .INPUTS
    System.String
    Accepts MemberId values from the pipeline.

    .OUTPUTS
    None
    This command performs a Graph API call and does not emit output.

    .EXAMPLE
    Add-GraphReference -ObjectId $groupId -MemberId $userId

    Adds the specified user as a member of the group.

    .EXAMPLE
    Add-GraphReference -ObjectId $groupId -ReferenceType owners -MemberId $userId -PermissiveModify

    Adds the specified user as a group owner and ignores the request if the reference already exists.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $ObjectId,
        [Parameter()]
        [ValidateSet('groups','applications','servicePrincipals')]
        [string]$objectType = 'groups',
        [Parameter()]
        [ValidateSet('members', 'owners')]
        [string]$ReferenceType = 'members',
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$MemberId,
        [switch]$PermissiveModify
    )

    begin
    {
        $uri = New-GraphUri -Uri "/$objectType/$ObjectId/$ReferenceType/`$ref"
    }
    process
    {
        $body = @{
            "@odata.id" = $script:graphConnection.GetReference($MemberId)
        } | ConvertTo-Json
        try
        {
            # we want this to throw, so to honor the -PermissiveModify switch
            Invoke-GraphWithRetry -Method Post -Uri $uri -Body $body -ErrorAction Stop
            Write-Verbose "User with ID $MemberId added to $ReferenceType of $ObjectId."
        }
        catch
        {
            $details = $_ | ConvertFrom-GraphErrorRecord
            if($details.error.message -match 'object references already exist' -and $PermissiveModify)
            {
                Write-Verbose -Message "User with ID $MemberId is already a $ReferenceType of $ObjectId."
            }
            else
            {
                Write-Error -ErrorRecord $_
            }
        }
    }
}
