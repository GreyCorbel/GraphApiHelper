function Remove-GraphReference
{
    <#
    .SYNOPSIS
    Removes a reference from a Microsoft Graph object.

    .DESCRIPTION
    Removes a reference from a Microsoft Graph group, application, or service principal.
    This is typically used to remove members or owners by deleting the corresponding $ref link.

    .PARAMETER ObjectId
    The identifier of the Microsoft Graph object that owns the reference.

    .PARAMETER objectType
    The Microsoft Graph object type. Valid values are groups, applications, and servicePrincipals.

    .PARAMETER ReferenceType
    The reference collection to update. Valid values are members and owners.

    .PARAMETER MemberId
    The identifier of the object being removed from the reference collection.

    .PARAMETER PermissiveModify
    Suppresses errors when the reference does not exist.

    .EXAMPLE
    Remove-GraphReference -ObjectId $groupId -MemberId $userId

    Removes the specified user from the group members collection.

    .EXAMPLE
    Remove-GraphReference -ObjectId $groupId -ReferenceType owners -MemberId $userId -PermissiveModify

    Removes the specified user from the group owners collection and ignores the request if the reference is already missing.
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
    }
    process
    {
        $uri = New-GraphUri -Uri "/$objectType/$ObjectId/$ReferenceType/$MemberId/`$ref"
        try
        {
            Invoke-GraphWithRetry -Method Delete -Uri $uri -ErrorAction Stop
            Write-Verbose "User with ID $MemberId removed from $ReferenceType of $ObjectId."
        }
        catch
        {
            $ex = $_.Exception
            if($ex.Response.StatusCode -eq 404 -and $PermissiveModify)
            {
                Write-Verbose -Message "User with ID $MemberId is not in $ReferenceType of $ObjectId."
            }
            else
            {
                Write-Error -ErrorRecord $_
            }
        }
    }
}
