function Add-GraphLargeFile
{
    <#
    .SYNOPSIS
    Uploads large files to Microsoft Graph using the resumable upload protocol
    
    .DESCRIPTION
    Uploads large files to Microsoft Graph (OneDrive, SharePoint, etc.) using the resumable upload session API.
    This function handles files of any size by splitting them into chunks and uploading them sequentially.
    
    The upload uses 5MB chunks (320KB * 16) which is optimal for Graph API uploads and supports resumable uploads
    in case of network interruptions. The function automatically creates an upload session and manages the chunked
    upload process.
    
    .PARAMETER LocalFilePath
    The full path to the local file to upload. The file must exist and be readable.
    
    .PARAMETER GraphFilePath
    The Microsoft Graph API path where the file should be uploaded, excluding the ':/createUploadSession' suffix.
    Example: 'https://graph.microsoft.com/v1.0/me/drive/root:/Documents/myfile.pdf'
    
    .EXAMPLE
    Add-GraphLargeFile -LocalFilePath 'C:\Files\presentation.pptx' -GraphFilePath 'https://graph.microsoft.com/v1.0/me/drive/root:/Documents/presentation.pptx'
    
    Uploads a PowerPoint file to the current user's OneDrive Documents folder.
    
    .EXAMPLE
    Add-GraphLargeFile -LocalFilePath 'C:\Videos\training.mp4' -GraphFilePath 'https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/Videos/training.mp4' -Verbose
    
    Uploads a video file to a SharePoint site's Videos folder with verbose output showing upload progress.

    .OUTPUTS
    None
    The function streams upload chunk requests and does not emit the final driveItem object.

    .INPUTS
    None
    This command does not accept pipeline input.
    
    .NOTES
    - Uses 5MB chunks for optimal performance
    - Automatically handles upload session creation
    - Supports conflict behavior of 'replace' - existing files will be overwritten
    - Uses Invoke-GraphWithRetry internally for reliability
    - Enable -Verbose to see detailed upload progress
    - Uses the authentication factory configured via Set-GraphAadFactory
    
    .LINK
    https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $LocalFilePath,
        [Parameter(Mandatory)]
        $GraphFilePath
    )

    begin
    {
        $chunkSize = 320KB * 16 # 5MB chunks
        $graphUri = New-GraphUri -Uri "$GraphFilePath"
    }
    process
    {
        $item = Get-Item -Path $LocalFilePath
        $fileSize = $item.length
        $fileStream = [System.IO.File]::OpenRead($item.FullName)
        Write-Verbose "Filesize: $fileSize"
        Write-Verbose "Chunksize: $chunkSize"
        try {
            $payload =  @{
                item = @{
                    '@microsoft.graph.conflictBehavior' = 'replace' 
                }
            }
            Write-Verbose "Requesting upload session on $graphUri`:/createUploadSession"
            $uploadSession = Invoke-GraphWithRetry `
                -RequestUri "$graphUri`:/createUploadSession" `
                -method Post `
                -body ($payload | ConvertTo-Json -Depth 10) `
                -ContentType 'application/json' `
                -ErrorAction Stop
    
            $uploadUrl = $uploadSession.uploadUrl
            Write-Verbose "UploadUrl: $uploadUrl"
            $offset = 0
            
            while ($offset -lt $fileSize) {
                $bytesToRead = [Math]::Min($chunkSize, $fileSize - $offset)
                $buffer = New-Object byte[] $bytesToRead
                $bytesRead = $fileStream.Read($buffer, 0, $bytesToRead)
    
                if ($bytesRead -gt 0) {
                    $contentRange = "bytes $offset-$($offset + $bytesRead - 1)/$fileSize"
                    Write-Verbose "Writing range: $contentRange"
                    Invoke-GraphWithRetry `
                        -RequestUri $uploadUrl `
                        -method Put `
                        -body $buffer `
                        -headers @{ 'Content-Range' = $contentRange } `
                        -ContentType 'application/octet-stream' | out-null
                    $offset += $bytesRead
                }
            }
        }
        finally {
            $fileStream.Close()
        }
    }
}
