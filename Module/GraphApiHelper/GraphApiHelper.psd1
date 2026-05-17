@{
    RootModule = 'GraphApiHelper.psm1'
    ModuleVersion = '1.0.6'
    GUID = 'f7f4d1f4-1b1b-4b1b-8b1b-1b1b1b1b1b1b'
    Author = 'Jiri Formacek'
    CompanyName = 'GreyCorbel Solutions'
    CompatiblePSEditions = @('Core')
    Description = 'Module provides simple commands for working with Microsoft Graph API, such as GET/POST/PATCH/DELETE requests, handling large file upload, retry logic, etc.'
    PowerShellVersion = '5.1'
    
    FunctionsToExport = @('Add-GraphLargeFile','Add-GraphReference','Get-GraphAuthorizationHeader','Get-GraphData','Invoke-GraphBatch','Invoke-GraphWithRetry','New-GraphBatchRequest','New-GraphUri','Remove-GraphReference','Set-GraphAadFactory','Set-GraphAiLogger','Set-GraphBaseUri','Set-GraphScopes')
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    
    PrivateData = @{
        PSData = @{
            Tags = @('Microsoft', 'Graph', 'Office365')
            LicenseUri = 'https://github.com/GreyCorbel/GraphApiHelper/blob/main/LICENSE'
            ProjectUri = 'https://github.com/GreyCorbel/GraphApiHelper'
            ReleaseNotes = 'Initial release'
            ExternalModuleDependencies = @('AadAuthenticationFactory')
            # Prerelease = ''
        }
    }
}

