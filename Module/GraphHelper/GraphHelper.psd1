@{
    RootModule = 'GraphHelper.psm1'
    ModuleVersion = '1.0.1'
    GUID = 'f7f4d1f4-1b1b-4b1b-8b1b-1b1b1b1b1b1b'
    Author = 'Jiri Formacek'
    CompanyName = 'GreyCorbel Solutions'
    CompatiblePSEditions = @('Core')
    Description = 'Module provides basic operation with Microsoft Graph API, such as GET/POST/PATCH/DELETE requests, handling large file upload, retry logic, etc.'
    PowerShellVersion = '5.1'
    
    FunctionsToExport = @('Add-GraphLargeFile','Get-GraphAuthorizationHeader','Get-GraphData','Invoke-GraphWithRetry','Set-GraphAadFactory','Set-GraphAiLogger','Set-GraphBaseUri','Set-GraphScopes')
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    
    PrivateData = @{
        PSData = @{
            Tags = @('Microsoft', 'Graph', 'Office365')
            LicenseUri = 'https://github.com/GreyCorbel/GraphHelper/blob/main/LICENSE'
            ProjectUri = 'https://github.com/GreyCorbel/GraphHelper'
            ReleaseNotes = 'Initial release'
            ExternalModuleDependencies = @('AadAuthenticationFactory','AiLogging')
            # Prerelease = ''
        }
    }
}

