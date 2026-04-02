$script:graphConnection = [PSCustomObject]@{
    FactoryName = 'graph'
    AiLogger = $null
    BaseUri = 'https://graph.microsoft.com/v1.0'
    GraphScope = @('https://graph.microsoft.com/.default')
}