@{
    RootModule = 'SPMigrationDB-DataOperations.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'b2c3d4e5-f6g7-8901-bcde-f23456789012'
    Author = 'SharePoint Migration Team'
    CompanyName = 'Your Organization'
    Copyright = '(c) 2025 Your Organization. All rights reserved.'
    Description = 'PowerShell module for SQLite data operations for SharePoint/OneDrive migration tracking. Provides CRUD operations for migration data management.'
    PowerShellVersion = '5.0'
    FunctionsToExport = @(
        'New-MigrationProject',
        'Get-MigrationProject',
        'Set-MigrationProject',
        'Remove-MigrationProject',
        'Add-MigrationUser',
        'Get-MigrationUser',
        'Set-MigrationUser',
        'Add-MigrationLog',
        'Get-MigrationLog'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    RequiredAssemblies = @('System.Data.SQLite.dll')
    FileList = @(
        'SPMigrationDB-DataOperations.psm1',
        'SPMigrationDB-DataOperations.psd1',
        'System.Data.SQLite.dll'
    )
    PrivateData = @{
        PSData = @{
            Tags = @('SharePoint', 'OneDrive', 'Migration', 'SQLite', 'Database', 'DataOperations')
            LicenseUri = ''
            ProjectUri = ''
            IconUri = ''
            ReleaseNotes = 'Initial release of SharePoint Migration Data Operations module'
        }
    }
}