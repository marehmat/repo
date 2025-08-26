@{
    RootModule = 'SPMigrationDB-DatabaseAdmin.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author = 'SharePoint Migration Team'
    CompanyName = 'Your Organization'
    Copyright = '(c) 2025 Your Organization. All rights reserved.'
    Description = 'PowerShell module for SQLite database administration for SharePoint/OneDrive migration tracking. Provides CRUD operations for database and table management.'
    PowerShellVersion = '5.0'
    FunctionsToExport = @(
        'New-MigrationDatabase',
        'Get-MigrationDatabaseSchema',
        'Test-MigrationDatabase',
        'Get-MigrationDatabaseTables',
        'Backup-MigrationDatabase',
        'Remove-MigrationDatabase'
    )
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    RequiredAssemblies = @('System.Data.SQLite.dll')
    FileList = @(
        'SPMigrationDB-DatabaseAdmin.psm1',
        'SPMigrationDB-DatabaseAdmin.psd1',
        'System.Data.SQLite.dll'
    )
    PrivateData = @{
        PSData = @{
            Tags = @('SharePoint', 'OneDrive', 'Migration', 'SQLite', 'Database', 'Administration')
            LicenseUri = ''
            ProjectUri = ''
            IconUri = ''
            ReleaseNotes = 'Initial release of SharePoint Migration Database Administration module'
        }
    }
}