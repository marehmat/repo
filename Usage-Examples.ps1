# SharePoint Migration Database Modules - Usage Examples
# PowerShell 5.0 Compatible

# Prerequisites:
# 1. Download System.Data.SQLite.dll from https://system.data.sqlite.org/downloads/
#    - Choose the appropriate version for your system (x86/x64)
#    - Place the DLL in the same directory as the module files
# 2. Import both modules

# Import the modules
Import-Module -Name "C:\Path\To\SPMigrationDB-DatabaseAdmin.psd1" -Force
Import-Module -Name "C:\Path\To\SPMigrationDB-DataOperations.psd1" -Force

# Set variables for examples
$DatabasePath = "C:\Migration\MigrationTracking.db"

# ================================
# DATABASE ADMINISTRATION EXAMPLES
# ================================

# Example 1: Create a new migration database
New-MigrationDatabase -DatabasePath $DatabasePath -Verbose

# Example 2: Test database connection
$connectionTest = Test-MigrationDatabase -DatabasePath $DatabasePath
Write-Host "Database connection test result: $connectionTest"

# Example 3: Get database table information
$tables = Get-MigrationDatabaseTables -DatabasePath $DatabasePath
Write-Host "Database contains $($tables.Count) tables:"
$tables | Format-Table TableName, RowCount

# Example 4: Create a backup of the database
$backupPath = "C:\Migration\Backup\MigrationTracking_$(Get-Date -Format 'yyyyMMdd_HHmmss').db"
Backup-MigrationDatabase -DatabasePath $DatabasePath -BackupPath $backupPath -Verbose

# Example 5: Get the database schema (useful for documentation)
$schema = Get-MigrationDatabaseSchema
Write-Host "Database schema:"
Write-Host $schema

# ================================
# DATA OPERATIONS EXAMPLES
# ================================

# Example 6: Create a new migration project
$project = New-MigrationProject -DatabasePath $DatabasePath -ProjectName "Contoso to Fabrikam Migration" -SourceTenant "contoso.onmicrosoft.com" -TargetTenant "fabrikam.onmicrosoft.com" -CreatedBy "admin@contoso.com" -Description "Full tenant migration including OneDrive and SharePoint sites"

Write-Host "Created project with ID: $($project.ProjectId)"

# Example 7: Get all migration projects
$allProjects = Get-MigrationProject -DatabasePath $DatabasePath
Write-Host "Total projects: $($allProjects.Count)"
$allProjects | Format-Table ProjectId, ProjectName, Status, CreatedDate

# Example 8: Add users to the migration project
$users = @(
    @{ Source = "john.doe@contoso.com"; Target = "john.doe@fabrikam.com"; DataSizeGB = 2.5; ItemCount = 150 },
    @{ Source = "jane.smith@contoso.com"; Target = "jane.smith@fabrikam.com"; DataSizeGB = 1.8; ItemCount = 95 },
    @{ Source = "bob.wilson@contoso.com"; Target = "bob.wilson@fabrikam.com"; DataSizeGB = 3.2; ItemCount = 220 }
)

foreach ($user in $users) {
    $migrationUser = Add-MigrationUser -DatabasePath $DatabasePath -ProjectId $project.ProjectId -SourceUserPrincipalName $user.Source -TargetUserPrincipalName $user.Target -DataSizeGB $user.DataSizeGB -ItemCount $user.ItemCount
    
    Write-Host "Added user: $($user.Source) -> $($user.Target) (User ID: $($migrationUser.UserId))"
}

# Example 9: Get all users for the project
$projectUsers = Get-MigrationUser -DatabasePath $DatabasePath -ProjectId $project.ProjectId
Write-Host "Project has $($projectUsers.Count) users:"
$projectUsers | Format-Table UserId, SourceUserPrincipalName, TargetUserPrincipalName, Status, DataSizeGB

# Example 10: Update project status
$updatedProject = Set-MigrationProject -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Status "InProgress"
Write-Host "Project status updated to: $($updatedProject.Status)"

# Example 11: Simulate user migration progress
$firstUser = $projectUsers[0]

# Start migration for first user
$startedUser = Set-MigrationUser -DatabasePath $DatabasePath -UserId $firstUser.UserId -Status "InProgress" -StartTime (Get-Date)
Add-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Message "Started migration for user: $($firstUser.SourceUserPrincipalName)" -LogLevel "Info" -Source "Migration-Script"

# Simulate completion
Start-Sleep -Seconds 2
$completedUser = Set-MigrationUser -DatabasePath $DatabasePath -UserId $firstUser.UserId -Status "Completed" -EndTime (Get-Date)
Add-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Message "Completed migration for user: $($firstUser.SourceUserPrincipalName)" -LogLevel "Info" -Source "Migration-Script"

# Example 12: Add some log entries for tracking
Add-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Message "Migration project started" -LogLevel "Info" -Source "Start-Migration"
Add-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Message "Establishing cross-tenant relationship" -LogLevel "Info" -Details "Source: contoso.onmicrosoft.com, Target: fabrikam.onmicrosoft.com" -Source "Set-TenantRelationship"
Add-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Message "User pre-creation completed" -LogLevel "Info" -Details "3 users created successfully" -Source "New-TargetUsers"

# Simulate a warning
Add-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Message "Large file detected during migration" -LogLevel "Warning" -Details "File size: 2.1GB, User: john.doe@contoso.com" -Source "Copy-UserData"

# Example 13: Get log entries
$allLogs = Get-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId
Write-Host "Project has $($allLogs.Count) log entries:"
$allLogs | Format-Table LogId, LogLevel, Message, Timestamp, Source

# Get only error and warning logs
$importantLogs = Get-MigrationLog -DatabasePath $DatabasePath -ProjectId $project.ProjectId -LogLevel "Warning"
Write-Host "Warning logs: $($importantLogs.Count)"

# Example 14: Get migration progress summary
$projectSummary = Get-MigrationProject -DatabasePath $DatabasePath -ProjectId $project.ProjectId
$userSummary = Get-MigrationUser -DatabasePath $DatabasePath -ProjectId $project.ProjectId

Write-Host "=== MIGRATION PROGRESS SUMMARY ==="
Write-Host "Project: $($projectSummary.ProjectName)"
Write-Host "Status: $($projectSummary.Status)"
Write-Host "Total Users: $($userSummary.Count)"
Write-Host "Completed: $(($userSummary | Where-Object {$_.Status -eq 'Completed'}).Count)"
Write-Host "In Progress: $(($userSummary | Where-Object {$_.Status -eq 'InProgress'}).Count)"
Write-Host "Pending: $(($userSummary | Where-Object {$_.Status -eq 'Pending'}).Count)"
Write-Host "Failed: $(($userSummary | Where-Object {$_.Status -eq 'Failed'}).Count)"

$totalDataSize = ($userSummary | Measure-Object -Property DataSizeGB -Sum).Sum
Write-Host "Total Data Size: $totalDataSize GB"

# Example 15: Export data to CSV for reporting
$userSummary | Export-Csv -Path "C:\Migration\Reports\UserMigrationStatus.csv" -NoTypeInformation
Write-Host "User migration status exported to CSV"

# ================================
# CLEANUP EXAMPLES (USE WITH CAUTION)
# ================================

# Example 16: Remove a specific project (uncomment to use)
# Remove-MigrationProject -DatabasePath $DatabasePath -ProjectId $project.ProjectId -Force

# Example 17: Remove the entire database (uncomment to use)
# Remove-MigrationDatabase -DatabasePath $DatabasePath -Force

Write-Host "=== Examples completed successfully! ==="