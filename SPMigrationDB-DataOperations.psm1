#
# SPMigrationDB.DataOperations Module
# PowerShell Module for SQLite Data Operations (CRUD on Migration Data)
# Compatible with PowerShell 5.x
# Author: SharePoint Migration Team
# Version: 1.0
#

# Import required assemblies for SQLite
Add-Type -Path "$PSScriptRoot\System.Data.SQLite.dll" -ErrorAction SilentlyContinue

#region Private Functions

function Get-SQLiteConnectionString {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )
    
    $builder = New-Object System.Data.SQLite.SQLiteConnectionStringBuilder
    $builder.DataSource = $DatabasePath
    $builder.Version = 3
    $builder.JournalMode = "WAL"
    $builder.Pooling = $true
    $builder.FailIfMissing = $true
    
    return $builder.ConnectionString
}

function Invoke-SQLiteQuery {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        [Parameter(Mandatory = $true)]
        [string]$Query,
        [hashtable]$Parameters = @{}
    )
    
    try {
        $connectionString = Get-SQLiteConnectionString -DatabasePath $DatabasePath
        $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
        $connection.Open()
        
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        
        foreach ($param in $Parameters.GetEnumerator()) {
            $null = $command.Parameters.AddWithValue("@$($param.Key)", $param.Value)
        }
        
        $adapter = New-Object System.Data.SQLite.SQLiteDataAdapter($command)
        $dataset = New-Object System.Data.DataSet
        $null = $adapter.Fill($dataset)
        
        $connection.Close()
        
        return $dataset.Tables[0].Rows
    }
    catch {
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
        }
        throw
    }
}

function Invoke-SQLiteNonQuery {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        [Parameter(Mandatory = $true)]
        [string]$Query,
        [hashtable]$Parameters = @{}
    )
    
    try {
        $connectionString = Get-SQLiteConnectionString -DatabasePath $DatabasePath
        $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
        $connection.Open()
        
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        
        foreach ($param in $Parameters.GetEnumerator()) {
            $null = $command.Parameters.AddWithValue("@$($param.Key)", $param.Value)
        }
        
        $result = $command.ExecuteNonQuery()
        $connection.Close()
        
        return $result
    }
    catch {
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
        }
        throw
    }
}

function Invoke-SQLiteScalar {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        [Parameter(Mandatory = $true)]
        [string]$Query,
        [hashtable]$Parameters = @{}
    )
    
    try {
        $connectionString = Get-SQLiteConnectionString -DatabasePath $DatabasePath
        $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
        $connection.Open()
        
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        
        foreach ($param in $Parameters.GetEnumerator()) {
            $null = $command.Parameters.AddWithValue("@$($param.Key)", $param.Value)
        }
        
        $result = $command.ExecuteScalar()
        $connection.Close()
        
        return $result
    }
    catch {
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
        }
        throw
    }
}

function ConvertTo-PSObject {
    param(
        [Parameter(Mandatory = $true)]
        $DataRow
    )
    
    $obj = New-Object PSObject
    
    foreach ($column in $DataRow.Table.Columns) {
        $value = $DataRow[$column.ColumnName]
        if ($value -is [System.DBNull]) {
            $value = $null
        }
        $obj | Add-Member -MemberType NoteProperty -Name $column.ColumnName -Value $value
    }
    
    return $obj
}

#endregion

#region Migration Projects Functions

<#
.SYNOPSIS
Creates a new migration project.

.DESCRIPTION
Creates a new migration project in the database to track SharePoint/OneDrive migration activities.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectName
The name of the migration project.

.PARAMETER SourceTenant
The source tenant URL or name.

.PARAMETER TargetTenant
The target tenant URL or name.

.PARAMETER CreatedBy
The user creating the project.

.PARAMETER Description
Optional description of the migration project.

.EXAMPLE
New-MigrationProject -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectName "Contoso to Fabrikam Migration" -SourceTenant "contoso.onmicrosoft.com" -TargetTenant "fabrikam.onmicrosoft.com" -CreatedBy "admin@contoso.com"
#>
function New-MigrationProject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [Parameter(Mandatory = $true)]
        [string]$ProjectName,
        
        [Parameter(Mandatory = $true)]
        [string]$SourceTenant,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetTenant,
        
        [Parameter(Mandatory = $true)]
        [string]$CreatedBy,
        
        [string]$Description = ""
    )
    
    $query = @"
INSERT INTO MigrationProjects (ProjectName, SourceTenant, TargetTenant, CreatedBy, Description)
VALUES (@ProjectName, @SourceTenant, @TargetTenant, @CreatedBy, @Description);
SELECT last_insert_rowid();
"@
    
    $parameters = @{
        ProjectName = $ProjectName
        SourceTenant = $SourceTenant
        TargetTenant = $TargetTenant
        CreatedBy = $CreatedBy
        Description = $Description
    }
    
    try {
        $projectId = Invoke-SQLiteScalar -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        Write-Verbose "Migration project created with ID: $projectId"
        
        return Get-MigrationProject -DatabasePath $DatabasePath -ProjectId $projectId
    }
    catch {
        Write-Error "Failed to create migration project: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Gets migration projects from the database.

.DESCRIPTION
Retrieves migration project information from the database. Can get all projects or a specific project by ID.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectId
Optional. The ID of a specific project to retrieve.

.EXAMPLE
Get-MigrationProject -DatabasePath "C:\Migration\MigrationTracking.db"

.EXAMPLE
Get-MigrationProject -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1
#>
function Get-MigrationProject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [int]$ProjectId = 0
    )
    
    if ($ProjectId -gt 0) {
        $query = "SELECT * FROM MigrationProjects WHERE ProjectId = @ProjectId"
        $parameters = @{ ProjectId = $ProjectId }
    }
    else {
        $query = "SELECT * FROM MigrationProjects ORDER BY CreatedDate DESC"
        $parameters = @{}
    }
    
    try {
        $results = Invoke-SQLiteQuery -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        $projects = @()
        foreach ($row in $results) {
            $projects += ConvertTo-PSObject -DataRow $row
        }
        
        return $projects
    }
    catch {
        Write-Error "Failed to get migration projects: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Updates a migration project.

.DESCRIPTION
Updates the properties of an existing migration project.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectId
The ID of the project to update.

.PARAMETER Status
The new status of the project.

.PARAMETER Description
The new description of the project.

.EXAMPLE
Set-MigrationProject -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1 -Status "InProgress"
#>
function Set-MigrationProject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [Parameter(Mandatory = $true)]
        [int]$ProjectId,
        
        [ValidateSet('Planning', 'InProgress', 'Completed', 'Failed', 'Cancelled')]
        [string]$Status,
        
        [string]$Description
    )
    
    $updateFields = @()
    $parameters = @{ ProjectId = $ProjectId }
    
    if ($Status) {
        $updateFields += "Status = @Status"
        $parameters.Status = $Status
    }
    
    if ($Description) {
        $updateFields += "Description = @Description"
        $parameters.Description = $Description
    }
    
    if ($updateFields.Count -eq 0) {
        Write-Warning "No fields specified for update"
        return
    }
    
    $updateFields += "ModifiedDate = CURRENT_TIMESTAMP"
    
    $query = "UPDATE MigrationProjects SET $($updateFields -join ', ') WHERE ProjectId = @ProjectId"
    
    try {
        $result = Invoke-SQLiteNonQuery -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        if ($result -eq 1) {
            Write-Verbose "Migration project updated successfully"
            return Get-MigrationProject -DatabasePath $DatabasePath -ProjectId $ProjectId
        }
        else {
            Write-Warning "No project found with ID: $ProjectId"
        }
    }
    catch {
        Write-Error "Failed to update migration project: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Removes a migration project.

.DESCRIPTION
Deletes a migration project and all associated data from the database.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectId
The ID of the project to remove.

.PARAMETER Force
If specified, suppresses confirmation prompts.

.EXAMPLE
Remove-MigrationProject -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1 -Force
#>
function Remove-MigrationProject {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [Parameter(Mandatory = $true)]
        [int]$ProjectId,
        
        [switch]$Force
    )
    
    $project = Get-MigrationProject -DatabasePath $DatabasePath -ProjectId $ProjectId
    
    if (-not $project) {
        Write-Warning "Project with ID $ProjectId not found"
        return
    }
    
    if ($Force -or $PSCmdlet.ShouldProcess("Project: $($project.ProjectName)", "Remove migration project and all associated data")) {
        try {
            # Delete in order due to foreign key constraints
            $queries = @(
                "DELETE FROM MigrationLogs WHERE ProjectId = @ProjectId",
                "DELETE FROM MigrationActivities WHERE ProjectId = @ProjectId",
                "DELETE FROM MigrationUsers WHERE ProjectId = @ProjectId",
                "DELETE FROM MigrationSites WHERE ProjectId = @ProjectId",
                "DELETE FROM TenantRelationships WHERE ProjectId = @ProjectId",
                "DELETE FROM IdentityMapping WHERE ProjectId = @ProjectId",
                "DELETE FROM MigrationProjects WHERE ProjectId = @ProjectId"
            )
            
            $parameters = @{ ProjectId = $ProjectId }
            
            foreach ($query in $queries) {
                $null = Invoke-SQLiteNonQuery -DatabasePath $DatabasePath -Query $query -Parameters $parameters
            }
            
            Write-Verbose "Migration project and associated data removed successfully"
        }
        catch {
            Write-Error "Failed to remove migration project: $($_.Exception.Message)"
            throw
        }
    }
}

#endregion

#region Migration Users Functions

<#
.SYNOPSIS
Adds a user to a migration project.

.DESCRIPTION
Adds a user to be migrated in the specified migration project.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectId
The ID of the migration project.

.PARAMETER SourceUserPrincipalName
The source user's UPN.

.PARAMETER TargetUserPrincipalName
The target user's UPN.

.PARAMETER SourceOneDriveUrl
Optional. The source OneDrive URL.

.PARAMETER TargetOneDriveUrl
Optional. The target OneDrive URL.

.PARAMETER DataSizeGB
Optional. The size of data to migrate in GB.

.PARAMETER ItemCount
Optional. The number of items to migrate.

.EXAMPLE
Add-MigrationUser -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1 -SourceUserPrincipalName "user@contoso.com" -TargetUserPrincipalName "user@fabrikam.com"
#>
function Add-MigrationUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [Parameter(Mandatory = $true)]
        [int]$ProjectId,
        
        [Parameter(Mandatory = $true)]
        [string]$SourceUserPrincipalName,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetUserPrincipalName,
        
        [string]$SourceOneDriveUrl,
        
        [string]$TargetOneDriveUrl,
        
        [double]$DataSizeGB,
        
        [int]$ItemCount
    )
    
    $query = @"
INSERT INTO MigrationUsers (ProjectId, SourceUserPrincipalName, TargetUserPrincipalName, SourceOneDriveUrl, TargetOneDriveUrl, DataSizeGB, ItemCount)
VALUES (@ProjectId, @SourceUserPrincipalName, @TargetUserPrincipalName, @SourceOneDriveUrl, @TargetOneDriveUrl, @DataSizeGB, @ItemCount);
SELECT last_insert_rowid();
"@
    
    $parameters = @{
        ProjectId = $ProjectId
        SourceUserPrincipalName = $SourceUserPrincipalName
        TargetUserPrincipalName = $TargetUserPrincipalName
        SourceOneDriveUrl = $SourceOneDriveUrl
        TargetOneDriveUrl = $TargetOneDriveUrl
        DataSizeGB = $DataSizeGB
        ItemCount = $ItemCount
    }
    
    try {
        $userId = Invoke-SQLiteScalar -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        Write-Verbose "Migration user added with ID: $userId"
        
        return Get-MigrationUser -DatabasePath $DatabasePath -UserId $userId
    }
    catch {
        Write-Error "Failed to add migration user: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Gets migration users from the database.

.DESCRIPTION
Retrieves migration user information from the database. Can get all users for a project or a specific user by ID.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectId
Optional. The ID of the project to get users for.

.PARAMETER UserId
Optional. The ID of a specific user to retrieve.

.EXAMPLE
Get-MigrationUser -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1

.EXAMPLE
Get-MigrationUser -DatabasePath "C:\Migration\MigrationTracking.db" -UserId 1
#>
function Get-MigrationUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [int]$ProjectId = 0,
        
        [int]$UserId = 0
    )
    
    if ($UserId -gt 0) {
        $query = "SELECT * FROM MigrationUsers WHERE UserId = @UserId"
        $parameters = @{ UserId = $UserId }
    }
    elseif ($ProjectId -gt 0) {
        $query = "SELECT * FROM MigrationUsers WHERE ProjectId = @ProjectId ORDER BY CreatedDate"
        $parameters = @{ ProjectId = $ProjectId }
    }
    else {
        $query = "SELECT * FROM MigrationUsers ORDER BY CreatedDate"
        $parameters = @{}
    }
    
    try {
        $results = Invoke-SQLiteQuery -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        $users = @()
        foreach ($row in $results) {
            $users += ConvertTo-PSObject -DataRow $row
        }
        
        return $users
    }
    catch {
        Write-Error "Failed to get migration users: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Updates a migration user.

.DESCRIPTION
Updates the properties of an existing migration user.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER UserId
The ID of the user to update.

.PARAMETER Status
The new status of the user migration.

.PARAMETER DataSizeGB
The size of data migrated in GB.

.PARAMETER ItemCount
The number of items migrated.

.PARAMETER StartTime
The start time of the migration.

.PARAMETER EndTime
The end time of the migration.

.PARAMETER ErrorMessage
Any error message from the migration.

.EXAMPLE
Set-MigrationUser -DatabasePath "C:\Migration\MigrationTracking.db" -UserId 1 -Status "InProgress" -StartTime (Get-Date)
#>
function Set-MigrationUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [Parameter(Mandatory = $true)]
        [int]$UserId,
        
        [ValidateSet('Pending', 'InProgress', 'Completed', 'Failed', 'Skipped')]
        [string]$Status,
        
        [double]$DataSizeGB,
        
        [int]$ItemCount,
        
        [datetime]$StartTime,
        
        [datetime]$EndTime,
        
        [string]$ErrorMessage
    )
    
    $updateFields = @()
    $parameters = @{ UserId = $UserId }
    
    if ($Status) {
        $updateFields += "Status = @Status"
        $parameters.Status = $Status
    }
    
    if ($DataSizeGB) {
        $updateFields += "DataSizeGB = @DataSizeGB"
        $parameters.DataSizeGB = $DataSizeGB
    }
    
    if ($ItemCount) {
        $updateFields += "ItemCount = @ItemCount"
        $parameters.ItemCount = $ItemCount
    }
    
    if ($StartTime) {
        $updateFields += "StartTime = @StartTime"
        $parameters.StartTime = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    if ($EndTime) {
        $updateFields += "EndTime = @EndTime"
        $parameters.EndTime = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    if ($ErrorMessage) {
        $updateFields += "ErrorMessage = @ErrorMessage"
        $parameters.ErrorMessage = $ErrorMessage
    }
    
    if ($updateFields.Count -eq 0) {
        Write-Warning "No fields specified for update"
        return
    }
    
    $updateFields += "ModifiedDate = CURRENT_TIMESTAMP"
    
    $query = "UPDATE MigrationUsers SET $($updateFields -join ', ') WHERE UserId = @UserId"
    
    try {
        $result = Invoke-SQLiteNonQuery -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        if ($result -eq 1) {
            Write-Verbose "Migration user updated successfully"
            return Get-MigrationUser -DatabasePath $DatabasePath -UserId $UserId
        }
        else {
            Write-Warning "No user found with ID: $UserId"
        }
    }
    catch {
        Write-Error "Failed to update migration user: $($_.Exception.Message)"
        throw
    }
}

#endregion

#region Migration Logs Functions

<#
.SYNOPSIS
Adds a log entry to the migration database.

.DESCRIPTION
Adds a detailed log entry for migration tracking and troubleshooting.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectId
The ID of the migration project.

.PARAMETER Message
The log message.

.PARAMETER LogLevel
The severity level of the log entry.

.PARAMETER ActivityId
Optional. The ID of the associated activity.

.PARAMETER Details
Optional. Additional details for the log entry.

.PARAMETER Source
Optional. The source of the log entry (script name, function, etc.).

.EXAMPLE
Add-MigrationLog -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1 -Message "User migration started" -LogLevel "Info" -Source "Start-UserMigration"
#>
function Add-MigrationLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [Parameter(Mandatory = $true)]
        [int]$ProjectId,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [ValidateSet('Info', 'Warning', 'Error', 'Debug')]
        [string]$LogLevel = 'Info',
        
        [int]$ActivityId,
        
        [string]$Details,
        
        [string]$Source
    )
    
    $query = @"
INSERT INTO MigrationLogs (ProjectId, ActivityId, LogLevel, Message, Details, Source)
VALUES (@ProjectId, @ActivityId, @LogLevel, @Message, @Details, @Source);
SELECT last_insert_rowid();
"@
    
    $parameters = @{
        ProjectId = $ProjectId
        ActivityId = $ActivityId
        LogLevel = $LogLevel
        Message = $Message
        Details = $Details
        Source = $Source
    }
    
    try {
        $logId = Invoke-SQLiteScalar -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        Write-Verbose "Migration log entry added with ID: $logId"
        
        return $logId
    }
    catch {
        Write-Error "Failed to add migration log: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Gets migration log entries from the database.

.DESCRIPTION
Retrieves migration log entries from the database with optional filtering.

.PARAMETER DatabasePath
The path to the SQLite database file.

.PARAMETER ProjectId
Optional. The ID of the project to get logs for.

.PARAMETER LogLevel
Optional. Filter by log level.

.PARAMETER StartDate
Optional. Filter logs from this date.

.PARAMETER EndDate
Optional. Filter logs to this date.

.PARAMETER Limit
Optional. Maximum number of log entries to return (default: 1000).

.EXAMPLE
Get-MigrationLog -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1 -LogLevel "Error"

.EXAMPLE
Get-MigrationLog -DatabasePath "C:\Migration\MigrationTracking.db" -ProjectId 1 -StartDate (Get-Date).AddDays(-7) -Limit 500
#>
function Get-MigrationLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [int]$ProjectId = 0,
        
        [ValidateSet('Info', 'Warning', 'Error', 'Debug')]
        [string]$LogLevel,
        
        [datetime]$StartDate,
        
        [datetime]$EndDate,
        
        [int]$Limit = 1000
    )
    
    $whereConditions = @()
    $parameters = @{}
    
    if ($ProjectId -gt 0) {
        $whereConditions += "ProjectId = @ProjectId"
        $parameters.ProjectId = $ProjectId
    }
    
    if ($LogLevel) {
        $whereConditions += "LogLevel = @LogLevel"
        $parameters.LogLevel = $LogLevel
    }
    
    if ($StartDate) {
        $whereConditions += "Timestamp >= @StartDate"
        $parameters.StartDate = $StartDate.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    if ($EndDate) {
        $whereConditions += "Timestamp <= @EndDate"
        $parameters.EndDate = $EndDate.ToString("yyyy-MM-dd HH:mm:ss")
    }
    
    $query = "SELECT * FROM MigrationLogs"
    
    if ($whereConditions.Count -gt 0) {
        $query += " WHERE " + ($whereConditions -join " AND ")
    }
    
    $query += " ORDER BY Timestamp DESC LIMIT @Limit"
    $parameters.Limit = $Limit
    
    try {
        $results = Invoke-SQLiteQuery -DatabasePath $DatabasePath -Query $query -Parameters $parameters
        
        $logs = @()
        foreach ($row in $results) {
            $logs += ConvertTo-PSObject -DataRow $row
        }
        
        return $logs
    }
    catch {
        Write-Error "Failed to get migration logs: $($_.Exception.Message)"
        throw
    }
}

#endregion

# Export public functions
Export-ModuleMember -Function @(
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