#
# SPMigrationDB.DatabaseAdmin Module
# PowerShell Module for SQLite Database Administration (Table Management)
# Compatible with PowerShell 5.x
# Author: SharePoint Migration Team
# Version: 1.0
#

# Import required assemblies for SQLite
Add-Type -Path "$PSScriptRoot\System.Data.SQLite.dll" -ErrorAction SilentlyContinue

#region Private Functions

function Test-SQLiteAssembly {
    try {
        $null = [System.Data.SQLite.SQLiteConnection]::new()
        return $true
    }
    catch {
        return $false
    }
}

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
    $builder.FailIfMissing = $false
    
    return $builder.ConnectionString
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

#endregion

#region Public Functions

<#
.SYNOPSIS
Creates a new SQLite database for migration tracking.

.DESCRIPTION
Creates a new SQLite database with the predefined schema for SharePoint/OneDrive migration tracking.

.PARAMETER DatabasePath
The full path where the SQLite database file should be created.

.PARAMETER Force
If specified, overwrites an existing database file.

.EXAMPLE
New-MigrationDatabase -DatabasePath "C:\Migration\MigrationTracking.db"

.EXAMPLE
New-MigrationDatabase -DatabasePath "C:\Migration\MigrationTracking.db" -Force
#>
function New-MigrationDatabase {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [switch]$Force
    )
    
    if (-not (Test-SQLiteAssembly)) {
        throw "SQLite assembly not found. Please ensure System.Data.SQLite.dll is in the module directory."
    }
    
    if (Test-Path $DatabasePath) {
        if ($Force) {
            if ($PSCmdlet.ShouldProcess($DatabasePath, "Remove existing database")) {
                Remove-Item $DatabasePath -Force
            }
        } else {
            throw "Database file already exists: $DatabasePath. Use -Force to overwrite."
        }
    }
    
    $directory = Split-Path $DatabasePath -Parent
    if ($directory -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force
    }
    
    try {
        if ($PSCmdlet.ShouldProcess($DatabasePath, "Create migration database")) {
            # Create the database and tables
            $schema = Get-MigrationDatabaseSchema
            
            $null = Invoke-SQLiteNonQuery -DatabasePath $DatabasePath -Query $schema
            
            Write-Verbose "Migration database created successfully at: $DatabasePath"
            
            return Get-Item $DatabasePath
        }
    }
    catch {
        Write-Error "Failed to create migration database: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Gets the database schema SQL for migration tracking.

.DESCRIPTION
Returns the complete SQL schema for creating all tables required for migration tracking.

.EXAMPLE
$schema = Get-MigrationDatabaseSchema
#>
function Get-MigrationDatabaseSchema {
    [CmdletBinding()]
    param()
    
    $schema = @"
-- Migration tracking database schema for SharePoint/OneDrive tenant migration

-- Main migration projects table
CREATE TABLE IF NOT EXISTS MigrationProjects (
    ProjectId INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectName TEXT NOT NULL,
    SourceTenant TEXT NOT NULL,
    TargetTenant TEXT NOT NULL,
    CreatedBy TEXT NOT NULL,
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    Status TEXT CHECK(Status IN ('Planning', 'InProgress', 'Completed', 'Failed', 'Cancelled')) DEFAULT 'Planning',
    Description TEXT
);

-- Users to be migrated
CREATE TABLE IF NOT EXISTS MigrationUsers (
    UserId INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectId INTEGER NOT NULL,
    SourceUserPrincipalName TEXT NOT NULL,
    TargetUserPrincipalName TEXT NOT NULL,
    SourceOneDriveUrl TEXT,
    TargetOneDriveUrl TEXT,
    Status TEXT CHECK(Status IN ('Pending', 'InProgress', 'Completed', 'Failed', 'Skipped')) DEFAULT 'Pending',
    DataSizeGB REAL,
    ItemCount INTEGER,
    StartTime DATETIME,
    EndTime DATETIME,
    ErrorMessage TEXT,
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ProjectId) REFERENCES MigrationProjects(ProjectId)
);

-- SharePoint sites to be migrated
CREATE TABLE IF NOT EXISTS MigrationSites (
    SiteId INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectId INTEGER NOT NULL,
    SourceSiteUrl TEXT NOT NULL,
    TargetSiteUrl TEXT NOT NULL,
    SiteTitle TEXT,
    Template TEXT,
    Status TEXT CHECK(Status IN ('Pending', 'InProgress', 'Completed', 'Failed', 'Skipped')) DEFAULT 'Pending',
    DataSizeGB REAL,
    ItemCount INTEGER,
    StartTime DATETIME,
    EndTime DATETIME,
    ErrorMessage TEXT,
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ProjectId) REFERENCES MigrationProjects(ProjectId)
);

-- Detailed migration activities/tasks
CREATE TABLE IF NOT EXISTS MigrationActivities (
    ActivityId INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectId INTEGER NOT NULL,
    UserId INTEGER,
    SiteId INTEGER,
    ActivityType TEXT NOT NULL,
    ActivityName TEXT NOT NULL,
    Status TEXT CHECK(Status IN ('Pending', 'InProgress', 'Completed', 'Failed', 'Skipped')) DEFAULT 'Pending',
    StartTime DATETIME,
    EndTime DATETIME,
    Duration INTEGER,
    Notes TEXT,
    ErrorMessage TEXT,
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ProjectId) REFERENCES MigrationProjects(ProjectId),
    FOREIGN KEY (UserId) REFERENCES MigrationUsers(UserId),
    FOREIGN KEY (SiteId) REFERENCES MigrationSites(SiteId)
);

-- Migration logs for detailed tracking
CREATE TABLE IF NOT EXISTS MigrationLogs (
    LogId INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectId INTEGER NOT NULL,
    ActivityId INTEGER,
    LogLevel TEXT CHECK(LogLevel IN ('Info', 'Warning', 'Error', 'Debug')) DEFAULT 'Info',
    Message TEXT NOT NULL,
    Details TEXT,
    Timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
    Source TEXT,
    FOREIGN KEY (ProjectId) REFERENCES MigrationProjects(ProjectId),
    FOREIGN KEY (ActivityId) REFERENCES MigrationActivities(ActivityId)
);

-- Cross-tenant relationship tracking
CREATE TABLE IF NOT EXISTS TenantRelationships (
    RelationshipId INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectId INTEGER NOT NULL,
    SourceTenant TEXT NOT NULL,
    TargetTenant TEXT NOT NULL,
    RelationshipStatus TEXT CHECK(RelationshipStatus IN ('Establishing', 'Active', 'Removing', 'Removed', 'Failed')) DEFAULT 'Establishing',
    TrustEstablishedDate DATETIME,
    TrustRemovedDate DATETIME,
    ErrorMessage TEXT,
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ProjectId) REFERENCES MigrationProjects(ProjectId)
);

-- Identity mapping for users and groups
CREATE TABLE IF NOT EXISTS IdentityMapping (
    MappingId INTEGER PRIMARY KEY AUTOINCREMENT,
    ProjectId INTEGER NOT NULL,
    SourceIdentity TEXT NOT NULL,
    TargetIdentity TEXT NOT NULL,
    IdentityType TEXT CHECK(IdentityType IN ('User', 'Group', 'SecurityGroup')) DEFAULT 'User',
    MappingStatus TEXT CHECK(MappingStatus IN ('Pending', 'Validated', 'Failed')) DEFAULT 'Pending',
    CreatedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ProjectId) REFERENCES MigrationProjects(ProjectId)
);

-- Create indexes for performance
CREATE INDEX IF NOT EXISTS idx_migrationusers_project ON MigrationUsers(ProjectId);
CREATE INDEX IF NOT EXISTS idx_migrationsites_project ON MigrationSites(ProjectId);
CREATE INDEX IF NOT EXISTS idx_migrationactivities_project ON MigrationActivities(ProjectId);
CREATE INDEX IF NOT EXISTS idx_migrationlogs_project ON MigrationLogs(ProjectId);
CREATE INDEX IF NOT EXISTS idx_migrationlogs_timestamp ON MigrationLogs(Timestamp);
CREATE INDEX IF NOT EXISTS idx_tenantrelationships_project ON TenantRelationships(ProjectId);
CREATE INDEX IF NOT EXISTS idx_identitymapping_project ON IdentityMapping(ProjectId);
"@
    
    return $schema
}

<#
.SYNOPSIS
Tests the connection to a migration database.

.DESCRIPTION
Verifies that the SQLite database exists and can be accessed.

.PARAMETER DatabasePath
The path to the SQLite database file.

.EXAMPLE
Test-MigrationDatabase -DatabasePath "C:\Migration\MigrationTracking.db"
#>
function Test-MigrationDatabase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )
    
    if (-not (Test-Path $DatabasePath)) {
        return $false
    }
    
    try {
        $result = Invoke-SQLiteScalar -DatabasePath $DatabasePath -Query "SELECT 1"
        return ($result -eq 1)
    }
    catch {
        Write-Warning "Database connection test failed: $($_.Exception.Message)"
        return $false
    }
}

<#
.SYNOPSIS
Gets information about tables in the migration database.

.DESCRIPTION
Retrieves metadata about all tables in the migration database, including table names and row counts.

.PARAMETER DatabasePath
The path to the SQLite database file.

.EXAMPLE
Get-MigrationDatabaseTables -DatabasePath "C:\Migration\MigrationTracking.db"
#>
function Get-MigrationDatabaseTables {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )
    
    if (-not (Test-MigrationDatabase -DatabasePath $DatabasePath)) {
        throw "Cannot access database: $DatabasePath"
    }
    
    try {
        $connectionString = Get-SQLiteConnectionString -DatabasePath $DatabasePath
        $connection = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
        $connection.Open()
        
        $tables = @()
        
        # Get table names
        $command = $connection.CreateCommand()
        $command.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
        $reader = $command.ExecuteReader()
        
        while ($reader.Read()) {
            $tableName = $reader["name"]
            $reader2 = $null
            
            try {
                # Get row count for each table
                $command2 = $connection.CreateCommand()
                $command2.CommandText = "SELECT COUNT(*) FROM [$tableName]"
                $rowCount = $command2.ExecuteScalar()
                
                $tables += [PSCustomObject]@{
                    TableName = $tableName
                    RowCount = [int]$rowCount
                }
            }
            catch {
                $tables += [PSCustomObject]@{
                    TableName = $tableName
                    RowCount = 0
                }
            }
        }
        
        $reader.Close()
        $connection.Close()
        
        return $tables
    }
    catch {
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
        }
        Write-Error "Failed to get database tables: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Backs up a migration database.

.DESCRIPTION
Creates a backup copy of the migration database.

.PARAMETER DatabasePath
The path to the source SQLite database file.

.PARAMETER BackupPath
The path where the backup should be created.

.PARAMETER Force
If specified, overwrites an existing backup file.

.EXAMPLE
Backup-MigrationDatabase -DatabasePath "C:\Migration\MigrationTracking.db" -BackupPath "C:\Migration\Backup\MigrationTracking_20231201.db"
#>
function Backup-MigrationDatabase {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        
        [switch]$Force
    )
    
    if (-not (Test-MigrationDatabase -DatabasePath $DatabasePath)) {
        throw "Source database not found or inaccessible: $DatabasePath"
    }
    
    if (Test-Path $BackupPath) {
        if ($Force) {
            if ($PSCmdlet.ShouldProcess($BackupPath, "Remove existing backup")) {
                Remove-Item $BackupPath -Force
            }
        } else {
            throw "Backup file already exists: $BackupPath. Use -Force to overwrite."
        }
    }
    
    $backupDirectory = Split-Path $BackupPath -Parent
    if ($backupDirectory -and -not (Test-Path $backupDirectory)) {
        New-Item -ItemType Directory -Path $backupDirectory -Force
    }
    
    try {
        if ($PSCmdlet.ShouldProcess($DatabasePath, "Create backup to $BackupPath")) {
            Copy-Item $DatabasePath $BackupPath -Force
            
            Write-Verbose "Database backup created successfully: $BackupPath"
            
            return Get-Item $BackupPath
        }
    }
    catch {
        Write-Error "Failed to create database backup: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
Removes a migration database.

.DESCRIPTION
Deletes the specified migration database file.

.PARAMETER DatabasePath
The path to the SQLite database file to remove.

.PARAMETER Force
If specified, suppresses confirmation prompts.

.EXAMPLE
Remove-MigrationDatabase -DatabasePath "C:\Migration\MigrationTracking.db" -Force
#>
function Remove-MigrationDatabase {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath,
        
        [switch]$Force
    )
    
    if (-not (Test-Path $DatabasePath)) {
        Write-Warning "Database file not found: $DatabasePath"
        return
    }
    
    if ($Force -or $PSCmdlet.ShouldProcess($DatabasePath, "Remove migration database")) {
        try {
            Remove-Item $DatabasePath -Force
            Write-Verbose "Migration database removed successfully: $DatabasePath"
        }
        catch {
            Write-Error "Failed to remove database: $($_.Exception.Message)"
            throw
        }
    }
}

#endregion

# Export public functions
Export-ModuleMember -Function @(
    'New-MigrationDatabase',
    'Get-MigrationDatabaseSchema',
    'Test-MigrationDatabase',
    'Get-MigrationDatabaseTables',
    'Backup-MigrationDatabase',
    'Remove-MigrationDatabase'
)