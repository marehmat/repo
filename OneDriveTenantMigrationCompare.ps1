# ================================
# CONFIGURATION
# ================================
$SourceTenantAdminUrl = "https://sourcetenant-admin.sharepoint.com"
$DestTenantAdminUrl = "https://desttenant-admin.sharepoint.com"

$SourceServiceAccount = "sourceadmin@sourcetenant.com"
$DestServiceAccount = "destadmin@desttenant.com"

$SourceFolderPath = "Documents/TargetFolder"  # Specific folder in source

$InputCsv = "C:\Path\To\users.csv"  # CSV with SourceUPN and DestUPN columns
$OutputCsv = "C:\Path\To\OneDriveMigrationValidation.csv"

# ================================
# INITIALIZE OUTPUT
# ================================
$ValidationResults = @()

# ================================
# FUNCTION: GET FILES FROM ONEDRIVE
# ================================
function Get-OneDriveFiles {
    param(
        [string]$SiteUrl,
        [string]$FolderPath,  # Empty for entire OneDrive
        [bool]$IsEntireOneDrive = $false
    )
    
    try {
        if ($IsEntireOneDrive) {
            Write-Host "    Getting ALL files from entire OneDrive..." -ForegroundColor Gray
            # Get root Documents folder
            $DocumentsFolder = Get-PnPFolder -Url "Documents" -ErrorAction Stop
            $AllFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $DocumentsFolder.ServerRelativeUrl -ItemType File -Recursive -ErrorAction Stop
        }
        else {
            Write-Host "    Getting files from folder: $FolderPath..." -ForegroundColor Gray
            $Folder = Get-PnPFolder -Url $FolderPath -ErrorAction Stop
            $AllFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $Folder.ServerRelativeUrl -ItemType File -Recursive -ErrorAction Stop
        }
        
        return $AllFiles
    }
    catch {
        Write-Warning "    Failed to retrieve files: $_"
        return $null
    }
}

# ================================
# FUNCTION: CALCULATE FILE STATISTICS
# ================================
function Get-FileStatistics {
    param(
        [array]$Files
    )
    
    if ($null -eq $Files -or $Files.Count -eq 0) {
        return @{
            TotalFiles = 0
            TotalSizeMB = 0
            LastModified = "N/A"
            FileList = @()
        }
    }
    
    $TotalSize = ($Files | Measure-Object -Property Length -Sum).Sum
    if ($null -eq $TotalSize) { $TotalSize = 0 }
    
    $SizeMB = [math]::Round($TotalSize / 1MB, 2)
    $LastModified = ($Files | Sort-Object -Property TimeLastModified -Descending | Select-Object -First 1).TimeLastModified
    
    # Create simplified file list for comparison
    $FileList = $Files | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name
            Size = $_.Length
            Modified = $_.TimeLastModified
            RelativePath = $_.ServerRelativeUrl
        }
    }
    
    return @{
        TotalFiles = $Files.Count
        TotalSizeMB = $SizeMB
        LastModified = $LastModified
        FileList = $FileList
    }
}

# ================================
# FUNCTION: COMPARE SOURCE AND DESTINATION
# ================================
function Compare-MigrationData {
    param(
        [array]$SourceFiles,
        [array]$DestFiles
    )
    
    $SourceFileNames = $SourceFiles | Select-Object -ExpandProperty Name
    $DestFileNames = $DestFiles | Select-Object -ExpandProperty Name
    
    # Files only in source (not migrated)
    $MissingInDest = @($SourceFileNames | Where-Object { $_ -notin $DestFileNames })
    
    # Files only in destination (extra files)
    $ExtraInDest = @($DestFileNames | Where-Object { $_ -notin $SourceFileNames })
    
    # Files in both - check for newer versions in destination
    $CommonFiles = $SourceFileNames | Where-Object { $_ -in $DestFileNames }
    $NewerInDest = 0
    
    foreach ($FileName in $CommonFiles) {
        $SourceFile = $SourceFiles | Where-Object { $_.Name -eq $FileName }
        $DestFile = $DestFiles | Where-Object { $_.Name -eq $FileName }
        
        if ($DestFile.Modified -gt $SourceFile.Modified) {
            $NewerInDest++
        }
    }
    
    return @{
        MissingInDestCount = $MissingInDest.Count
        MissingInDestFiles = ($MissingInDest -join "; ")
        ExtraInDestCount = $ExtraInDest.Count
        ExtraInDestFiles = ($ExtraInDest -join "; ")
        NewerInDestCount = $NewerInDest
        CommonFilesCount = $CommonFiles.Count
    }
}

# ================================
# CONNECT TO SOURCE ADMIN
# ================================
Write-Host "`n======================================" -ForegroundColor Cyan
Write-Host "ONEDRIVE MIGRATION VALIDATION" -ForegroundColor Cyan
Write-Host "======================================`n" -ForegroundColor Cyan

Write-Host "Connecting to SOURCE tenant admin..." -ForegroundColor Cyan
Connect-SPOService -Url $SourceTenantAdminUrl

# ================================
# PROCESS EACH USER
# ================================
$Users = Import-Csv -Path $InputCsv
$TotalUsers = $Users.Count
$CurrentUser = 0

foreach ($User in $Users) {
    $CurrentUser++
    $SourceUPN = $User.SourceUPN
    $DestUPN = $User.DestUPN
    
    Write-Host "`n[$CurrentUser/$TotalUsers] Processing Migration Mapping" -ForegroundColor Yellow
    Write-Host "  Source: $SourceUPN" -ForegroundColor Yellow
    Write-Host "  Destination: $DestUPN" -ForegroundColor Yellow
    Write-Host "================================================" -ForegroundColor Yellow
    
    # Format OneDrive URLs for each tenant with respective UPNs
    $SourceUPNFormatted = $SourceUPN -replace '[@\.]', '_'
    $DestUPNFormatted = $DestUPN -replace '[@\.]', '_'
    
    $SourceSiteUrl = "https://sourcetenant-my.sharepoint.com/personal/$SourceUPNFormatted"
    $DestSiteUrl = "https://desttenant-my.sharepoint.com/personal/$DestUPNFormatted"
    
    $SourceStats = $null
    $DestStats = $null
    $ComparisonResults = $null
    
    try {
        # ========================================
        # PROCESS SOURCE TENANT
        # ========================================
        Write-Host "  [SOURCE] Processing..." -ForegroundColor Cyan
        
        try {
            Set-SPOUser -Site $SourceSiteUrl -LoginName $SourceServiceAccount -IsSiteCollectionAdmin $true -ErrorAction Stop
            Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin -ErrorAction Stop
            
            $SourceFiles = Get-OneDriveFiles -SiteUrl $SourceSiteUrl -FolderPath $SourceFolderPath -IsEntireOneDrive $false
            $SourceStats = Get-FileStatistics -Files $SourceFiles
            
            Write-Host "    Files: $($SourceStats.TotalFiles) | Size: $($SourceStats.TotalSizeMB) MB" -ForegroundColor Green
        }
        finally {
            Set-SPOUser -Site $SourceSiteUrl -LoginName $SourceServiceAccount -IsSiteCollectionAdmin $false -ErrorAction SilentlyContinue
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        
        # ========================================
        # PROCESS DESTINATION TENANT
        # ========================================
        Write-Host "  [DESTINATION] Processing..." -ForegroundColor Cyan
        
        # Disconnect from source and connect to destination admin
        Disconnect-SPOService -ErrorAction SilentlyContinue
        Connect-SPOService -Url $DestTenantAdminUrl -ErrorAction Stop
        
        try {
            Set-SPOUser -Site $DestSiteUrl -LoginName $DestServiceAccount -IsSiteCollectionAdmin $true -ErrorAction Stop
            Connect-PnPOnline -Url $DestSiteUrl -UseWebLogin -ErrorAction Stop
            
            $DestFiles = Get-OneDriveFiles -SiteUrl $DestSiteUrl -FolderPath "" -IsEntireOneDrive $true
            $DestStats = Get-FileStatistics -Files $DestFiles
            
            Write-Host "    Files: $($DestStats.TotalFiles) | Size: $($DestStats.TotalSizeMB) MB" -ForegroundColor Green
        }
        finally {
            Set-SPOUser -Site $DestSiteUrl -LoginName $DestServiceAccount -IsSiteCollectionAdmin $false -ErrorAction SilentlyContinue
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        
        # Reconnect to source admin for next iteration
        Disconnect-SPOService -ErrorAction SilentlyContinue
        Connect-SPOService -Url $SourceTenantAdminUrl -ErrorAction SilentlyContinue
        
        # ========================================
        # COMPARISON
        # ========================================
        Write-Host "  [COMPARISON] Analyzing differences..." -ForegroundColor Cyan
        
        if ($SourceStats -and $DestStats) {
            $ComparisonResults = Compare-MigrationData -SourceFiles $SourceStats.FileList -DestFiles $DestStats.FileList
            
            Write-Host "    Files in both: $($ComparisonResults.CommonFilesCount)" -ForegroundColor White
            Write-Host "    Missing in destination: $($ComparisonResults.MissingInDestCount)" -ForegroundColor $(if ($ComparisonResults.MissingInDestCount -gt 0) { "Red" } else { "Green" })
            Write-Host "    Extra in destination: $($ComparisonResults.ExtraInDestCount)" -ForegroundColor White
            Write-Host "    Newer in destination: $($ComparisonResults.NewerInDestCount)" -ForegroundColor Magenta
        }
        
        # ========================================
        # STORE RESULTS
        # ========================================
        $ValidationResults += [PSCustomObject]@{
            SourceUPN = $SourceUPN
            DestUPN = $DestUPN
            'Source_FileCount' = if ($SourceStats) { $SourceStats.TotalFiles } else { "Error" }
            'Source_SizeMB' = if ($SourceStats) { $SourceStats.TotalSizeMB } else { "Error" }
            'Source_LastModified' = if ($SourceStats) { $SourceStats.LastModified } else { "Error" }
            'Dest_FileCount' = if ($DestStats) { $DestStats.TotalFiles } else { "Error" }
            'Dest_SizeMB' = if ($DestStats) { $DestStats.TotalSizeMB } else { "Error" }
            'Dest_LastModified' = if ($DestStats) { $DestStats.LastModified } else { "Error" }
            'FilesInBoth' = if ($ComparisonResults) { $ComparisonResults.CommonFilesCount } else { "N/A" }
            'MissingInDest_Count' = if ($ComparisonResults) { $ComparisonResults.MissingInDestCount } else { "N/A" }
            'MissingInDest_Files' = if ($ComparisonResults) { $ComparisonResults.MissingInDestFiles } else { "N/A" }
            'ExtraInDest_Count' = if ($ComparisonResults) { $ComparisonResults.ExtraInDestCount } else { "N/A" }
            'ExtraInDest_Files' = if ($ComparisonResults) { $ComparisonResults.ExtraInDestFiles } else { "N/A" }
            'NewerInDest_Count' = if ($ComparisonResults) { $ComparisonResults.NewerInDestCount } else { "N/A" }
            'ValidationStatus' = if ($ComparisonResults.MissingInDestCount -eq 0) { "PASS" } else { "FAIL" }
        }
        
        Write-Host "  [RESULT] Validation: $($ValidationResults[-1].ValidationStatus)" -ForegroundColor $(if ($ValidationResults[-1].ValidationStatus -eq "PASS") { "Green" } else { "Red" })
        
    }
    catch {
        Write-Warning "Critical error processing $SourceUPN -> $DestUPN: $_"
        
        $ValidationResults += [PSCustomObject]@{
            SourceUPN = $SourceUPN
            DestUPN = $DestUPN
            'Source_FileCount' = "Error"
            'Source_SizeMB' = "Error"
            'Source_LastModified' = "Error"
            'Dest_FileCount' = "Error"
            'Dest_SizeMB' = "Error"
            'Dest_LastModified' = "Error"
            'FilesInBoth' = "Error"
            'MissingInDest_Count' = "Error"
            'MissingInDest_Files' = "Error"
            'ExtraInDest_Count' = "Error"
            'ExtraInDest_Files' = "Error"
            'NewerInDest_Count' = "Error"
            'ValidationStatus' = "ERROR"
        }
    }
}

# ================================
# EXPORT RESULTS
# ================================
Write-Host "`n======================================" -ForegroundColor Cyan
Write-Host "VALIDATION COMPLETE" -ForegroundColor Cyan
Write-Host "======================================`n" -ForegroundColor Cyan

$ValidationResults | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "Report saved to: $OutputCsv" -ForegroundColor Green

# ================================
# SUMMARY STATISTICS
# ================================
$PassCount = ($ValidationResults | Where-Object { $_.ValidationStatus -eq "PASS" }).Count
$FailCount = ($ValidationResults | Where-Object { $_.ValidationStatus -eq "FAIL" }).Count
$ErrorCount = ($ValidationResults | Where-Object { $_.ValidationStatus -eq "ERROR" }).Count

Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  Total User Mappings: $TotalUsers" -ForegroundColor White
Write-Host "  Passed: $PassCount" -ForegroundColor Green
Write-Host "  Failed: $FailCount" -ForegroundColor Red
Write-Host "  Errors: $ErrorCount" -ForegroundColor Yellow
Write-Host "`nMigration validation complete!" -ForegroundColor Cyan
