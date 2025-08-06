<#
.SYNOPSIS
    Generate SharePoint Site and Document Library Permissions Report
.DESCRIPTION
    This script extracts all permissions from a SharePoint site and identifies
    document libraries with unique permissions, exporting results to CSV
.PARAMETER SiteURL
    The SharePoint site URL to analyze
.PARAMETER ReportPath
    Path where the CSV report will be saved
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteURL,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = "C:\temp\SharePointPermissionsReport.csv"
)

# Global array to store all permission data
$Global:PermissionData = @()

# Function to get permissions for any SharePoint object
Function Get-SharePointPermissions {
    param(
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)][string]$ObjectType,
        [Parameter(Mandatory=$true)][string]$ObjectTitle,
        [Parameter(Mandatory=$true)][string]$ObjectURL
    )
    
    # Check if object has unique permissions
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments
    
    # Get all role assignments for the object
    $RoleAssignments = $Object.RoleAssignments
    
    foreach ($RoleAssignment in $RoleAssignments) {
        # Get the permission levels and member details
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
        
        # Get permission levels (exclude Limited Access)
        $PermissionLevels = ($RoleAssignment.RoleDefinitionBindings | 
                           Where-Object { $_.Name -ne "Limited Access" } | 
                           Select-Object -ExpandProperty Name) -join ", "
        
        # Skip if no meaningful permissions
        if ([string]::IsNullOrEmpty($PermissionLevels)) { continue }
        
        # Get member details
        $Member = $RoleAssignment.Member
        $PrincipalType = $Member.PrincipalType
        
        # Handle SharePoint Groups
        if ($PrincipalType -eq "SharePointGroup") {
            try {
                # Get group members
                $GroupMembers = Get-PnPGroupMember -Identity $Member.LoginName -ErrorAction SilentlyContinue
                
                if ($GroupMembers.Count -gt 0) {
                    foreach ($GroupMember in $GroupMembers) {
                        $Global:PermissionData += [PSCustomObject]@{
                            Object = $ObjectType
                            Title = $ObjectTitle
                            URL = $ObjectURL
                            HasUniquePermissions = $HasUniquePermissions
                            PrincipalType = "GroupMember"
                            PrincipalName = $GroupMember.Title
                            PrincipalLoginName = $GroupMember.LoginName
                            GroupName = $Member.Title
                            PermissionLevels = $PermissionLevels
                            GrantedThrough = "SharePoint Group: $($Member.Title)"
                        }
                    }
                } else {
                    # Empty group
                    $Global:PermissionData += [PSCustomObject]@{
                        Object = $ObjectType
                        Title = $ObjectTitle
                        URL = $ObjectURL
                        HasUniquePermissions = $HasUniquePermissions
                        PrincipalType = "SharePointGroup"
                        PrincipalName = $Member.Title
                        PrincipalLoginName = $Member.LoginName
                        GroupName = ""
                        PermissionLevels = $PermissionLevels
                        GrantedThrough = "SharePoint Group (Empty)"
                    }
                }
            }
            catch {
                Write-Warning "Could not retrieve members for group: $($Member.Title)"
                $Global:PermissionData += [PSCustomObject]@{
                    Object = $ObjectType
                    Title = $ObjectTitle
                    URL = $ObjectURL
                    HasUniquePermissions = $HasUniquePermissions
                    PrincipalType = "SharePointGroup"
                    PrincipalName = $Member.Title
                    PrincipalLoginName = $Member.LoginName
                    GroupName = ""
                    PermissionLevels = $PermissionLevels
                    GrantedThrough = "SharePoint Group (Access Denied)"
                }
            }
        }
        else {
            # Direct user permissions
            $Global:PermissionData += [PSCustomObject]@{
                Object = $ObjectType
                Title = $ObjectTitle
                URL = $ObjectURL
                HasUniquePermissions = $HasUniquePermissions
                PrincipalType = $PrincipalType
                PrincipalName = $Member.Title
                PrincipalLoginName = $Member.LoginName
                GroupName = ""
                PermissionLevels = $PermissionLevels
                GrantedThrough = "Direct Permissions"
            }
        }
    }
}

# Function to process document libraries
Function Get-DocumentLibraryPermissions {
    param(
        [Parameter(Mandatory=$true)]$Web
    )
    
    Write-Host "Analyzing Document Libraries..." -ForegroundColor Yellow
    
    # Get all document libraries (exclude system lists)
    $Lists = Get-PnPList -Web $Web | Where-Object { 
        $_.BaseTemplate -eq 101 -and  # Document Library
        $_.Hidden -eq $false -and
        $_.Title -notin @("Form Templates", "Site Assets", "Site Pages", "Images", "Pages", "Preservation Hold Library", "Style Library")
    }
    
    foreach ($List in $Lists) {
        Write-Host "`tProcessing Library: $($List.Title)" -ForegroundColor Green
        
        # Get list with role assignments
        $ListWithPermissions = Get-PnPList -Identity $List.Id -Includes RoleAssignments
        
        # Check if library has unique permissions
        if ($ListWithPermissions.HasUniqueRoleAssignments) {
            Write-Host "`t`tLibrary has unique permissions" -ForegroundColor Cyan
            Get-SharePointPermissions -Object $ListWithPermissions -ObjectType "Document Library" -ObjectTitle $List.Title -ObjectURL "$($Web.Url)/$($List.DefaultViewUrl)"
        }
        else {
            # Record that it inherits permissions
            $Global:PermissionData += [PSCustomObject]@{
                Object = "Document Library"
                Title = $List.Title
                URL = "$($Web.Url)/$($List.DefaultViewUrl)"
                HasUniquePermissions = $false
                PrincipalType = "Inherited"
                PrincipalName = "Inherits from parent site"
                PrincipalLoginName = ""
                GroupName = ""
                PermissionLevels = "Inherited"
                GrantedThrough = "Permission Inheritance"
            }
        }
    }
}

# Main execution
try {
    Write-Host "Starting SharePoint Permissions Analysis..." -ForegroundColor Yellow
    Write-Host "Site URL: $SiteURL" -ForegroundColor White
    Write-Host "Report will be saved to: $ReportPath" -ForegroundColor White
    Write-Host ""
    
    # Connect to SharePoint site
    Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteURL -Interactive
    
    # Get the root web with role assignments
    Write-Host "Getting site permissions..." -ForegroundColor Yellow
    $Web = Get-PnPWeb -Includes RoleAssignments
    
    # Get site collection administrators first
    Write-Host "Getting Site Collection Administrators..." -ForegroundColor Yellow
    $SiteAdmins = Get-PnPSiteCollectionAdmin
    $SiteAdminNames = ($SiteAdmins | Select-Object -ExpandProperty LoginName) -join ", "
    
    # Add site collection administrators to report
    $Global:PermissionData += [PSCustomObject]@{
        Object = "Site Collection"
        Title = $Web.Title
        URL = $Web.Url
        HasUniquePermissions = $true
        PrincipalType = "SiteCollectionAdmin"
        PrincipalName = $SiteAdminNames
        PrincipalLoginName = $SiteAdminNames
        GroupName = ""
        PermissionLevels = "Site Collection Administrator"
        GrantedThrough = "Site Collection Administration"
    }
    
    # Get site permissions
    Write-Host "Analyzing site permissions..." -ForegroundColor Yellow
    Get-SharePointPermissions -Object $Web -ObjectType "Site" -ObjectTitle $Web.Title -ObjectURL $Web.Url
    
    # Get document library permissions
    Get-DocumentLibraryPermissions -Web $Web
    
    # Export to CSV
    Write-Host ""
    Write-Host "Exporting results to CSV..." -ForegroundColor Yellow
    $Global:PermissionData | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    
    Write-Host ""
    Write-Host "✓ Analysis Complete!" -ForegroundColor Green
    Write-Host "✓ Total permissions found: $($Global:PermissionData.Count)" -ForegroundColor Green
    Write-Host "✓ Report saved to: $ReportPath" -ForegroundColor Green
    
    # Show summary statistics
    $UniqueLibraries = $Global:PermissionData | Where-Object { $_.Object -eq "Document Library" -and $_.HasUniquePermissions -eq $true } | Select-Object -Property Title -Unique
    Write-Host "✓ Document libraries with unique permissions: $($UniqueLibraries.Count)" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline
        Write-Host "Disconnected from SharePoint." -ForegroundColor Gray
    }
    catch {
        # Ignore disconnect errors
    }
}
