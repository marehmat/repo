<#
.SYNOPSIS
Mimic a user's permissions to another user on a SharePoint Online site.

.PARAMETER AdminUrl
SharePoint admin center URL (e.g. https://contoso-admin.sharepoint.com)

.PARAMETER GroupId
Microsoft 365 Group Id (GUID) of the site (optional if SiteUrl is given)

.PARAMETER SiteUrl
Full site URL (optional if GroupId is given)

.PARAMETER SourceUPN
User principal name of the user to mimic

.PARAMETER TargetUPN
User principal name of the user to receive equivalent permissions

.PARAMETER LogPath
Path to a CSV log file (created if missing)

.PARAMETER WhatIf
Preview mode; discover and log intended changes, but do not apply them

.NOTES
- Requires: Windows PowerShell 5.x
- Modules: Microsoft.Online.SharePoint.PowerShell, PnP.PowerShell (v1.9.x), Microsoft.Graph (Groups scope)
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  [Parameter(Mandatory=$true)] [string]$AdminUrl,
  [Parameter(Mandatory=$false)] [Guid]$GroupId,
  [Parameter(Mandatory=$false)] [string]$SiteUrl,
  [Parameter(Mandatory=$true)] [string]$SourceUPN,
  [Parameter(Mandatory=$true)] [string]$TargetUPN,
  [Parameter(Mandatory=$true)] [string]$LogPath,
  [switch]$WhatIf
)

#region Helpers
function Write-ChangeLog {
  param(
    [string]$Scope,      # M365Group | SPGroup | Web | List
    [string]$Action,     # AddMember | AddOwner | AddGroupMember | AddRole
    [string]$ObjectId,   # GroupId | WebUrl | ListTitle
    [string]$Role,       # e.g., 'Member','Owner','Read','Contribute'
    [string]$Details
  )
  $row = [pscustomobject]@{
    TimeUtc     = (Get-Date).ToUniversalTime().ToString("s")
    Scope       = $Scope
    Action      = $Action
    Object      = $ObjectId
    Role        = $Role
    SourceUPN   = $SourceUPN
    TargetUPN   = $TargetUPN
    Details     = $Details
  }
  $exists = Test-Path -Path $LogPath
  $row | Export-Csv -Path $LogPath -Append -NoTypeInformation -Force -Encoding UTF8
  if(-not $exists){ Write-Host "Log created: $LogPath" -ForegroundColor Cyan }
}

function Ensure-Module {
  param([string]$Name,[string]$MinVersion)
  if(-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Host "Installing $Name..." -ForegroundColor Yellow
    Install-Module $Name -MinimumVersion $MinVersion -Scope CurrentUser -Force -ErrorAction Stop
  }
  Import-Module $Name -MinimumVersion $MinVersion -ErrorAction Stop
}

function Resolve-SiteUrlFromGroup {
  param([Guid]$GroupId)
  # Requires connection to SPO Admin
  $site = Get-SPOSite -Limit All -Detailed | Where-Object { $_.GroupId -eq $GroupId } | Select-Object -First 1
  if(-not $site){ throw "No site found connected to GroupId $GroupId" }
  return $site.Url
}
#endregion Helpers

try {
  # Ensure modules (use specific versions where requested)
  Ensure-Module -Name Microsoft.Online.SharePoint.PowerShell -MinVersion "16.0.23512.12000"
  Ensure-Module -Name PnP.PowerShell -MinVersion "1.9.0"
  Ensure-Module -Name Microsoft.Graph -MinVersion "1.27.0"

  # Connect to SPO Admin
  Write-Host "Connecting to SharePoint Online Admin: $AdminUrl" -ForegroundColor Cyan
  Connect-SPOService -Url $AdminUrl

  # Resolve SiteUrl if only GroupId is provided
  if(-not $SiteUrl -and $GroupId){
    $SiteUrl = Resolve-SiteUrlFromGroup -GroupId $GroupId
    Write-Host "Resolved site: $SiteUrl" -ForegroundColor Green
  }
  if(-not $SiteUrl){ throw "SiteUrl is required if GroupId is not supplied." }

  # Get associated GroupId if not provided
  if(-not $GroupId){
    $siteInfo = Get-SPOSite -Identity $SiteUrl -Detailed
    $GroupId  = $siteInfo.GroupId
    if(-not $GroupId){ Write-Host "Site is not group-connected (no M365 Group)." -ForegroundColor Yellow }
  }

  # PnP connection (Windows PowerShell 5 + v1.9)
  Write-Host "Connecting PnP to $SiteUrl..." -ForegroundColor Cyan
  Connect-PnPOnline -Url $SiteUrl -Interactive

  # Connect to Microsoft Graph for group membership operations (if GroupId available)
  if($GroupId){
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All" | Out-Null
    Select-MgProfile -Name "v1.0" | Out-Null
  }

  # ------------------------------
  # 1) M365 Group (Owner/Member)
  # ------------------------------
  if($GroupId){
    # Check if Source is Owner
    $isSourceOwner   = $false
    $isSourceMember  = $false

    try {
      $owners = Get-MgGroupOwner -GroupId $GroupId -All
      if($owners | Where-Object {$_.UserPrincipalName -eq $SourceUPN}) { $isSourceOwner = $true }
    } catch { }

    # Check membership (direct)
    try {
      $members = Get-MgGroupMember -GroupId $GroupId -All
      if($members | Where-Object {$_.AdditionalProperties.userPrincipalName -eq $SourceUPN -or $_.UserPrincipalName -eq $SourceUPN}) {
        $isSourceMember = $true
      }
    } catch { }

    # Ensure target exists as user object
    $targetUser = Get-MgUser -Filter "userPrincipalName eq '$TargetUPN'" -All | Select-Object -First 1
    if(-not $targetUser){ throw "Target user not found in Entra ID: $TargetUPN" }

    if($isSourceOwner){
      Write-ChangeLog -Scope "M365Group" -Action "AddOwner" -ObjectId $GroupId -Role "Owner" -Details "Mirror owner because source is owner"
      if(-not $WhatIf){
        # Add owner if not already
        $existingOwners = Get-MgGroupOwner -GroupId $GroupId -All
        if(-not ($existingOwners | Where-Object Id -eq $targetUser.Id)){
          New-MgGroupOwner -GroupId $GroupId -DirectoryObjectId $targetUser.Id | Out-Null
        }
      }
    } elseif($isSourceMember){
      Write-ChangeLog -Scope "M365Group" -Action "AddMember" -ObjectId $GroupId -Role "Member" -Details "Mirror membership because source is member"
      if(-not $WhatIf){
        $existingMembers = Get-MgGroupMember -GroupId $GroupId -All
        if(-not ($existingMembers | Where-Object Id -eq $targetUser.Id)){
          New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $targetUser.Id | Out-Null
        }
      }
    } else {
      Write-Host "Source is neither owner nor direct member of the M365 Group; skipping M365-level grant." -ForegroundColor Yellow
    }
  }

  # ---------------------------------
  # 2) SharePoint Groups (site scope)
  # ---------------------------------
  $spGroups = Get-SPOSiteGroup -Site $SiteUrl
  foreach($g in $spGroups){
    # Use PnP to check membership
    $isMember = $false
    try {
      $member = Get-PnPGroupMember -Group $g.Title -User $SourceUPN -ErrorAction SilentlyContinue
      if($member){ $isMember = $true }
    } catch { }

    if($isMember){
      Write-ChangeLog -Scope "SPGroup" -Action "AddGroupMember" -ObjectId $g.Title -Role "GroupMember" -Details "Add to SP group because source is member"
      if(-not $WhatIf){
        # Add via SPO Management (preferred)
        Add-SPOUser -Site $SiteUrl -LoginName $TargetUPN -Group $g.Title -ErrorAction Stop | Out-Null
      }
    }
  }

  # ----------------------------------------------------
  # 3) Direct Web (site) permissions (role assignments)
  # ----------------------------------------------------
  $web = Get-PnPWeb -Includes RoleAssignments
  $webRAs = $web.RoleAssignments
  foreach($ra in $webRAs){
    $member = Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings
    # If direct user assignment to SourceUPN, mirror to target
    if($member.Member.PrincipalType -eq "User" -and $member.Member.LoginName -like "*|$SourceUPN"){
      $roles = $member.RoleDefinitionBindings | ForEach-Object { $_.Name }
      $roleText = ($roles -join ";")
      Write-ChangeLog -Scope "Web" -Action "AddRole" -ObjectId $SiteUrl -Role $roleText -Details "Direct web role(s) copied from source to target"
      if(-not $WhatIf){
        foreach($r in $roles){
          Set-PnPWebPermission -User $TargetUPN -AddRole $r | Out-Null
        }
      }
    }
  }

  # -------------------------------------------------------------------
  # 4) Document libraries with unique permissions (not inheriting web)
  # -------------------------------------------------------------------
  $lists = Get-PnPList -Includes Title, BaseType, Hidden, HasUniqueRoleAssignments, RoleAssignments
  $docLibs = $lists | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $_.HasUniqueRoleAssignments }

  foreach($list in $docLibs){
    $listTitle = $list.Title
    $las = Get-PnPProperty -ClientObject $list -Property RoleAssignments
    foreach($ra in $las){
      $raProps = Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings
      $member = $raProps.Member
      $roles  = $raProps.RoleDefinitionBindings | ForEach-Object { $_.Name }

      # Case A: direct user assignment to SourceUPN at list level
      if($member.PrincipalType -eq "User" -and $member.LoginName -like "*|$SourceUPN"){
        Write-ChangeLog -Scope "List" -Action "AddRole" -ObjectId $listTitle -Role ($roles -join ";") -Details "Direct list role(s) copied from source"
        if(-not $WhatIf){
          foreach($r in $roles){
            Set-PnPListPermission -Identity $listTitle -User $TargetUPN -AddRole $r | Out-Null
          }
        }
      }

      # Case B: source has access via an SP group at list level
      if($member.PrincipalType -eq "SharePointGroup"){
        # Is source inside this group?
        $gm = Get-PnPGroupMember -Group $member.Title -User $SourceUPN -ErrorAction SilentlyContinue
        if($gm){
          Write-ChangeLog -Scope "List" -Action "AddGroupMember" -ObjectId $member.Title -Role "GroupMember" -Details "Target added to SP group used at list scope"
          if(-not $WhatIf){
            # Adding to the group once at site scope suffices; Add-SPOUser ensures membership
            Add-SPOUser -Site $SiteUrl -LoginName $TargetUPN -Group $member.Title -ErrorAction SilentlyContinue | Out-Null
          }
        }
      }
    }
  }

  Write-Host "Completed. Review change log at: $LogPath" -ForegroundColor Green

} catch {
  Write-Error $_.Exception.Message
  throw
}




<#
Example usage:

.\Mimic-SPOPerms.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -GroupId "00000000-0000-0000-0000-000000000000" `
  -SourceUPN "alice@contoso.com" `
  -TargetUPN "bob@contoso.com" `
  -LogPath "C:\Temp\MimicPerms_log.csv"

  
To preview without making changes:

.\Mimic-SPOPerms.ps1 -AdminUrl https://contoso-admin.sharepoint.com -SiteUrl https://contoso.sharepoint.com/sites/TeamA -SourceUPN alice@contoso.com -TargetUPN bob@contoso.com -LogPath C:\Temp\log.csv -WhatIf
#>


Version 2


$AdminUrl = "https://contoso-admin.sharepoint.com"
$GroupId  = $null   # or [Guid]"00000000-0000-0000-0000-000000000000"
$SiteUrl  = "https://contoso.sharepoint.com/sites/TeamA"
$SourceUPN = "alice@contoso.com"
$TargetUPN = "bob@contoso.com"
$LogPath   = "C:\Temp\MimicPerms_log.csv"
$WhatIf    = $false  # set $true to preview only


function Write-ChangeLog {
  param([string]$Scope,[string]$Action,[string]$ObjectId,[string]$Role,[string]$Details)
  [pscustomobject]@{
    TimeUtc   = (Get-Date).ToUniversalTime().ToString("s")
    Scope     = $Scope
    Action    = $Action
    Object    = $ObjectId
    Role      = $Role
    SourceUPN = $SourceUPN
    TargetUPN = $TargetUPN
    Details   = $Details
  } | Export-Csv -Path $LogPath -Append -NoTypeInformation -Force -Encoding UTF8
}

function Ensure-Module { param([string]$Name,[string]$MinVersion)
  if(-not (Get-Module -ListAvailable -Name $Name)){ Install-Module $Name -MinimumVersion $MinVersion -Scope CurrentUser -Force }
  Import-Module $Name -MinimumVersion $MinVersion -ErrorAction Stop
}

try {
  Ensure-Module -Name Microsoft.Online.SharePoint.PowerShell -MinVersion "16.0.23512.12000"
  Ensure-Module -Name PnP.PowerShell -MinVersion "1.9.0"

  Connect-SPOService -Url $AdminUrl

  if(-not $SiteUrl -and $GroupId){
    $SiteUrl = (Get-SPOSite -Limit All -Detailed | Where-Object { $_.GroupId -eq $GroupId } | Select-Object -First 1).Url
    if(-not $SiteUrl){ throw "No site found bound to GroupId $GroupId" }
  }
  if(-not $SiteUrl){ throw "Provide SiteUrl or GroupId." }

  Connect-PnPOnline -Url $SiteUrl -Interactive

  # 1) SharePoint groups: mirror membership
  $spGroups = Get-SPOSiteGroup -Site $SiteUrl
  foreach($g in $spGroups){
    $inGroup = $false
    try {
      $inGroup = [bool](Get-PnPGroupMember -Group $g.Title -User $SourceUPN -ErrorAction SilentlyContinue)
    } catch {}
    if($inGroup){
      Write-ChangeLog -Scope "SPGroup" -Action "AddGroupMember" -ObjectId $g.Title -Role "GroupMember" -Details "Add target where source is member"
      if(-not $WhatIf){
        Add-SPOUser -Site $SiteUrl -LoginName $TargetUPN -Group $g.Title -ErrorAction SilentlyContinue | Out-Null
      }
    }
  }

  # 2) Site-level direct role assignments
  $web = Get-PnPWeb -Includes RoleAssignments
  foreach($ra in $web.RoleAssignments){
    $props = Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings
    if($props.Member.PrincipalType -eq "User" -and $props.Member.LoginName -like "*|$SourceUPN"){
      $roles = $props.RoleDefinitionBindings | ForEach-Object { $_.Name }
      if($roles.Count -gt 0){
        Write-ChangeLog -Scope "Web" -Action "AddRole" -ObjectId $SiteUrl -Role ($roles -join ";") -Details "Copy direct web roles"
        if(-not $WhatIf){ foreach($r in $roles){ Set-PnPWebPermission -User $TargetUPN -AddRole $r | Out-Null } }
      }
    }
  }

  # 3) Document libraries with unique permissions
  $lists = Get-PnPList -Includes Title, BaseType, Hidden, HasUniqueRoleAssignments, RoleAssignments
  $docLibs = $lists | Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden -and $_.HasUniqueRoleAssignments }
  foreach($list in $docLibs){
    $las = Get-PnPProperty -ClientObject $list -Property RoleAssignments
    foreach($ra in $las){
      $p = Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings
      $roles = $p.RoleDefinitionBindings | ForEach-Object { $_.Name }

      # Direct user at list level
      if($p.Member.PrincipalType -eq "User" -and $p.Member.LoginName -like "*|$SourceUPN"){
        Write-ChangeLog -Scope "List" -Action "AddRole" -ObjectId $list.Title -Role ($roles -join ";") -Details "Copy list roles"
        if(-not $WhatIf){ foreach($r in $roles){ Set-PnPListPermission -Identity $list.Title -User $TargetUPN -AddRole $r | Out-Null } }
      }

      # Access via SP group at list level
      if($p.Member.PrincipalType -eq "SharePointGroup"){
        $isMember = Get-PnPGroupMember -Group $p.Member.Title -User $SourceUPN -ErrorAction SilentlyContinue
        if($isMember){
          Write-ChangeLog -Scope "List" -Action "AddGroupMember" -ObjectId $p.Member.Title -Role "GroupMember" -Details "Ensure target in list-permission group"
          if(-not $WhatIf){ Add-SPOUser -Site $SiteUrl -LoginName $TargetUPN -Group $p.Member.Title -ErrorAction SilentlyContinue | Out-Null }
        }
      }
    }
  }

  Write-Host "Done. Review log at $LogPath" -ForegroundColor Green
}
catch {
  Write-Error $_.Exception.Message
  throw
}







