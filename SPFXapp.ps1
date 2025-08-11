# Requires: Windows PowerShell 5.1, PnP.PowerShell 1.12.0
# Purpose: Inventory all installed SPFx (sppkg) apps (tenant and site scopes) across the tenant
# Auth: Uses -UseWebLogin for PS5 compatibility per request

param(
  [Parameter(Mandatory=$true)]
  [string]$TenantAdminUrl,            # e.g. https://contoso-admin.sharepoint.com
  [Parameter(Mandatory=$false)]
  [string]$OutputCsvPath = ".\SPFx-App-Installations.csv",
  [Parameter(Mandatory=$false)]
  [int]$DegreeOfParallelism = 6       # Tune for throttling vs speed
)

# Helper: Create a new PnP connection with WebLogin
function New-PnPConnectionWeb {
  param([string]$Url)
  # WebLogin is PS5-friendly and avoids interactive device code flow
  return Connect-PnPOnline -Url $Url -UseWebLogin -ReturnConnection
}

# Helper: Safely run code with retry (throttling/backoff)
function Invoke-WithRetry {
  param(
    [scriptblock]$Script,
    [int]$MaxRetries = 5
  )
  $delay = 2
  for ($i=0; $i -le $MaxRetries; $i++) {
    try { return & $Script }
    catch {
      if ($i -eq $MaxRetries) { throw }
      Start-Sleep -Seconds $delay
      $delay = [Math]::Min($delay*2, 60)
    }
  }
}

# 1) Connect to Admin Center
$adminConn = Invoke-WithRetry { New-PnPConnectionWeb -Url $TenantAdminUrl }

# 2) Get tenant app catalog URL (if present)
$tenantAppCatalogUrl = $null
try {
  $tenantAppCatalogUrl = (Invoke-WithRetry { Get-PnPTenantAppCatalogUrl -Connection $adminConn })
} catch {}

# 3) Get all site collections
$allSites = Invoke-WithRetry { Get-PnPTenantSite -Connection $adminConn -IncludeOneDriveSites:$false -Filter "Template ne 'SRCHCEN#0'" } # exclude classic search center noisily

# 4) Collect apps from Tenant app catalog (published/available centrally)
$results = New-Object System.Collections.Concurrent.ConcurrentBag[PSObject]

if ($tenantAppCatalogUrl) {
  $catConn = Invoke-WithRetry { New-PnPConnectionWeb -Url $tenantAppCatalogUrl }
  # Get all apps in tenant catalog and where enabled
  $tenantApps = Invoke-WithRetry { Get-PnPApp -Connection $catConn -Scope Tenant -ErrorAction Stop }
  foreach ($app in $tenantApps) {
    $obj = [pscustomobject]@{
      Scope                 = 'TenantCatalog'
      SiteUrl               = $tenantAppCatalogUrl
      AppTitle              = $app.Title
      AppId                 = $app.Id
      ProductId             = $app.ProductId
      Version               = $app.AppCatalogVersion
      Deployed              = $app.Deployed
      Enabled               = $app.Enabled
      InstalledVersion      = $app.InstalledVersion
      IsClientSideSolution  = $app.IsClientSideSolution
      CanUpgrade            = $app.CanUpgrade
      FromTenantAppCatalog  = $true
      Source                = $app.Source
    }
    $results.Add($obj)
  }
}

# 5) Function to scan a site collection for installed apps (site scope)
$siteScanScript = {
  param($siteUrl)

  # Import PnP if needed in runspace
  if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell module is required in runspaces."
  }

  function New-PnPConnectionWebLocal {
    param([string]$Url)
    return Connect-PnPOnline -Url $Url -UseWebLogin -ReturnConnection
  }

  function Invoke-WithRetryLocal {
    param([scriptblock]$Script, [int]$MaxRetries = 5)
    $delay = 2
    for ($i=0; $i -le $MaxRetries; $i++) {
      try { return & $Script }
      catch {
        if ($i -eq $MaxRetries) { throw }
        Start-Sleep -Seconds $delay
        $delay = [Math]::Min($delay*2, 60)
      }
    }
  }

  $localResults = @()
  try {
    $conn = Invoke-WithRetryLocal { New-PnPConnectionWebLocal -Url $siteUrl }

    # Gather site-scoped installed apps
    $siteApps = Invoke-WithRetryLocal { Get-PnPApp -Connection $conn -Scope Site -ErrorAction Stop }
    foreach ($app in $siteApps) {
      $localResults += [pscustomobject]@{
        Scope                 = 'Site'
        SiteUrl               = $siteUrl
        AppTitle              = $app.Title
        AppId                 = $app.Id
        ProductId             = $app.ProductId
        Version               = $app.AppCatalogVersion
        Deployed              = $app.Deployed
        Enabled               = $app.Enabled
        InstalledVersion      = $app.InstalledVersion
        IsClientSideSolution  = $app.IsClientSideSolution
        CanUpgrade            = $app.CanUpgrade
        FromTenantAppCatalog  = $false
        Source                = $app.Source
      }
    }

    # Optional: If the site has a site collection app catalog, also query that catalog for availability
    try {
      $hasSiteCatalog = Invoke-WithRetryLocal { Get-PnPSiteCollectionAppCatalog -Connection $conn -CurrentSite -ErrorAction Stop }
      if ($hasSiteCatalog) {
        # Switch connection to the root web (current) and read apps available in this site catalog
        $scApps = Invoke-WithRetryLocal { Get-PnPApp -Connection $conn -Scope Tenant -ErrorAction Stop }
        foreach ($app in $scApps) {
          $localResults += [pscustomobject]@{
            Scope                 = 'SiteCollectionCatalog'
            SiteUrl               = $siteUrl
            AppTitle              = $app.Title
            AppId                 = $app.Id
            ProductId             = $app.ProductId
            Version               = $app.AppCatalogVersion
            Deployed              = $app.Deployed
            Enabled               = $app.Enabled
            InstalledVersion      = $app.InstalledVersion
            IsClientSideSolution  = $app.IsClientSideSolution
            CanUpgrade            = $app.CanUpgrade
            FromTenantAppCatalog  = $false
            Source                = $app.Source
          }
        }
      }
    } catch {
      # ignore if no site collection app catalog
    }
  } catch {
    $localResults += [pscustomobject]@{
      Scope                 = 'Error'
      SiteUrl               = $siteUrl
      AppTitle              = $null
      AppId                 = $null
      ProductId             = $null
      Version               = $null
      Deployed              = $null
      Enabled               = $null
      InstalledVersion      = $null
      IsClientSideSolution  = $null
      CanUpgrade            = $null
      FromTenantAppCatalog  = $null
      Source                = $null
      Error                 = $_.Exception.Message
    }
  }

  return $localResults
}

# 6) Parallel scan sites with runspaces (throttled)
$iss = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$psModulePath = ($env:PSModulePath -split ';' | Where-Object { Test-Path $_ -and (Get-ChildItem $_ -Recurse -Filter 'PnP.PowerShell.psd1' -ErrorAction SilentlyContinue | Select-Object -First 1) }) | Select-Object -First 1
if ($psModulePath) {
  $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry 'env:PSModulePath' $env:PSModulePath $null))
}
$pool = [runspacefactory]::CreateRunspacePool(1, [Math]::Max(1,$DegreeOfParallelism), $iss, $Host)
$pool.Open()

$jobs = @()
foreach ($site in $allSites) {
  $ps = [powershell]::Create()
  $ps.RunspacePool = $pool
  [void]$ps.AddScript($siteScanScript).AddArgument($site.Url)
  $jobs += [pscustomobject]@{
    PS = $ps
    Handle = $ps.BeginInvoke()
    Url = $site.Url
  }
}

foreach ($job in $jobs) {
  try {
    $data = $job.PS.EndInvoke($job.Handle)
    foreach ($row in $data) { $results.Add($row) }
  } catch {
    $results.Add([pscustomobject]@{
      Scope                 = 'Error'
      SiteUrl               = $job.Url
      Error                 = $_.Exception.Message
    })
  } finally {
    $job.PS.Dispose()
  }
}

$pool.Close()
$pool.Dispose()

# 7) Output CSV
$ordered = $results | Select-Object Scope,SiteUrl,AppTitle,AppId,ProductId,Version,InstalledVersion,Deployed,Enabled,IsClientSideSolution,CanUpgrade,FromTenantAppCatalog,Source,Error
$ordered | Sort-Object SiteUrl, AppTitle | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Completed. Output: $OutputCsvPath" -ForegroundColor Green
