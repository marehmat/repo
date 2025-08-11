param(
    [Parameter(Mandatory=$true)]
    [string]$TenantAdminUrl,                         # e.g. https://contoso-admin.sharepoint.com

    [Parameter(Mandatory=$false)]
    [string]$WorkDir = (Get-Location).Path,          # Directory for CSVs

    [Parameter(Mandatory=$false)]
    [int]$DegreeOfParallelism = 6,                   # Tune to avoid throttling (start 6–8)

    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 1000,                          # Smooth traffic, keep memory steady

    [Parameter(Mandatory=$false)]
    [int]$MaxAgeHoursForSiteList = 72,               # Rebuild site list if older than this

    [Parameter(Mandatory=$false)]
    [switch]$ForceRefreshSites                       # Force rebuild of the AllSites.csv
)

# ----------------------------------------------------------
# Paths
# ----------------------------------------------------------
$AllSitesCsv      = Join-Path $WorkDir "AllSites.csv"
$InstalledAppsCsv = Join-Path $WorkDir "InstalledApps.csv"

# ----------------------------------------------------------
# Helpers
# ----------------------------------------------------------
function Ensure-WorkDir {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Invoke-WithBackoff {
    param(
        [scriptblock]$Script,
        [int]$MaxRetries = 6,
        [int]$InitialDelay = 2,
        [int]$MaxDelay = 60
    )
    $delay = $InitialDelay
    for ($i=0; $i -le $MaxRetries; $i++) {
        try {
            return & $Script
        } catch {
            if ($i -eq $MaxRetries) { throw }
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, $MaxDelay)
        }
    }
}

function Get-OrBuild-SiteList {
    param(
        [string]$AdminUrl,
        [string]$CsvPath,
        [int]$MaxAgeHours,
        [switch]$Force
    )

    $needsBuild = $true
    if (Test-Path $CsvPath) {
        if (-not $Force) {
            $age = (Get-Date) - (Get-Item $CsvPath).LastWriteTime
            if ($age.TotalHours -le $MaxAgeHours) {
                Write-Host "Using cached site list ($([int]$age.TotalHours)h old): $CsvPath" -ForegroundColor Cyan
                $needsBuild = $false
            } else {
                Write-Host "Cached site list is stale ($([int]$age.TotalHours)h). Rebuilding..." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Force refresh requested. Rebuilding site list..." -ForegroundColor Yellow
        }
    } else {
        Write-Host "No cached site list found. Building new one..." -ForegroundColor Yellow
    }

    if ($needsBuild) {
        # Connect admin center with WebLogin (PS5 compatible)
        $adminConn = Invoke-WithBackoff { Connect-PnPOnline -Url $AdminUrl -UseWebLogin -ReturnConnection }

        # Enumerate site collections (exclude OneDrive by parameter)
        $sites = Invoke-WithBackoff { Get-PnPTenantSite -Connection $adminConn -IncludeOneDriveSites:$false }

        # Local filtering to skip non-target templates (adjust as needed)
        $sites = $sites | Where-Object {
            $_.Template -ne 'SRCHCEN#0' -and      # Classic search center
            $_.Template -ne 'APPCATALOG#0'        # Tenant app catalog site
        }

        # Save a clean, minimal CSV (expand if you need more fields later)
        $sites | Select-Object Url, Title, Template, Owner |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        Write-Host "Saved $($sites.Count) sites to $CsvPath" -ForegroundColor Green
        return $sites | Select-Object Url, Title, Template, Owner
    }

    return Import-Csv -Path $CsvPath
}

# ----------------------------------------------------------
# Main
# ----------------------------------------------------------
Ensure-WorkDir -Path $WorkDir

# Validate PnP module presence up front (optional but helpful)
try {
    $null = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction Stop
} catch {
    Write-Host "PnP.PowerShell module not found. Please install PnP.PowerShell 1.12.0 before running." -ForegroundColor Red
    Write-Host "Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope AllUsers" -ForegroundColor Yellow
    throw
}

# Build or load the site list
$allSites = Get-OrBuild-SiteList -AdminUrl $TenantAdminUrl -CsvPath $AllSitesCsv -MaxAgeHours $MaxAgeHoursForSiteList -Force:$ForceRefreshSites
if (-not $allSites -or $allSites.Count -eq 0) {
    throw "No sites found to scan."
}

# Prepare output bag
$results = New-Object System.Collections.Concurrent.ConcurrentBag[object]

# Runspace pool for bounded concurrency
$maxThreads = [Math]::Max(1, $DegreeOfParallelism)
$pool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
$pool.Open()

# Break into batches to smooth traffic and lower peak memory
$siteUrls = $allSites | Select-Object -ExpandProperty Url
$siteCount = $siteUrls.Count
$processed = 0

Write-Host "Scanning $siteCount sites with DOP=$maxThreads, BatchSize=$BatchSize ..." -ForegroundColor Cyan

for ($ofs = 0; $ofs -lt $siteUrls.Count; $ofs += $BatchSize) {
    $start = $ofs
    $end = [Math]::Min($ofs + $BatchSize - 1, $siteUrls.Count - 1)
    $batch = $siteUrls[$start..$end]
    Write-Host "Processing batch [$($start+1)-$($end+1)] of $siteCount..." -ForegroundColor Cyan

    $jobs = @()
    foreach ($url in $batch) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool

        # Worker: connect to site, get installed apps once, return rows or error
        $null = $ps.AddScript({
            param($u)

            Import-Module PnP.PowerShell -MinimumVersion 1.12.0 -ErrorAction Stop

            function Invoke-WithBackoffLocal {
                param([scriptblock]$S, [int]$Max=6, [int]$Init=2, [int]$MaxD=60)
                $d=$Init
                for($i=0;$i -le $Max;$i++){
                    try { return & $S } catch {
                        if ($i -eq $Max) { throw }
                        Start-Sleep -Seconds $d
                        $d=[Math]::Min($d*2,$MaxD)
                    }
                }
            }

            $rows = @()
            try {
                # Short-lived connection per site; WebLogin
                $conn = Invoke-WithBackoffLocal { Connect-PnPOnline -Url $u -UseWebLogin -ReturnConnection }

                # One call per site scope to list installed apps
                $apps = Invoke-WithBackoffLocal { Get-PnPApp -Connection $conn -Scope Site -ErrorAction Stop }

                foreach($a in $apps){
                    $rows += [pscustomobject]@{
                        SiteUrl              = $u
                        AppTitle             = $a.Title
                        AppId                = $a.Id
                        ProductId            = $a.ProductId
                        InstalledVersion     = $a.InstalledVersion
                        Deployed             = $a.Deployed
                        Enabled              = $a.Enabled
                        IsClientSideSolution = $a.IsClientSideSolution
                        Source               = $a.Source
                        CollectedAt          = (Get-Date).ToString("o")
                        Error                = $null
                    }
                }

                if ($apps.Count -eq 0) {
                    # Optional: emit a row indicating no apps for visibility; comment out to skip
                    $rows += [pscustomobject]@{
                        SiteUrl              = $u
                        AppTitle             = $null
                        AppId                = $null
                        ProductId            = $null
                        InstalledVersion     = $null
                        Deployed             = $null
                        Enabled              = $null
                        IsClientSideSolution = $null
                        Source               = $null
                        CollectedAt          = (Get-Date).ToString("o")
                        Error                = $null
                    }
                }
            } catch {
                $rows += [pscustomobject]@{
                    SiteUrl              = $u
                    AppTitle             = $null
                    AppId                = $null
                    ProductId            = $null
                    InstalledVersion     = $null
                    Deployed             = $null
                    Enabled              = $null
                    IsClientSideSolution = $null
                    Source               = $null
                    CollectedAt          = (Get-Date).ToString("o")
                    Error                = $_.Exception.Message
                }
            }

            return $rows
        }).AddArgument($url)

        $jobs += [pscustomobject]@{ PS=$ps; Handle=$ps.BeginInvoke() }
    }

    foreach($j in $jobs){
        try {
            $data = $j.PS.EndInvoke($j.Handle)
            foreach($row in $data){ $results.Add($row) }
        } finally {
            $j.PS.Dispose()
            $processed++
            if ($processed % 1000 -eq 0) {
                Write-Host "Progress: $processed/$siteCount sites processed..." -ForegroundColor DarkGray
            }
        }
    }

    # Gentle pacing between batches to avoid RU spikes
    Start-Sleep -Seconds 5
}

$pool.Close()
$pool.Dispose()

# Export results
$ordered = $results | Select-Object SiteUrl,AppTitle,AppId,ProductId,InstalledVersion,Deployed,Enabled,IsClientSideSolution,Source,CollectedAt,Error
$ordered | Export-Csv -Path $InstalledAppsCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done. Output written to: $InstalledAppsCsv" -ForegroundColor Green
Write-Host "Site list cached at: $AllSitesCsv" -ForegroundColor Green

# ------------------------ Usage notes ------------------------
# - Requires Windows PowerShell 5.1 and PnP.PowerShell 1.12.0.
# - Auth via -UseWebLogin (as requested). A login browser window will appear per process as needed.
# - Start with DegreeOfParallelism=6–8; increase only if throttling is minimal.
# - AllSites.csv is auto-created if missing; use -ForceRefreshSites to rebuild it on demand.
# - Filters exclude classic search center and tenant app catalog sites; adjust as needed.
