# Requires: Windows PowerShell 5.x, PnP PowerShell 1.9.x module
# Auth: WebLogin or DeviceLogin (no certificates, no provisioning template APIs)

param(
    [Parameter(Mandatory=$true)]
    [string]$SourceSiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$SourceListTitle,

    [Parameter(Mandatory=$true)]
    [string]$DestinationSiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$NewListTitle,

    [string]$NewListDescription = "",
    [ValidateSet("GenericList","DocumentLibrary","Announcements","Contacts","Calendar","Tasks","IssueTracking")]
    [string]$TemplateType = "GenericList",

    [switch]$UseDeviceLogin,   # alternative to WebLogin if browser SSO issues
    [switch]$ShowDebug
)

function Connect-NonInteractive {
    param([string]$Url, [switch]$DeviceLogin)
    if ($DeviceLogin) {
        Connect-PnPOnline -Url $Url -DeviceLogin
    }
    else {
        # Legacy browser popup login
        Connect-PnPOnline -Url $Url -UseWebLogin
    }
}

function Get-OotbInternalNames {
    # Common OOTB internal names to exclude; this list can be extended if needed
    @(
        "ID","ContentType","ContentTypeId","Title","Modified","Created","Author","Editor",
        "Attachments","Edit","LinkTitle","LinkTitleNoMenu","_UIVersionString","_ModerationStatus",
        "FileLeafRef","FileRef","File_x0020_Type","FSObjType","UniqueId","_HasCopyDestinations",
        "_CopySource","CheckoutUser","CheckedOutTitle","Created_x0020_Date","Modified_x0020_Date",
        "DocIcon","ItemChildCount","FolderChildCount","AppAuthor","AppEditor","SelectTitle",
        "PreviewOnForm","ParentVersionString","ParentLeafName","ParentUniqueId","ContentVersion",
        "Order","GUID","ComplianceAssetId"
    )
}

function Get-ListSchemaObject {
    param([string]$ListTitle)

    $list = Get-PnPList -Identity $ListTitle -Includes EnableContentTypes,ContentTypes,DefaultView,Views,Title,OnQuickLaunch,Description,BaseTemplate
    if (-not $list) { throw "List '$ListTitle' not found." }

    # Map template display name to base template integer if user passed a string
    $baseTemplate = $list.BaseTemplate

    # Get custom fields: exclude OOTB, sealed, readonly
    $ootb = Get-OotbInternalNames
    $allFields = Get-PnPField -List $ListTitle -Includes InternalName,Sealed,ReadOnlyField,SchemaXml,Hidden
    $customFields = $allFields | Where-Object {
        ($_.Sealed -eq $false) -and
        ($_.ReadOnlyField -eq $false) -and
        ($_.Hidden -eq $false) -and
        ($ootb -notcontains $_.InternalName)
    }

    $fieldXmls = @()
    foreach ($f in $customFields) {
        $fieldXmls += $f.SchemaXml
    }

    # Content types (skip the default Item CT unless it was explicitly added/extended)
    $cts = @()
    if ($list.EnableContentTypes) {
        $cts = $list.ContentTypes | Where-Object { $_.Sealed -eq $false -and $_.ReadOnly -eq $false -and $_.StringId -ne "0x01" } | ForEach-Object {
            [PSCustomObject]@{
                Name     = $_.Name
                Id       = $_.StringId
                Group    = $_.Group
                ReadOnly = $_.ReadOnly
            }
        }
    }

    # Views (only public views can be recreated programmatically)
    $views = @()
    foreach ($v in $list.Views) {
        if ($v.PersonalView -eq $true) { continue }
        $viewFields = @()
        foreach ($vf in $v.ViewFields) { $viewFields += $vf }
        $queryXml = $v.ViewQuery
        $rowLimit = $v.RowLimit
        $isPaged  = $v.Paged
        $default  = $v.DefaultView

        $views += [PSCustomObject]@{
            Title     = $v.Title
            Default   = [bool]$default
            Fields    = $viewFields
            QueryXml  = $queryXml
            RowLimit  = [int]$rowLimit
            Paged     = [bool]$isPaged
            ViewType  = $v.Type                 # e.g., HTML, Grid, Calendar (creation support varies)
        }
    }

    # Build schema object
    $schema = [PSCustomObject]@{
        List = [PSCustomObject]@{
            Title          = $list.Title
            Description    = $list.Description
            BaseTemplate   = [int]$baseTemplate
            OnQuickLaunch  = [bool]$list.OnQuickLaunch
            EnableCTs      = [bool]$list.EnableContentTypes
        }
        Fields = $fieldXmls
        ContentTypes = $cts
        Views = $views
    }

    if ($ShowDebug) {
        Write-Host "Captured fields:" $schema.Fields.Count
        Write-Host "Captured content types:" $schema.ContentTypes.Count
        Write-Host "Captured views:" $schema.Views.Count
    }

    return $schema
}

function New-ListFromSchemaObject {
    param(
        [Parameter(Mandatory=$true)][psobject]$Schema,
        [Parameter(Mandatory=$true)][string]$NewTitle,
        [string]$NewDescription = "",
        [int]$BaseTemplateOverride
    )

    $baseTemplate = if ($BaseTemplateOverride) { $BaseTemplateOverride } else { $Schema.List.BaseTemplate }
    if ($ShowDebug) { Write-Host "Creating list '$NewTitle' with template $baseTemplate" }

    # Create list
    $newList = New-PnPList -Title $NewTitle -Template $baseTemplate -OnQuickLaunch:$Schema.List.OnQuickLaunch -EnableContentTypes:$Schema.List.EnableCTs -Description $NewDescription -ErrorAction Stop

    # Add custom fields
    foreach ($fieldXml in $Schema.Fields) {
        # Replace List/Source references if present
        $sanitizedXml = $fieldXml
        # Ensure field does not already exist
        try {
            $field = Get-PnPField -List $NewTitle | Where-Object { $_.SchemaXml -eq $sanitizedXml -or $_.InternalName -match '^\s*Name="' }
        } catch { $field = $null }

        try {
            Add-PnPFieldFromXml -List $NewTitle -FieldXml $sanitizedXml -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning "Failed to add field from XML: $($_.Exception.Message)"
        }
    }

    # Attach content types (if any)
    foreach ($ct in $Schema.ContentTypes) {
        # Try bind by ID first; fallback by Name if ID not found
        $ctDef = Get-PnPContentType -Identity $ct.Id -ErrorAction SilentlyContinue
        if (-not $ctDef) {
            $ctDef = Get-PnPContentType | Where-Object { $_.Name -eq $ct.Name }
        }
        if ($ctDef) {
            try {
                Add-PnPContentTypeToList -List $NewTitle -ContentType $ctDef -ErrorAction Stop | Out-Null
            } catch {
                Write-Warning "Failed to add CT '$($ct.Name)': $($_.Exception.Message)"
            }
        }
        else {
            Write-Warning "Content Type '$($ct.Name)' not found in destination site; skipping."
        }
    }

    # Recreate views (public only)
    foreach ($view in $Schema.Views) {
        try {
            # Create the view; CAML Query can be re-applied via -Query
            $newView = Add-PnPView -List $NewTitle -Title $view.Title -Fields $view.Fields -RowLimit $view.RowLimit -Paged:$view.Paged -Query $view.QueryXml -ErrorAction Stop
            if ($view.Default -eq $true) {
                Set-PnPView -List $NewTitle -Identity $view.Title -Values @{ DefaultView = $true } | Out-Null
            }
        } catch {
            Write-Warning "Failed to create view '$($view.Title)': $($_.Exception.Message)"
        }
    }

    return $newList
}

# MAIN

# 1) Connect to SOURCE
Write-Host "Connect to SOURCE: $SourceSiteUrl"
Connect-NonInteractive -Url $SourceSiteUrl -DeviceLogin:$UseDeviceLogin

# 2) Capture schema object
$schema = Get-ListSchemaObject -ListTitle $SourceListTitle

# 3) Connect to DESTINATION
Write-Host "Connect to DESTINATION: $DestinationSiteUrl"
Connect-NonInteractive -Url $DestinationSiteUrl -DeviceLogin:$UseDeviceLogin

# 4) Create new list from schema
# Map template type keyword to integer if user supplied a friendly name
$tplMap = @{
    "GenericList"     = 100
    "DocumentLibrary" = 101
    "Announcements"   = 104
    "Contacts"        = 105
    "Calendar"        = 106
    "Tasks"           = 107
    "IssueTracking"   = 1100
}
$tplInt = if ($tplMap.ContainsKey($TemplateType)) { $tplMap[$TemplateType] } else { $tplMap["GenericList"] }

$newList = New-ListFromSchemaObject -Schema $schema -NewTitle $NewListTitle -NewDescription $NewListDescription -BaseTemplateOverride $tplInt

Write-Host "Done. Created list '$NewListTitle' at $DestinationSiteUrl"
