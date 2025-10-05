# Script 1: Export on-prem list fields to JSON (SSOM)
# Run in SharePoint Management Shell on a SharePoint server

param(
  [Parameter(Mandatory=$true)] [string]$WebUrl,
  [Parameter(Mandatory=$true)] [string]$ListTitle,
  [string]$OutFile = ".\ListFields-$($ListTitle).json"
)

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$site = Get-SPSite $WebUrl  # Requires server-side object model session [web:35]
$web  = $site.OpenWeb()
$list = $web.Lists[$ListTitle]

# Helper to safely get property
function Get-Prop($obj, $name) {
  try { return $obj.$name } catch { return $null }
}

$fields = @()

foreach ($f in $list.Fields) {
  # Skip internal system columns unless needed
  if ($f.Sealed -or $f.ReadOnlyField) { } # still include, but flag

  $fieldInfo = [ordered]@{
    Id                    = $f.Id.Guid
    InternalName          = $f.InternalName
    Title                 = $f.Title
    Group                 = $f.Group
    TypeDisplayName       = $f.TypeDisplayName
    TypeAsString          = $f.TypeAsString
    FieldTypeKind         = $f.Type.ToString()
    Hidden                = $f.Hidden
    Required              = $f.Required
    ReadOnly              = $f.ReadOnlyField
    Sealed                = $f.Sealed
    DefaultValue          = (Get-Prop $f 'DefaultValue')
    SchemaXml             = $f.SchemaXml
    Indexed               = (Get-Prop $f 'Indexed')
    EnforceUniqueValues   = (Get-Prop $f 'EnforceUniqueValues')
    Description           = (Get-Prop $f 'Description')
    MaxLength             = (Get-Prop $f 'MaxLength')
    Choices               = @()
    FillInChoice          = $null
    LookupListId          = $null
    LookupField           = $null
    AllowMultipleValues   = $null
    TermSetId             = $null
    TermGroupId           = $null
    AnchorId              = $null
    IsMultilineRichText   = $null
  }

  # Choice/MultiChoice
  if ($f.Type -eq [Microsoft.SharePoint.SPFieldType]::Choice -or
      $f.Type -eq [Microsoft.SharePoint.SPFieldType]::MultiChoice) {
    $cf = [Microsoft.SharePoint.SPFieldChoice]$f
    $fieldInfo.Choices = @($cf.Choices)
    $fieldInfo.FillInChoice = $cf.FillInChoice
    $fieldInfo.AllowMultipleValues = ($f.Type -eq [Microsoft.SharePoint.SPFieldType]::MultiChoice)
  }

  # Lookup
  if ($f.Type -eq [Microsoft.SharePoint.SPFieldType]::Lookup) {
    $lf = [Microsoft.SharePoint.SPFieldLookup]$f
    $fieldInfo.LookupListId = $lf.LookupList
    $fieldInfo.LookupField  = $lf.LookupField
    $fieldInfo.AllowMultipleValues = $lf.AllowMultipleValues
  }

  # Multi-line text settings
  if ($f.Type -eq [Microsoft.SharePoint.SPFieldType]::Note) {
    $nf = [Microsoft.SharePoint.SPFieldMultiLineText]$f
    $fieldInfo.IsMultilineRichText = $nf.RichText
    $fieldInfo.NumberOfLines = $nf.NumberOfLines
  }

  # Managed Metadata (Taxonomy) via schema attributes
  $xml = [xml]$f.SchemaXml
  $tmm = $xml.Field.Attributes | Where-Object { $_.Name -in @('SspId','TermSetId','AnchorId') }
  if ($tmm) {
    $fieldInfo.TermSetId   = ($xml.Field.Attributes['TermSetId']).Value
    $fieldInfo.AnchorId    = ($xml.Field.Attributes['AnchorId']).Value
    $fieldInfo.SspId       = ($xml.Field.Attributes['SspId']).Value
  }

  $fields += [pscustomobject]$fieldInfo
}

$fields | ConvertTo-Json -Depth 6 | Out-File -Encoding UTF8 $OutFile

Write-Host "Exported $($fields.Count) fields to $OutFile"
$web.Dispose(); $site.Dispose()
