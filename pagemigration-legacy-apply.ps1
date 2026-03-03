<#
  Option 1: Export each modern page from the SOURCE site as a PnP provisioning template (XML),
  automatically convert Full-Width sections to normal one-column sections (GROUP sites don't support full-width),
  then apply the template to the TARGET site, and finally publish the page (best effort).
  Also optionally copies files from "Site Assets" (files only; skips folders).

  Run in: Windows PowerShell 5.1
  Module: SharePointPnPPowerShellOnline (legacy)
#>

param(
  [Parameter(Mandatory = $true)]
  [string]$SourceSiteUrl,

  [Parameter(Mandatory = $true)]
  [string]$TargetSiteUrl,

  [Parameter(Mandatory = $true)]
  [string]$WorkingFolder,

  [switch]$CopySiteAssets
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-FolderIfMissing {
  param([Parameter(Mandatory = $true)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

function Convert-FullWidthSectionsInTemplate {
  param([Parameter(Mandatory = $true)][string]$TemplatePath)

  # Pragmatic conversion: GROUP sites can't use full-width sections.
  # The export commonly includes "OneColumnFullWidth" and/or FullWidth="true".
  $xml = Get-Content -LiteralPath $TemplatePath -Raw

  $xml = $xml -replace "OneColumnFullWidth", "OneColumn"
  $xml = $xml -replace 'FullWidth="true"', 'FullWidth="false"'

  Set-Content -LiteralPath $TemplatePath -Value $xml -Encoding UTF8
}

Import-Module SharePointPnPPowerShellOnline -ErrorAction Stop

# --- Prepare folders
New-FolderIfMissing -Path $WorkingFolder
$pagesFolder  = Join-Path $WorkingFolder "pages-templates"
$assetsFolder = Join-Path $WorkingFolder "site-assets"
New-FolderIfMissing -Path $pagesFolder
if ($CopySiteAssets) { New-FolderIfMissing -Path $assetsFolder }

# --- Connect to source
Write-Host "Connecting to SOURCE (UseWebLogin)..."
Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin

Write-Host "Reading pages from Source: $SourceSiteUrl"
$srcItems = Get-PnPListItem -List "Site Pages" -PageSize 2000 -Fields "FileLeafRef","FileRef","FileSystemObjectType"

$pages = $srcItems | Where-Object {
  ($_.FileSystemObjectType -eq "File") -and
  ($_.FieldValues["FileLeafRef"] -ne $null) -and
  ($_.FieldValues["FileLeafRef"].ToString().ToLower().EndsWith(".aspx"))
}

Write-Host ("Found {0} pages in 'Site Pages'." -f $pages.Count)

# --- Migrate pages
$errorsLog = Join-Path $WorkingFolder "migration-errors.log"
if (Test-Path -LiteralPath $errorsLog) { Remove-Item -LiteralPath $errorsLog -Force }

foreach ($p in $pages) {
  $pageName = $p.FieldValues["FileLeafRef"]
  if (-not $pageName) { continue }

  $templatePath = Join-Path $pagesFolder ($pageName.ToString().Replace(".aspx", ".xml"))

  try {
    Write-Host "Exporting page to template: $pageName"
    Export-PnPClientSidePage -Identity $pageName -Out $templatePath -Force

    # Convert full-width sections for GROUP sites
    Convert-FullWidthSectionsInTemplate -TemplatePath $templatePath

    Write-Host "Connecting to TARGET (UseWebLogin)..."
    Connect-PnPOnline -Url $TargetSiteUrl -UseWebLogin

    Write-Host "Applying template to target: $pageName"
    Apply-PnPProvisioningTemplate -Path $templatePath -Handlers Pages

    # Best-effort publish
    try {
      Write-Host "Publishing page (best effort): $pageName"
      Set-PnPClientSidePage -Identity $pageName -Publish -ErrorAction Stop | Out-Null
    } catch {
      Write-Host "Publish skipped/failed for $pageName (non-fatal)."
    }

    # Back to source for next page
    Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin
  }
  catch {
    $msg = "[PAGE ERROR] $pageName :: $($_.Exception.Message)"
    Write-Host $msg
    Add-Content -LiteralPath $errorsLog -Value $msg

    # Attempt to reconnect to source and continue
    try { Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin } catch {}
    continue
  }
}

# --- Copy Site Assets (optional; files only)
if ($CopySiteAssets) {
  Write-Host "Copying 'Site Assets' from source to target..."

  try {
    Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin

    $assetItems = Get-PnPListItem -List "Site Assets" -PageSize 2000 -Fields "FileLeafRef","FileRef","FileSystemObjectType"
    $assetFiles = $assetItems | Where-Object {
      ($_.FileSystemObjectType -eq "File") -and
      ($_.FieldValues["FileRef"] -ne $null) -and
      ($_.FieldValues["FileLeafRef"] -ne $null)
    }

    foreach ($a in $assetFiles) {
      $fileName = $a.FieldValues["FileLeafRef"]
      $fileRef  = $a.FieldValues["FileRef"]
      if (-not $fileName -or -not $fileRef) { continue }

      try {
        Write-Host "Downloading asset: $fileName"
        Get-PnPFile -Url $fileRef -Path $assetsFolder -FileName $fileName -AsFile -Force

        $localPath = Join-Path $assetsFolder $fileName
        if (Test-Path -LiteralPath $localPath) {
          Write-Host "Uploading asset: $fileName"
          Connect-PnPOnline -Url $TargetSiteUrl -UseWebLogin
          Add-PnPFile -Path $localPath -Folder "SiteAssets" -Overwrite | Out-Null
        }

        Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin
      }
      catch {
        $msg = "[ASSET ERROR] $fileName :: $($_.Exception.Message)"
        Write-Host $msg
        Add-Content -LiteralPath $errorsLog -Value $msg

        # Best-effort reconnect back to source to continue
        try { Connect-PnPOnline -Url $SourceSiteUrl -UseWebLogin } catch {}
        continue
      }
    }
  }
  catch {
    $msg = "[ASSET PHASE ERROR] $($_.Exception.Message)"
    Write-Host $msg
    Add-Content -LiteralPath $errorsLog -Value $msg
  }
}

Write-Host "Done."
Write-Host "Target Site Pages: $TargetSiteUrl/SitePages"
Write-Host "Templates saved in: $pagesFolder"
Write-Host "Errors log: $errorsLog"