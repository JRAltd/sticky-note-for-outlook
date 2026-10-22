<#
.SYNOPSIS
  Configures a shareable, publicly installable Outlook Add-in manifest (manifest.public.xml
  and dist/manifest.xml) pointing to your public HTTPS hosting URL (such as GitHub Pages:
  https://<username>.github.io/<repo-name>).

.EXAMPLE
  .\scripts\Set-PublicHostUrl.ps1 -PublicUrl "https://yourusername.github.io/sticky-note-for-outlook"
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$PublicUrl
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot

$CleanUrl = $PublicUrl.Trim().TrimEnd("/")
if (-not $CleanUrl.StartsWith("https://")) {
  throw "PublicUrl must start with https:// (Outlook requires HTTPS). Example: https://yourusername.github.io/sticky-note-for-outlook"
}

$UriObj = [System.Uri]$CleanUrl
$OriginOnly = "$($UriObj.Scheme)://$($UriObj.Authority)"

$RawManifest = Get-Content (Join-Path $ProjectRoot "manifest.xml") -Raw

# 1. Replace <AppDomain> with the origin only (Outlook requires <AppDomain> without path segments)
$PublicManifest = $RawManifest.Replace(
  "<AppDomain>https://localhost:3000</AppDomain>",
  "<AppDomain>$OriginOnly</AppDomain>"
)

# 2. Replace all asset, taskpane, and command URLs with the full base URL (supports GitHub Pages /<repo> subpaths)
$PublicManifest = $PublicManifest.Replace("https://localhost:3000", $CleanUrl)
$PublicManifest = $PublicManifest.Replace(
  "<SupportUrl DefaultValue=`"$CleanUrl/taskpane.html`"/>",
  "<SupportUrl DefaultValue=`"$CleanUrl/support.html`"/>"
)

$OutPublicManifest = Join-Path $ProjectRoot "manifest.public.xml"
$OutDistManifest   = Join-Path $ProjectRoot "dist\manifest.xml"

Set-Content -Path $OutPublicManifest -Value $PublicManifest -Encoding UTF8
Set-Content -Path $OutDistManifest   -Value $PublicManifest -Encoding UTF8

Write-Host "Updated public Outlook manifest!" -ForegroundColor Green
Write-Host " -> AppDomain Origin   : $OriginOnly" -ForegroundColor Cyan
Write-Host " -> Base Asset URL     : $CleanUrl" -ForegroundColor Cyan
Write-Host " -> Shareable Manifest : $OutPublicManifest" -ForegroundColor Cyan
Write-Host " -> Hosted Manifest    : $OutDistManifest" -ForegroundColor Cyan
