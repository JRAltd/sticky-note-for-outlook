<#
.SYNOPSIS
  Creates a distributable ZIP package (email-notes-complete.zip) containing the full
  source tree, pre-built dist bundle, icons, and Outlook manifests.
#>
$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $ProjectRoot

$ZipOutput = Join-Path $ProjectRoot "email-notes-complete.zip"
if (Test-Path $ZipOutput) {
  Remove-Item $ZipOutput -Force
}

$ItemsToInclude = @(
  "manifest.xml",
  "manifest.public.xml",
  "manifest.prod.xml",
  "manifest.json",
  "store-assets",
  "package.json",
  "tsconfig.json",
  "tsconfig.node.json",
  "vite.config.ts",
  ".env.example",
  "taskpane.html",
  "dashboard.html",
  "commands.html",
  "server.py",
  "Start-EmailNotes.ps1",
  "README.md",
  "PUBLISHING-GUIDE.md",
  "netlify.toml",
  "vercel.json",
  ".github",
  "public",
  "src",
  "dist",
  "scripts"
)

$ValidPaths = $ItemsToInclude | Where-Object { Test-Path (Join-Path $ProjectRoot $_) }
Compress-Archive -Path $ValidPaths -DestinationPath $ZipOutput -Force
Write-Host "Created deployment archive: $ZipOutput" -ForegroundColor Green
