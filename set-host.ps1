# Points manifest.xml at wherever you host the add-in files.
# Example:  .\set-host.ps1 -BaseUrl https://yourname.github.io/outlook-quick-sort
param(
  [Parameter(Mandatory = $true)]
  [string]$BaseUrl
)

$BaseUrl = $BaseUrl.TrimEnd('/')
if ($BaseUrl -notmatch '^https://') {
  throw "BaseUrl must start with https:// (Outlook only loads add-ins over HTTPS)."
}

$manifest = Join-Path $PSScriptRoot 'manifest.xml'
$xml = [System.IO.File]::ReadAllText($manifest)
$updated = [regex]::Replace($xml, 'https://[^"<\s]+?(?=/(taskpane\.html|assets/))', $BaseUrl)
[System.IO.File]::WriteAllText($manifest, $updated, (New-Object System.Text.UTF8Encoding($false)))

Write-Host "manifest.xml now points at $BaseUrl"
