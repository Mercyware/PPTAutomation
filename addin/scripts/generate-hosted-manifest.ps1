param(
  [Parameter(Mandatory = $true)]
  [string]$BaseUrl,
  [string]$BackendUrl = "",
  [string]$OutputPath = "manifest.hosted.xml"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Normalize-HttpsUrl([string]$Value, [string]$Name) {
  $trimmed = ($Value | ForEach-Object { $_.Trim() }).TrimEnd("/")
  if (-not $trimmed) {
    throw "$Name is required."
  }
  if (-not $trimmed.StartsWith("https://")) {
    throw "$Name must start with https://"
  }
  return $trimmed
}

$normalizedBaseUrl = Normalize-HttpsUrl -Value $BaseUrl -Name "BaseUrl"
$normalizedBackendUrl = if ($BackendUrl) {
  Normalize-HttpsUrl -Value $BackendUrl -Name "BackendUrl"
} else {
  $normalizedBaseUrl
}

$addinRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$manifestPath = Join-Path $addinRoot "manifest.xml"
if (-not (Test-Path -LiteralPath $manifestPath)) {
  throw "Could not find manifest.xml at $manifestPath"
}

$outputFullPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) {
  $OutputPath
} else {
  Join-Path $addinRoot $OutputPath
}

$content = Get-Content -LiteralPath $manifestPath -Raw
$content = $content.Replace("https://localhost:3100", $normalizedBaseUrl)
$content = $content.Replace("http://localhost:4000", $normalizedBackendUrl)

Set-Content -LiteralPath $outputFullPath -Value $content -NoNewline
Write-Host "Generated hosted manifest at: $outputFullPath"
