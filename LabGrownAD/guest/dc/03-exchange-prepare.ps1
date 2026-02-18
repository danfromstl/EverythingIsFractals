. C:\lab\00-helpers.ps1

$cfg = Read-LabConfig

function Find-ExchangeSetup {
  # Looks for <drive>:\Setup.exe
  foreach ($d in (Get-PSDrive -PSProvider FileSystem)) {
    $p = Join-Path $d.Root "Setup.exe"
    if (Test-Path $p) { return $p }
  }
  return $null
}

$setup = Find-ExchangeSetup
if (-not $setup) {
  Write-Host "Exchange Setup.exe not found. Mount the Exchange ISO to this VM and re-run."
  exit 0
}

$lic = "/IAcceptExchangeServerLicenseTerms_DiagnosticDataON"
$orgName = $cfg.exchange.organizationName
if ([string]::IsNullOrWhiteSpace($orgName)) { $orgName = $cfg.seed.companyName }
if ([string]::IsNullOrWhiteSpace($orgName)) { $orgName = "LabOrg" }

# Schema prep (requires Schema Admins + Enterprise Admins)
& $setup $lic "/PrepareSchema"
if ($LASTEXITCODE -ne 0) { throw "PrepareSchema failed with exit code $LASTEXITCODE" }

# Forest prep
& $setup $lic "/PrepareAD" "/OrganizationName:$orgName"
if ($LASTEXITCODE -ne 0) { throw "PrepareAD failed with exit code $LASTEXITCODE" }

# Domain prep
& $setup $lic "/PrepareAllDomains"
if ($LASTEXITCODE -ne 0) { throw "PrepareAllDomains failed with exit code $LASTEXITCODE" }

Write-Host "Exchange schema/forest/domain prep completed."
