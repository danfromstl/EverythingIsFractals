Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-LabConfig {
  param(
    [string]$Path = "$PSScriptRoot\..\..\config.local.json"
  )

  if (-not (Test-Path $Path)) {
    throw "Missing config.local.json at: $Path"
  }

  $cfg = Get-Content -Raw -Path $Path | ConvertFrom-Json

  # Minimal required fields (add more later if you want)
  foreach ($p in @(
    "lab.vmRoot",
    "vmware.baseDcVmx",
    "vmware.baseClientVmx",
    "vmware.cloneType",
    "domain.dnsName",
    "domain.netbiosName",
    "network.dcIp",
    "network.clientIp",
    "network.prefixLength",
    "network.dnsServer",
    "secrets.guestAdminPw",
    "secrets.safeModeAdminPw"
  )) {
    if (-not (Test-JsonPath -Obj $cfg -Path $p)) {
      throw "config.local.json is missing required field: $p"
    }
  }

  return $cfg
}

function Test-JsonPath {
  param(
    [Parameter(Mandatory)] $Obj,
    [Parameter(Mandatory)] [string]$Path
  )

  $cur = $Obj
  foreach ($part in $Path.Split(".")) {
    if ($null -eq $cur) { return $false }

    # ConvertFrom-Json returns PSCustomObject; properties are accessible via .psobject.Properties
    if ($cur.PSObject.Properties.Name -contains $part) {
      $cur = $cur.$part
    } else {
      return $false
    }
  }
  return $true
}

function Get-LabSecrets {
  param(
    [Parameter(Mandatory)] $Cfg
  )

  # Fix 2: keep everything in config.local.json by default,
  # but allow OPTIONAL env-var overrides if you ever want them.
  $guestPw = $Cfg.secrets.guestAdminPw
  $dsrmPw  = $Cfg.secrets.safeModeAdminPw

  $ovGuest = [Environment]::GetEnvironmentVariable("LAB_GUEST_ADMIN_PW")
  if (-not [string]::IsNullOrWhiteSpace($ovGuest)) { $guestPw = $ovGuest }

  $ovDsrm = [Environment]::GetEnvironmentVariable("LAB_DSRM_PW")
  if (-not [string]::IsNullOrWhiteSpace($ovDsrm)) { $dsrmPw = $ovDsrm }

  return [pscustomobject]@{
    guestAdminPw     = $guestPw
    safeModeAdminPw  = $dsrmPw
  }
}
