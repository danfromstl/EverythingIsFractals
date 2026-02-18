$ErrorActionPreference = "Stop"
. "$PSScriptRoot\lib\config.ps1"
. "$PSScriptRoot\lib\vmrun.ps1"

$cfg = Get-LabConfig
$sec = Get-LabSecrets -Cfg $cfg

$vmEnc = $cfg.vmware.encryptionPassword
if ([string]::IsNullOrWhiteSpace($vmEnc)) { $vmEnc = $sec.guestAdminPw }
Set-VmEncryptionPassword -Password $vmEnc

$labRoot = $cfg.lab.vmRoot

$dcVmx = Join-Path $labRoot "DC\DC.vmx"
$wsVmx = Join-Path $labRoot "CLIENT\CLIENT.vmx"

function Stop-LabVmIfRunning {
  param([Parameter(Mandatory)][string]$VmxPath)

  if (-not (Test-Path $VmxPath)) { return }

  $running = @(Get-RunningVmxPaths)
  if ($running -notcontains $VmxPath) {
    "VM already stopped: $VmxPath"
    return
  }

  try {
    Invoke-Vmrun -Args @("-T","ws","stop",$VmxPath,"soft") | Out-Null
  } catch {
    Invoke-Vmrun -Args @("-T","ws","stop",$VmxPath,"hard") | Out-Null
  }

  "Stop requested: $VmxPath"
}

Stop-LabVmIfRunning -VmxPath $dcVmx
Stop-LabVmIfRunning -VmxPath $wsVmx

"Lab is stopping."
