$ErrorActionPreference = "Stop"

. "$PSScriptRoot\lib\config.ps1"
. "$PSScriptRoot\lib\vmrun.ps1"

$cfg = Get-LabConfig
$sec = Get-LabSecrets -Cfg $cfg

# VM encryption password (falls back to guest pw)
$vmEnc = $cfg.vmware.encryptionPassword
if ([string]::IsNullOrWhiteSpace($vmEnc)) { $vmEnc = $sec.guestAdminPw }

Set-VmEncryptionPassword -Password $vmEnc

function Invoke-LabClone {
  param(
    [Parameter(Mandatory)][string]$Role,
    [Parameter(Mandatory)][string]$SourceVmx,
    [Parameter(Mandatory)][string]$DestVmx,
    [Parameter(Mandatory)][string]$CloneType
  )

  try {
    Invoke-Vmrun -Args @("-T","ws","clone",$SourceVmx,$DestVmx,$CloneType) | Out-Null
    return
  } catch {
    $hints = @()

    if (Test-VmxEncrypted -VmxPath $SourceVmx) {
      $hints += "Source VM appears encrypted (often due vTPM). vmrun clone commonly fails for encrypted templates."
      $hints += "Create an unencrypted template for automation, or clone this VM once in the VMware UI and use that clone as the new base VM."
    }

    $locks = @(Get-VmxLockDirectories -VmxPath $SourceVmx)
    if ($locks.Count -gt 0) {
      $hints += "Source VM has lock folders: $($locks -join '; '). Close VMware and clear stale *.lck folders before retrying."
    }

    $hintText = ""
    if ($hints.Count -gt 0) {
      $hintText = "`nHints:`n - " + ($hints -join "`n - ")
    }

    throw "Failed cloning $Role VM.`n$($_.Exception.Message)$hintText"
  }
}

$guestUser = "Administrator"
$guestPass = $sec.guestAdminPw

$labRoot = $cfg.lab.vmRoot
$dcVmx   = Join-Path $labRoot "DC\DC.vmx"
$wsVmx   = Join-Path $labRoot "CLIENT\CLIENT.vmx"

# Clone base VMs (linked is fast; full is independent but bigger)
$cloneType = $cfg.vmware.cloneType  # "linked" or "full"
if ($cloneType -notin @("linked","full")) {
  throw "vmware.cloneType must be 'linked' or 'full'. Current value: '$cloneType'"
}

foreach ($baseVmx in @($cfg.vmware.baseDcVmx, $cfg.vmware.baseClientVmx)) {
  if (-not (Test-Path $baseVmx)) {
    throw "Base VMX not found: $baseVmx"
  }
}

# Guard against overwrites
if (Test-Path $labRoot) {
  $existing = @(Get-ChildItem -Path $labRoot -Force -ErrorAction SilentlyContinue)
  if ($existing.Count -gt 0) {
    throw "Lab path already contains files: $labRoot. Run scripts/down.ps1, then remove/empty this folder (or choose a different lab.vmRoot)."
  }
}

# Create directories
New-Item -ItemType Directory -Force -Path (Split-Path $dcVmx) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $wsVmx) | Out-Null

Invoke-LabClone -Role "DC" -SourceVmx $cfg.vmware.baseDcVmx -DestVmx $dcVmx -CloneType $cloneType
Invoke-LabClone -Role "CLIENT" -SourceVmx $cfg.vmware.baseClientVmx -DestVmx $wsVmx -CloneType $cloneType

# Start both VMs headless
Invoke-Vmrun -Args @("-T","ws","start",$dcVmx,"nogui")
Invoke-Vmrun -Args @("-T","ws","start",$wsVmx,"nogui")

Wait-ToolsReady -VmxPath $dcVmx
Wait-ToolsReady -VmxPath $wsVmx

# Prepare transient payloads
$temp = New-Item -ItemType Directory -Force -Path (Join-Path $env:TEMP ("lab-" + [guid]::NewGuid().ToString()))
$configPath  = Join-Path $temp.FullName "config.json"
$secretsPath = Join-Path $temp.FullName "secrets.json"

$cfg | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -Path $configPath
$sec | ConvertTo-Json -Depth 5  | Set-Content -Encoding UTF8 -Path $secretsPath

# Ensure guest folders exist (best-effort)
Invoke-Vmrun -Args @("-T","ws","-gu",$guestUser,"-gp",$guestPass,"runProgramInGuest",$dcVmx,"cmd.exe","/c","mkdir C:\lab\dc 2>nul & mkdir C:\lab\client 2>nul")
Invoke-Vmrun -Args @("-T","ws","-gu",$guestUser,"-gp",$guestPass,"runProgramInGuest",$wsVmx,"cmd.exe","/c","mkdir C:\lab\dc 2>nul & mkdir C:\lab\client 2>nul")

# Copy config + secrets into each guest
Copy-ToGuest $dcVmx $guestUser $guestPass $configPath  "C:\lab\config.json"
Copy-ToGuest $dcVmx $guestUser $guestPass $secretsPath "C:\lab\secrets.json"
Copy-ToGuest $wsVmx $guestUser $guestPass $configPath  "C:\lab\config.json"
Copy-ToGuest $wsVmx $guestUser $guestPass $secretsPath "C:\lab\secrets.json"

# Copy common helper
$guestCommon = Resolve-Path "$PSScriptRoot\..\guest\common\00-helpers.ps1"
Copy-ToGuest $dcVmx $guestUser $guestPass $guestCommon "C:\lab\00-helpers.ps1"
Copy-ToGuest $wsVmx $guestUser $guestPass $guestCommon "C:\lab\00-helpers.ps1"

# Copy role scripts
Get-ChildItem "$PSScriptRoot\..\guest\dc\*.ps1" | ForEach-Object {
  Copy-ToGuest $dcVmx $guestUser $guestPass $_.FullName ("C:\lab\dc\" + $_.Name)
}
Get-ChildItem "$PSScriptRoot\..\guest\client\*.ps1" | ForEach-Object {
  Copy-ToGuest $wsVmx $guestUser $guestPass $_.FullName ("C:\lab\client\" + $_.Name)
}

# ---- DC provisioning ----
Invoke-PSInGuest $dcVmx $guestUser $guestPass "C:\lab\dc\01-network.ps1"
Wait-ToolsReady $dcVmx

Invoke-PSInGuest $dcVmx $guestUser $guestPass "C:\lab\dc\02-adds-forest.ps1"
Wait-ToolsReady $dcVmx

# Exchange schema prep (requires Exchange ISO mounted to DC VM)
Invoke-PSInGuest $dcVmx $guestUser $guestPass "C:\lab\dc\03-exchange-prepare.ps1"
Wait-ToolsReady $dcVmx

Invoke-PSInGuest $dcVmx $guestUser $guestPass "C:\lab\dc\04-seed-directory.ps1"
Wait-ToolsReady $dcVmx

# ---- Client provisioning ----
Invoke-PSInGuest $wsVmx $guestUser $guestPass "C:\lab\client\01-network.ps1"
Wait-ToolsReady $wsVmx

Invoke-PSInGuest $wsVmx $guestUser $guestPass "C:\lab\client\02-join-domain.ps1"
Wait-ToolsReady $wsVmx

Invoke-PSInGuest $wsVmx $guestUser $guestPass "C:\lab\client\03-smoke-test.ps1"
Wait-ToolsReady $wsVmx

# Cleanup
Remove-Item -Recurse -Force $temp.FullName

"Lab is up."
"DC VMX: $dcVmx"
"Client VMX: $wsVmx"
