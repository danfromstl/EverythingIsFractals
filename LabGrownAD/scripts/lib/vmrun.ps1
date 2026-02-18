Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Stores the VM encryption password (VMware "Access Control"/encryption password),
# NOT the Windows guest account password.
$script:VmEncPw = $null

function Set-VmEncryptionPassword {
  param([AllowNull()][string]$Password)
  if ([string]::IsNullOrWhiteSpace($Password)) {
    $script:VmEncPw = $null
  } else {
    $script:VmEncPw = $Password
  }
}

function Get-VmrunPath {
  $candidates = @(
    "$env:ProgramFiles\VMware\VMware Workstation\vmrun.exe",
    "${env:ProgramFiles(x86)}\VMware\VMware Workstation\vmrun.exe"
  )

  foreach ($p in $candidates) { if (Test-Path $p) { return $p } }

  $cmd = Get-Command vmrun.exe -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  throw "vmrun.exe not found. Install VMware Workstation Pro and ensure vmrun is available."
}

function Add-VmEncryptionFlag {
  param([Parameter(Mandatory)][string[]]$Args)

  if ([string]::IsNullOrWhiteSpace($script:VmEncPw)) {
    return ,$Args
  }

  # Insert -vp <pw> immediately after "-T ws" if present.
  $idxT = [Array]::IndexOf($Args, "-T")
  if ($idxT -ge 0 -and ($idxT + 1) -lt $Args.Length -and $Args[$idxT + 1] -eq "ws") {
    $head = @()
    if ($idxT -gt 0) { $head = $Args[0..($idxT+1)] } else { $head = $Args[0..1] }

    $tail = @()
    if (($idxT + 2) -le ($Args.Length - 1)) { $tail = $Args[($idxT+2)..($Args.Length-1)] }

    return @($head + @("-vp", $script:VmEncPw) + $tail)
  }

  # Fallback: just prefix it (still works in practice for many vmrun commands)
  return @(@("-vp", $script:VmEncPw) + $Args)
}

function Invoke-VmrunInternal {
  param([Parameter(Mandatory)][Alias("Args")][string[]]$VmrunArgs)

  $vmrun = Get-VmrunPath
  $finalArgs = Add-VmEncryptionFlag -Args $VmrunArgs

  $output = @(& $vmrun @finalArgs 2>&1 | ForEach-Object { "$_" })
  $exitCode = if ($null -eq $LASTEXITCODE) { 0 } else { [int]$LASTEXITCODE }

  return [pscustomobject]@{
    ExitCode  = $exitCode
    Output    = $output
    FinalArgs = $finalArgs
  }
}

function Format-VmrunFailure {
  param([Parameter(Mandatory)]$Result)

  $details = @($Result.Output | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [Environment]::NewLine
  $base = "vmrun failed: $($Result.FinalArgs -join ' ')"
  if ([string]::IsNullOrWhiteSpace($details)) {
    return $base
  }
  return "$base`n$details"
}

function Invoke-Vmrun {
  param([Parameter(Mandatory)][Alias("Args")][string[]]$VmrunArgs)

  $result = Invoke-VmrunInternal -VmrunArgs $VmrunArgs
  if ($result.ExitCode -ne 0) {
    throw (Format-VmrunFailure -Result $result)
  }

  if ($result.Output.Count -gt 0) {
    $result.Output
  }
}

function Get-RunningVmxPaths {
  $result = Invoke-VmrunInternal -VmrunArgs @("-T","ws","list")
  if ($result.ExitCode -ne 0) {
    throw (Format-VmrunFailure -Result $result)
  }

  return @($result.Output | Where-Object { $_ -match "\.vmx$" } | ForEach-Object { $_.Trim() })
}

function Wait-ToolsReady {
  param(
    [Parameter(Mandatory)] [string]$VmxPath,
    [int]$TimeoutSeconds = 600
  )

  $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
  while ((Get-Date) -lt $deadline) {
    try {
      $result = Invoke-VmrunInternal -VmrunArgs @("-T","ws","checkToolsState",$VmxPath)
      if ($result.ExitCode -eq 0 -and (($result.Output -join "`n") -match "running")) { return }
    } catch {}
    Start-Sleep -Seconds 5
  }
  throw "VMware Tools not ready in guest after $TimeoutSeconds seconds: $VmxPath"
}

function Test-VmxEncrypted {
  param([Parameter(Mandatory)][string]$VmxPath)

  if (-not (Test-Path $VmxPath)) { return $false }
  return [bool](Select-String -Path $VmxPath -Pattern "^\s*vmx\.encryptionType\s*=|^\s*encryption\.keySafe\s*=|^\s*encryptedVM\.guid\s*=" -Quiet -ErrorAction SilentlyContinue)
}

function Get-VmxLockDirectories {
  param([Parameter(Mandatory)][string]$VmxPath)

  if (-not (Test-Path $VmxPath)) { return @() }

  $vmDir = Split-Path -Parent $VmxPath
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($VmxPath)
  $locks = Get-ChildItem -Path $vmDir -Directory -Filter "$baseName*.lck" -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty FullName

  return @($locks)
}

function Copy-ToGuest {
  param(
    [string]$VmxPath,
    [string]$GuestUser,
    [string]$GuestPass,
    [string]$HostPath,
    [string]$GuestPath
  )
  Invoke-Vmrun -Args @("-T","ws","-gu",$GuestUser,"-gp",$GuestPass,"copyFileFromHostToGuest",$VmxPath,$HostPath,$GuestPath)
}

function Invoke-PSInGuest {
  param(
    [string]$VmxPath,
    [string]$GuestUser,
    [string]$GuestPass,
    [string]$Ps1PathInGuest,
    [string]$ArgsString = ""
  )

  $cmd = "powershell.exe"
  $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$Ps1PathInGuest`" $ArgsString"
  Invoke-Vmrun -Args @("-T","ws","-gu",$GuestUser,"-gp",$GuestPass,"runProgramInGuest",$VmxPath,$cmd,$arg)
}

function Run-PSInGuest {
  param(
    [string]$VmxPath,
    [string]$GuestUser,
    [string]$GuestPass,
    [string]$Ps1PathInGuest,
    [string]$ArgsString = ""
  )

  Invoke-PSInGuest -VmxPath $VmxPath -GuestUser $GuestUser -GuestPass $GuestPass -Ps1PathInGuest $Ps1PathInGuest -ArgsString $ArgsString
}
