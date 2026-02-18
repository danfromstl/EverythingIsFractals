function Get-HostOnlyAdapter {
  # Heuristic: NAT adapter usually has a default gateway; host-only usually does not.
  $up = Get-NetAdapter | Where-Object Status -eq 'Up'
  foreach ($a in $up) {
    $cfg = Get-NetIPConfiguration -InterfaceIndex $a.ifIndex
    if ($null -eq $cfg.IPv4DefaultGateway) { return $a }
  }
  throw "Could not locate host-only adapter (no IPv4 default gateway)."
}

function Set-HostOnlyStaticIp {
  param([string]$Ip, [int]$PrefixLength)

  $a = Get-HostOnlyAdapter
  # wipe existing IPv4 addresses on that adapter
  Get-NetIPAddress -InterfaceIndex $a.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue

  New-NetIPAddress -InterfaceIndex $a.ifIndex -IPAddress $Ip -PrefixLength $PrefixLength | Out-Null
}

function Set-HostOnlyDns {
  param([string[]]$DnsServers)

  $a = Get-HostOnlyAdapter
  Set-DnsClientServerAddress -InterfaceIndex $a.ifIndex -ServerAddresses $DnsServers
}

function Read-LabConfig {
  return (Get-Content "C:\lab\config.json" -Raw | ConvertFrom-Json)
}

function Read-LabSecrets {
  return (Get-Content "C:\lab\secrets.json" -Raw | ConvertFrom-Json)
}
