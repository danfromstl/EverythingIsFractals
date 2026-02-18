. C:\lab\00-helpers.ps1

$cfg = Read-LabConfig

$dcIp = $cfg.network.dcIp
if ([string]::IsNullOrWhiteSpace($dcIp)) { $dcIp = $cfg.dc.hostOnlyIp }

$dcName = $cfg.domain.dcHostname
if ([string]::IsNullOrWhiteSpace($dcName)) { $dcName = $cfg.dc.name }

if ([string]::IsNullOrWhiteSpace($dcIp)) { throw "Missing DC IP in config (network.dcIp or dc.hostOnlyIp)." }
if ([string]::IsNullOrWhiteSpace($dcName)) { throw "Missing DC hostname in config (domain.dcHostname or dc.name)." }

Set-HostOnlyStaticIp -Ip $dcIp -PrefixLength $cfg.network.prefixLength
Set-HostOnlyDns -DnsServers @("127.0.0.1")

if ($env:COMPUTERNAME -ne $dcName) {
  Rename-Computer -NewName $dcName -Force
  Restart-Computer -Force
}
