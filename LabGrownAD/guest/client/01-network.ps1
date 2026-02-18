. C:\lab\00-helpers.ps1

$cfg = Read-LabConfig

$clientIp = $cfg.network.clientIp
if ([string]::IsNullOrWhiteSpace($clientIp)) { $clientIp = $cfg.client.hostOnlyIp }

$dnsServer = $cfg.network.dnsServer
if ([string]::IsNullOrWhiteSpace($dnsServer)) { $dnsServer = $cfg.dc.hostOnlyIp }

$clientName = $cfg.domain.clientHostname
if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = $cfg.client.name }

if ([string]::IsNullOrWhiteSpace($clientIp)) { throw "Missing client IP in config (network.clientIp or client.hostOnlyIp)." }
if ([string]::IsNullOrWhiteSpace($dnsServer)) { throw "Missing DNS server in config (network.dnsServer or dc.hostOnlyIp)." }
if ([string]::IsNullOrWhiteSpace($clientName)) { throw "Missing client hostname in config (domain.clientHostname or client.name)." }

Set-HostOnlyStaticIp -Ip $clientIp -PrefixLength $cfg.network.prefixLength
Set-HostOnlyDns -DnsServers @($dnsServer)

if ($env:COMPUTERNAME -ne $clientName) {
  Rename-Computer -NewName $clientName -Force
  Restart-Computer -Force
}
