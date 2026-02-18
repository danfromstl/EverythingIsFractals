. C:\lab\00-helpers.ps1

$cfg = Read-LabConfig
$sec = Read-LabSecrets

$netbiosName = $cfg.domain.netbiosName
if ([string]::IsNullOrWhiteSpace($netbiosName)) { $netbiosName = $cfg.domain.netbios }

$domainName = $cfg.domain.dnsName
if ([string]::IsNullOrWhiteSpace($domainName)) { $domainName = $cfg.domain.fqdn }

if ([string]::IsNullOrWhiteSpace($netbiosName)) { throw "Missing NetBIOS name (domain.netbiosName or domain.netbios)." }
if ([string]::IsNullOrWhiteSpace($domainName)) { throw "Missing domain name (domain.dnsName or domain.fqdn)." }

$user = "$netbiosName\Administrator"
$pass = ConvertTo-SecureString $sec.guestAdminPw -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($user,$pass)

Add-Computer -DomainName $domainName -Credential $cred -Force -Restart
