. C:\lab\00-helpers.ps1

$cfg = Read-LabConfig
$sec = Read-LabSecrets

# AD DS role
Install-WindowsFeature AD-Domain-Services -IncludeManagementTools | Out-Null

$domainName = $cfg.domain.dnsName
if ([string]::IsNullOrWhiteSpace($domainName)) { $domainName = $cfg.domain.fqdn }

$netbiosName = $cfg.domain.netbiosName
if ([string]::IsNullOrWhiteSpace($netbiosName)) { $netbiosName = $cfg.domain.netbios }

$dsrmPw = $sec.safeModeAdminPw
if ([string]::IsNullOrWhiteSpace($dsrmPw)) { $dsrmPw = $sec.dsrmPw }

if ([string]::IsNullOrWhiteSpace($domainName)) { throw "Missing domain name (domain.dnsName or domain.fqdn)." }
if ([string]::IsNullOrWhiteSpace($netbiosName)) { throw "Missing NetBIOS name (domain.netbiosName or domain.netbios)." }
if ([string]::IsNullOrWhiteSpace($dsrmPw)) { throw "Missing DSRM password (secrets.safeModeAdminPw or secrets.dsrmPw)." }

$dsrm = ConvertTo-SecureString $dsrmPw -AsPlainText -Force

# Promote to new forest + install DNS
Install-ADDSForest `
  -DomainName $domainName `
  -DomainNetbiosName $netbiosName `
  -SafeModeAdministratorPassword $dsrm `
  -InstallDNS `
  -Force

# Install-ADDSForest reboots automatically unless suppressed
