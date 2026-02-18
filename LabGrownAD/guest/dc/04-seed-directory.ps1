. C:\lab\00-helpers.ps1

$cfg = Read-LabConfig
$sec = Read-LabSecrets
Import-Module ActiveDirectory

$domainName = $cfg.domain.dnsName
if ([string]::IsNullOrWhiteSpace($domainName)) { $domainName = $cfg.domain.fqdn }
if ([string]::IsNullOrWhiteSpace($domainName)) { throw "Missing domain name (domain.dnsName or domain.fqdn)." }

$userCount = [int]$cfg.seed.userCount
if ($userCount -le 0) { $userCount = 200 }

$hidePct = [double]$cfg.seed.hideFromAddressListsPercent
if ($hidePct -lt 0) { $hidePct = 0 }
if ($hidePct -gt 100) { $hidePct = 100 }
$hideChance = $hidePct / 100.0

$companyName = $cfg.seed.companyName
if ([string]::IsNullOrWhiteSpace($companyName)) { $companyName = "LabCo" }

$defaultPw = $sec.guestAdminPw
if ([string]::IsNullOrWhiteSpace($defaultPw)) { $defaultPw = "P@ssw0rd!Lab" }
$defaultPwSecure = ConvertTo-SecureString $defaultPw -AsPlainText -Force

# Quick schema check for msExchHideFromAddressLists (requires Exchange schema prep)
$schemaNC = (Get-ADRootDSE).schemaNamingContext
$hasHideAttr = $null -ne (Get-ADObject -SearchBase $schemaNC -LDAPFilter "(lDAPDisplayName=msExchHideFromAddressLists)" -ErrorAction SilentlyContinue)

# OUs
$baseDn = (Get-ADDomain).DistinguishedName
$ouCorp = "OU=Corp,$baseDn"
$ouUsers = "OU=Users,$ouCorp"
$ouHidden = "OU=Hidden,$ouCorp"

foreach ($ou in @("Corp","Users","Hidden")) {
  $dn = "OU=$ou,$baseDn"
  if (-not (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$dn)" -ErrorAction SilentlyContinue)) {
    New-ADOrganizationalUnit -Name $ou -Path $baseDn -ProtectedFromAccidentalDeletion $false | Out-Null
  }
}

# Simple deterministic fake names
$first = @("Ava","Ben","Cara","Dan","Eli","Faye","Gus","Hana","Ivan","Jules","Kara","Liam","Mina","Noah","Omar","Pia","Quinn","Rae","Seth","Tia")
$last  = @("Adams","Baker","Clark","Davis","Evans","Foster","Garcia","Hughes","Irwin","Jones","Klein","Lopez","Miller","Nguyen","Owens","Patel","Reed","Singh","Taylor","Young")

# Create managers
$managers = @()
for ($i=0; $i -lt 5; $i++) {
  $sam = "mgr$($i+1)"
  $name = "$($first[$i]) $($last[$i])"
  if (-not (Get-ADUser -Filter "sAMAccountName -eq '$sam'" -ErrorAction SilentlyContinue)) {
    New-ADUser -Name $name -SamAccountName $sam -UserPrincipalName "$sam@$domainName" `
      -Path $ouUsers -Enabled $true -AccountPassword $defaultPwSecure | Out-Null
  }
  $managers += (Get-ADUser -Identity $sam)
}

# Create employees and assign managers; hide ~10%
$rand = New-Object System.Random 42

for ($u=0; $u -lt $userCount; $u++) {
  $fn = $first[$rand.Next(0,$first.Count)]
  $ln = $last[$rand.Next(0,$last.Count)]
  $sam = ("u{0:000}" -f $u)
  $name = "$fn $ln"
  $mgr = $managers[$rand.Next(0,$managers.Count)]

  if (-not (Get-ADUser -Filter "sAMAccountName -eq '$sam'" -ErrorAction SilentlyContinue)) {
    New-ADUser -Name $name -SamAccountName $sam -UserPrincipalName "$sam@$domainName" `
      -Path $ouUsers -Enabled $true -Manager $mgr.DistinguishedName `
      -Title "Staff" -Department "Dept$($rand.Next(1,6))" -Company $companyName `
      -AccountPassword $defaultPwSecure | Out-Null
  }

  if ($hasHideAttr -and ($rand.NextDouble() -lt $hideChance)) {
    try {
      Set-ADUser -Identity $sam -Replace @{ msExchHideFromAddressLists = $true }
      # optionally move hidden users to OU=Hidden for convenience
      Move-ADObject -Identity (Get-ADUser $sam).DistinguishedName -TargetPath $ouHidden
    } catch {
      Write-Warning "Failed setting msExchHideFromAddressLists for ${sam}: $($_.Exception.Message)"
    }
  }
}

Write-Host "Seeded directory. Exchange hide attribute present: $hasHideAttr"
