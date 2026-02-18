Import-Module ActiveDirectory

$u = Get-ADUser -Filter * -ResultSetSize 5 -Properties manager,directReports,msExchHideFromAddressLists |
  Select-Object SamAccountName,Manager,@{n="DRs";e={ ($_.directReports | Measure-Object).Count }},msExchHideFromAddressLists

$u | Format-Table -AutoSize
