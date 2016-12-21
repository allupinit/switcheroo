Write-Host " "
Write-Host "This script has only been tested with Exchange 2010 from an Exchange Powershell Console"
Write-Host " "
Write-Host "You are about to **TEMPORARILY** add/move an e-mail address to the user of your choosing."
Write-Host " "
Write-Host "Written by Keith Steckert on December 21, 2016"
Write-Host "Copyleft (MIT License) 2016 SFB Solutions, LLC"
Write-Host " "
Write-Host " "

$emladdr = Read-Host -Prompt 'E-Mail Address'
$usrfrom = (Get-Recipient "$emladdr" -erroraction 'silentlycontinue').Name

if ($usrfrom -eq $null) {
  Write-Host "No matching recipient found for '$emladdr'"
  Write-Host "The address will be temporarily added to a user."
  Write-Host " "
} else {
  Write-Host "'$emladdr' is currently assigned to '$usrfrom'"
  Write-Host " "
}

$usrtoprompt = Read-Host -Prompt 'Temporarily assign it to'
$usrto = (Get-Recipient "$usrtoprompt" -erroraction 'silentlycontinue').Name
Write-Host " "


if ($usrto -eq $null) {
  Write-Host "'$usrtoprompt' was not found on the Exchange Server"
  Write-Host "Aborting!!"
  break
}

if ($usrfrom -eq $null) {
  Write-Host "Assigning '$emladdr' to '$usrto'..."
  Write-Host " "
} else {
  Write-Host "Switching '$emladdr' from '$usrfrom' to '$usrto'..."
  Write-Host " "
  Set-Mailbox "$usrfrom" -EmailAddresses @{remove="$emladdr"}
}

Set-Mailbox "$usrto" -EmailAddresses @{add="$emladdr"}

Write-Host "Press any key to undo changes ..."
Write-Host " "
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

if ($usrfrom -eq $null) {
  Write-Host "Removing '$emladdr' from '$usrto'..."
  Write-Host " "
  Set-Mailbox "$usrto" -EmailAddresses @{remove="$emladdr"}
  Write-Host " "
  Write-Host "'$emladdr' has been removed from '$usrto'."
  Write-Host " "
} else {
  Write-Host "Switching '$emladdr' back to '$usrfrom' from '$usrto'..."
  Write-Host " "
  Set-Mailbox "$usrto" -EmailAddresses @{remove="$emladdr"}
  Set-Mailbox "$usrfrom" -EmailAddresses @{add="$emladdr"}
  Write-Host " "
  Write-Host "'$emladdr' is back to normal."
  Write-Host " "
}