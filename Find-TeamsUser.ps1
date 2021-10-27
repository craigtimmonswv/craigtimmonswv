$user=Read-Host -Prompt "Enter email of person to find:"
Get-CsOnlineUser -Identity $user