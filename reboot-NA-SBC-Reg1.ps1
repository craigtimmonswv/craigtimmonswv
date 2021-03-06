$sbcs = "sbc01","sbc02","sbc03"
cls
	$Username = "Rest Username"
	$Password = "Rest Password"
	$date=Get-Date -UFormat %b.%d.%Y
	foreach ($address in $sbcs)
	{
	#Authenticate to SBC
#	[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
	[Net.ServicePointManager]::SecurityProtocol = 'TLS11','TLS12','ssl3'
	$LoginUrl = "https://" + $Address + ".yourdomain.com/rest/login"
	$LoginCredentials = "Username=" + $Username + "&Password=" + $Password
	 
	Invoke-RestMethod -Uri $LoginUrl -Method Post -Body $LoginCredentials -SessionVariable ps
	 
	#Backup Gateways
	$args = ""
	$BackupUrl = "https://" + $Address + ".yourdomain.com/rest/system?action=reboot"
	Invoke-RestMethod -Uri $BackupUrl -Method POST -WebSession $ps
	}