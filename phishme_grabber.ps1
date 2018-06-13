# PhishMe Search Script v1
# copyright by Artyom Ageyev
# 19/04/2018

Write-Host "================= PhishMe Search Script ===================="

# ==== VARIABLES =====
$attachment_dir = ""
$email = ""
$email_folder = "Inbox"

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI") 

#get phishme mail folder content
$folder = $namespace.Folders($email).Folders($email_folder)

while($true){

	#get time window
	$now = Get-date
	if ($now.Minute -lt 30)
	{
		$date_from = $now.Date.AddHours($now.Hour - 1).AddMinutes(30)
		$date_to = $now.Date.AddHours($now.Hour)
	} 
	else 
	{
		$date_from = $now.Date.AddHours($now.Hour)
		$date_to = $now.Date.AddHours($now.Hour).AddMinutes(30)
	}

	Write-Host "$("Working with emails from") $date_from $(" to ") $date_to"
	Write-Host "[INFO] Building mail database. It could take some time ... " -NoNewline
	$mails = $folder.items | where-object { ($_.ReceivedTime -ge $date_from) -and ($_.ReceivedTime -lt $date_to)}  
	# $mails = $folder.items | where-object { $_.ReceivedTime -gt [DateTime]::ParseExact($date, 'd/M/yyyy HH:mm:ss',[CultureInfo]::InvariantCulture)} 
	Write-Host "Done" -ForegroundColor Green

	#save attachments to NAS
	foreach ($mail in $mails)
	{
		$filename = "$($mail.ReceivedTime.ToString('yyyyMMdd-HHmmss')) - $($mail.attachments(1).FileName)"
		Write-Host "$("[INFO] Saving mail ") $filename"
		$mail.attachments(1).saveasfile($attachment_dir + $filename)
	}
    write-Host "[INFO] waiting for 5 minutes before loop"
	Sleep -Seconds (New-TimeSpan -Minute 15).TotalSeconds 
}