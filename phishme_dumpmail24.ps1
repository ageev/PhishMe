# PhishMe Search Script v1
# copyright by Artyom Ageyev
# 19/04/2018

Write-Host "================= PhishMe Search Script ===================="

# ==== < VARIABLES =====
$attachment_dir = ""
$data_dir = ""
$email = ""
$email_folder = "Inbox"
$lookback_time  = 24 #hours from now
# ========= /> =========

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI") 

#get phishme mail folder content
$folder = $namespace.Folders($email).Folders($email_folder)

	#get time window
    $now = Get-Date
	$date_from = (Get-date).AddHours(-$lookback_time)

	Write-Host "$("Working with emails from") $date_from $(" to ") $now"
	Write-Host "[INFO] Building mail database. It could take some time ... " -NoNewline
	$mails = $folder.items | where-object { $_.ReceivedTime -ge $date_from}  
	# $mails = $folder.items | where-object { $_.ReceivedTime -gt [DateTime]::ParseExact($date, 'd/M/yyyy HH:mm:ss',[CultureInfo]::InvariantCulture)} 
	Write-Host "Done" -ForegroundColor Green

	#save attachments to NAS
	foreach ($mail in $mails)
	{
		$filename = "$($mail.ReceivedTime.ToString('yyyyMMdd-HHmmss')) - $($mail.attachments(1).FileName)"
		Write-Host "$("[INFO] Saving mail ") $filename"
		$mail.attachments(1).saveasfile($attachment_dir + $filename)
	}

# building list of mails in a folder
Write-Host "$("[INFO] Now building list of mails to ") $data_dir $("phishme_data.csv ...")" -NoNewline
$files = Get-ChildItem "$attachment_dir" -Filter *.msg
foreach ($file in $files){    
    $msg = $outlook.CreateItemFromTemplate($file.FullName)
    $msg | select SentOn, SenderEmailAddress, SenderName, To, cc, Subject |Export-Csv -Path $($data_dir + "phishme_data.csv") -Delimiter ";" -Append
    }
Write-Host "Done" -ForegroundColor Green


# prevent autoclosure
Write-Host -NoNewLine 'Press any key to close...'
if (!$psISE) {$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')}