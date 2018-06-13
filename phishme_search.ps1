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

[int16]$lookback_time = Read-Host "Input lookup depth (hours)"
$date_from = (Get-date).AddHours(-$lookback_time)
Write-Host "$("Working with emails from") $date_from"
Write-Host "[INFO] Building mail database. It could take some time ... " -NoNewline
$mails = $folder.items | where-object { $_.ReceivedTime -gt $date_from}  
# $mails = $folder.items | where-object { $_.ReceivedTime -gt [DateTime]::ParseExact($date, 'd/M/yyyy HH:mm:ss',[CultureInfo]::InvariantCulture)} 
Write-Host "Done" -ForegroundColor Green

#get search string from user's input
$search_string = Read-Host "Enter the search string"

foreach($mail in $mails){

$tag = ($mail.HTMLBody) -split "`n" | sls $search_string

if ($tag) 
    {
    $filename = $mail.attachments(1).FileName
    Write-Host "$("[INFO] Saving mail ") $filename)"
    $mail.attachments(1).saveasfile($attachment_dir + $filename)
    }
}

Write-Host "$("[INFO] All mails were saved in ") $attachment_dir" -ForegroundColor Green

Write-Host -NoNewLine 'Press any key to close...'
if (!$psISE) {$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')}