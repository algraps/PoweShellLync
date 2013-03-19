param(
  [string]$UserEmail,
  [string]$ExportEmail
)
#********************************************************************************************************************************************
#Author: Alessandro Graps
#Date : 2013-03-19
#********************************************************************************************************************************************
#Start Function
#********************************************************************************************************************************************
function UnZipMe($zipfilename,$destination)
{
   $shellApplication = new-object -com shell.application
   $zipPackage = $shellApplication.NameSpace($zipfilename)
   $destinationFolder = $shellApplication.NameSpace($destination)
	# CopyHere vOptions Flag # 4 - Do not display a progress dialog box.
	# 16 - Respond with "Yes to All" for any dialog box that is displayed.
	$destinationFolder.CopyHere($zipPackage.Items(),20)
}

function ZipMe($srcdir,$ZipFileName,$zipFilepath )
{
		$zipFile = "$zipFilepath$zipFilename"
		#Prepare zip file
		if(-not (test-path($zipFile))) {
		    set-content $zipFile ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
		    (dir $zipFile).IsReadOnly = $false  
		}

		$shellApplication = new-object -com shell.application
		$zipPackage = $shellApplication.NameSpace($zipFile)
		$files = Get-ChildItem -Path $srcdir | where{! $_.PSIsContainer}
		
		foreach($file in $files) { 
		    $zipPackage.CopyHere($file.FullName)
			#using this method, sometimes files can be 'skipped'
			#this 'while' loop checks each file is added before moving to the next
		    while($zipPackage.Items().Item($file.name) -eq $null){
		        Start-sleep -seconds 1
		    }
		}
}

Function IsValidEmail   
{ 
	Param ([string] $In) 
	# Returns true if In is in valid e-mail format. 
	[system.Text.RegularExpressions.Regex]::IsMatch($In,  
              "^(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +  
              "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$");  
}

Function UpdateContacts($user)
{
	Write-Host "Processing - $email_new"
	#Tmp dir
	$folderTmp = "C:\TmpLync\"
	$folderTmpContact = "C:\TmpLync\TmpContacts"
	$zipFilename = "contacts.zip"
	$applicationPath = "C:\"

	#Delete dir
	Write-Host "Remove tempFolder $folderTmp"
	Remove-Item -Recurse -Force $folderTmp

	#********************************************************************************************************************************************
	#Start TMP FOLDER
	#********************************************************************************************************************************************
	##Clean and create all tmp dir
	if (!(Test-Path $folderTmp)) {
		# create it
		Write-Host "Create tempFolder $folderTmp"
		[void](new-item $folderTmp -itemType directory)
	}
	if (!(Test-Path $folderTmpContact)) {
		# create it
		Write-Host "Create tempFolder $folderTmpContact"
		[void](new-item $folderTmpContact -itemType directory)
	}
	#********************************************************************************************************************************************
	#End TMP FOLDER
	#********************************************************************************************************************************************
	#********************************************************************************************************************************************
	#Start UNZIP CONTACTS
	#********************************************************************************************************************************************
	##Unzip FILE
	Write-Host "Unzip file $applicationPath$zipFilename"
	$a = gci -Path $applicationPath -Filter $zipFilename
	##
	foreach($file in $a)
	{
	    Write-Host "Processing - $file" UnZipMe –zipfilename
	    UnZipMe –zipfilename $file.FullName -destination $folderTmpContact 
	}
	#********************************************************************************************************************************************
	#End UNZIP CONTACTS
	#********************************************************************************************************************************************
	#********************************************************************************************************************************************
	#Start XML MANIPOLATION
	#********************************************************************************************************************************************
	Write-Host "Start xml manipolation $folderTmpContact\DocItemSet.xml"
	$xml = New-Object XML
	$xml = [xml](get-content "$folderTmpContact\DocItemSet.xml")
	$emailXml = $xml.DocItemSet.DocItem[0].name
	$emails =$emailXml.Split(":")
	$email = $emails[2]
	$original_file = $folderTmpContact + "\DocItemSet.xml"
	$destination_file =  $folderTmpContact + "\DocItemSet.xml"
	(Get-Content $original_file) | Foreach-Object {
	    $_ -replace $email, $email_new      
	    } | Set-Content $destination_file
	Write-Host "Replace $email with $email_new"
	Write-Host "End xml manipolation $folderTmpContact\DocItemSet.xml"
	#********************************************************************************************************************************************
	#END XML MANIPOLATION
	#********************************************************************************************************************************************
	#********************************************************************************************************************************************
	#Start ZIP FILE
	#********************************************************************************************************************************************
	#ZIP FILE
	#Zip File src
	Write-Host "Create new file zip $ZipFileName"
	$srcdir = $folderTmpContact
	$zipFilepath = $applicationPath
	#Delete old contacts.zip
	$zipFile = $zipFilepath + $ZipFileName
	Remove-Item -Recurse -Force $zipFile
	ZipMe -srcdir $srcdir -zipFileName $ZipFileName -zipFilePath $zipFilepath 
	Write-Host "UpdateLync $ZipFileName"
	Update-CsUserData -FileName $zipFile -Confirm:$False
}

#********************************************************************************************************************************************
#End Function
#********************************************************************************************************************************************
Write-Host "Start processing"
Write-Host $UserEmail
Write-Host $ExportEmail

$email_new = $UserEmail

Write-Host "Start Import module"
Import-Module 'C:\Program Files\Common Files\Microsoft Lync Server 2013\Modules\Lync\Lync.psd1'
Write-Host "End Import module"


$zipFilename = "contacts.zip"
$applicationPath = "C:\"
$zipfile = $applicationPath + $zipFilename

if ((Test-Path $zipfile)) {
	# create it	
	Remove-Item -Recurse -Force $zipFile
}

Export-CsUserData -FileName $zipfile –UserFilter $ExportEmail -PoolFqdn "lync.server"

if ($email_new -eq "all")
{
	$users = Get-CsUser -Filter {RegistrarPool -eq "lync.server"} | select sipaddress
	foreach ($user in $users)
	{
		$usermail = $user.Split(":")
   		UpdateContacts -$user $usermail[1]
	}
}
else
{
	if (IsValidEmail($email_new))
	{
		UpdateContacts -$user $email_new
	}
	else
	{
		Write-Host "Invalid parameter"
	}
}

#********************************************************************************************************************************************
#End ZIP FILE
#********************************************************************************************************************************************
Write-Host "End processing"