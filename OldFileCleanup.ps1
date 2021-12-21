<#
.NAME 
    oldfilecleanup.ps1 
.SYNOPSIS
    Deletes old files based on number of days old

.DESCRIPTION
    Deletes old files based on number of days old

    Instructions: Modify the necessary variables under the     
    section "SET VARIABLES" below (lines 29-46). Especially    
    in the LogAreasToCleanArray variable
   
.PARAMETER <paramName>
   None (See above instructions)
.EXAMPLE
   None (just run it)
#>


############################################
#   SET LOCATION TO WHERE THE SCRIPT IS    #
############################################

#we do this to tell Powershell that we any relative directory references
$scriptpath = $MyInvocation.MyCommand.Path
write-host "Scriptpath = " $scriptpath
$dir = Split-Path $scriptpath
write-host "Dir = " $dir 
set-location $dir

#####################
#  USER VARIABLES   #
#####################

$Date = Get-Date -format s | foreach {$_ -replace "/", "-"} | foreach {$_ -replace ":", "."} 
$LogOutputPath = $($dir + ".\Log")
$global:Logfile = $($LogOutputPath + "\" + $date + '-FilesDeleted.txt')
$smtpServer = "mail.contoso.com"
$FromAddress = "LogCleanup@contoso.com"
[string[]]$DestMailboxes = @("<firstperson@contoso.com>","<secondperson@contoso.com>","<thirdperson@contoso.com>") 
#$TestingMode = "ON" #Comment this line for testing mode, which just shows a listing of old files
$global:ErrorCount=0
$global:SizeDeleted=0


$LogAreasToCleanArray = @(
    #Path,#OfDaysToKeep,FilesToInclude,FilesToExclude(optional)
    ('D:\location','30','*')
)

#-----------------------------------------------------------------------------------#
#                                                                                   # 
#	DO NOT EDIT ANY OF THE FOLLOWING CODE - EDITS ARE TO BE MADE ABOVE HERE ONLY!!  #
#                                                                                   #
#-----------------------------------------------------------------------------------#

###########################
# MAKE A VARIABLE DUSTBIN #
###########################

# Store all the start up variables so you can clean up when the script finishes.
if ($startupvariables) { try {Remove-Variable -Name startupvariables  -Scope Global -ErrorAction SilentlyContinue } catch { } }
New-Variable -force -name startupVariables -value ( Get-Variable | ForEach-Object { $_.Name } ) 

#######################
# MAKE OUTPUT FOLDERS #
#######################

if ((Test-Path -LiteralPath $LogOutputPath ) -eq $true) {
	write-output 'Log output dir found' 
}
else
{
	write-output "No log output dir found - making a new one" 
	New-Item -ItemType directory -Path $LogOutputPath
}

#####################
#   START LOGGING   #
#####################

Write-host $LogOutputPath
Write-host $Logfile 

Start-Transcript -path $Logfile 

#Let's start by writing the mode to the log first so it shows a strong cue about 
#whether to expect files being fake deleted or deleted for real...
if ($TestingMode) 
    {
	Write-output ""
	Write-output "#####################"
	Write-output "#   TESTING MODE    #"
	Write-output "#####################"
	Write-output ""
	}
	else
	{
	Write-output ""
	Write-output "#####################"
	Write-output "#  PRODUCTION MODE  #"
	Write-output "#####################"
	Write-output ""
	}

#####################
#   MAIN FUNCTION   #
#####################

Function LogDelete 
{
Param(
    [String]$LocationToClean,
    [String]$Limit,
    [String]$FilesToInclude,
    [String]$FilesToExclude
    )
	
	#if these parameters are missing, STOP the script and exit...
    if (-not ($LocationToClean)) { Throw "PARAMETER 1 Missing: You need a location to clean when calling the LogDelete function!"} 
    if (-not ($Limit)) { Throw "PARAMETER 2 Missing: You need a date limit when calling the LogDelete function!"} 
    if (-not ($FilesToInclude)) { Throw "PARAMETER 3 Missing: You need to specify files when calling the LogDelete function!"} 
		
	$DayLimit = (get-date).AddDays(-$Limit)

	$CleaningFilesLine = "LocationToClean: " + $LocationToClean 
	$OlderThanLine = "DayLimit: " + $daylimit.datetime
	$FilesIncludedLine = "FilesToInclude: " + $FilesToInclude 
	$FilesExcludedLine = "FilesToExclude: " + $FilesToExclude
	Write-output $CleaningFilesLine 
	Write-output $OlderThanLine 
	Write-output $FilesIncludedLine 
	Write-output $FilesExcludedLine 
	
	
	#this picks out the files that are older than the limit, and will feed that list to the foreach loop coming up
	Write-Output "Picking files to delete..." 
	#$FilesToDelete = @(Get-ChildItem -path $LocationToClean -Recurse -Include $FilesToInclude -Exclude $FilesToExclude | Where-object {!$_.psiscontainer -and $_.CreationTime -lt $DayLimit } )
	$FilesToDelete = @(Get-ChildItem $LocationToClean -Recurse -Include $FilesToInclude -Exclude $FilesToExclude) # Now the piped command doesn't work anymore
	if ($FilesToDelete.length -eq 0) {write-host "   no files to delete."}
	
	#Delete a newer file - this is here to force an error and test whether the script will attach the transcript to email
	#if ($locationtoclean -eq "D:\TitanFTP\srtFtpData\c1netpws01\apadmin\SSL_VPNBackup"){
	#write-host "Forcing an error..." -foregroundcolor red
	#remove-item "D:\TitanFTP\srtFtpData\c1netpws01\apadmin\SSL_VPNBackup\JuniperAccessLog-CHWSSLVPN-CHW-SSLVPN-C1-20161101-2303 - Copy.gz" -force}
	
	#loop through each file in the list for deleting or fake deleting (testing mode)
	foreach ($CurrentFile in $FilesToDelete) 
	{
		#write-host $currentfile | fl
		$lastWrite = (get-item $CurrentFile).LastWriteTime
		$isfolder =  (Get-Item $CurrentFile) -is [System.IO.DirectoryInfo]
		#write-host $lastWrite -foregroundcolor magenta
		
		if (($isfolder -ne "true") -and ($lastWrite -lt $DayLimit)){
		
			#get the file size to tab up
			$currentfilesize = $currentfile.length/1KB
			#write-host "file size is " $currentfilesize -foregroundcolor magenta `r`n
			$Global:SizeDeleted = $Global:SizeDeleted + $currentfilesize
			#write-host "Cumulative space deleted is " $Global:SizeDeleted  "KB" -foregroundcolor yellow `r`n
			
			if ($TestingMode) 
			{#Testing mode is ON - the line is NOT commented in the "Set Variables" section
			$DeletedMessageSuccess = "FAKE Deleting " + $CurrentFile.Name + " ...DONE!"
			Write-Output $DeletedMessageSuccess 
            if ((Test-Path -LiteralPath $CurrentFile ) -ne $true) {
                $Global:ErrorCount = $Global:ErrorCount + 1
                Write-host "Error count = " $Global:ErrorCount	-foregroundcolor RED
                }
            } 
			else 
			{
			#This is the real deal. Testing mode is OFF - the line is commented out in the "Set Variables" section.
			$DeletedMessageSuccess = "Deleting... " + $CurrentFile.Name 
			Remove-item $CurrentFile -force #-WhatIf 
			$ok = $? 
			write-output $ok
			#$ok tests to make sure the file deletion actually worked
			if ($ok -ne "true")
			    {
			        #The file deletion did NOT work, so let's update the log with great failure
			        $DeletedMessageFailure = $DeletedMessageSuccess + " ...FAILED!"
			        Write-Output $DeletedMessageFailure 
			        $Global:ErrorCount = $Global:ErrorCount + 1
			    } 
			else 
			    {
			        #The file deletion worked, so let's update the log with great success
			        $DeletedMessageSuccess = $DeletedMessageSuccess + " ...DONE!"
			        Write-Output $DeletedMessageSuccess 
			    }
			}
		}
		else
		{
		#In case we want it to log this info, my guess is no.
		#$SkipMessage = "Skipping " + $CurrentFile.Name 
		#Write-Output $SkipMessage
		}
	}
    
}

#####################
#   MAIL FUNCTION   #
#####################

Function EmailReport (){
    
    Param
    (
    [String]$CurrentEmailAddress
    )
    $msg = new-object Net.Mail.MailMessage
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$att = new-object Net.Mail.Attachment($global:LogFile)
	write-host "Attachment is: " $att -foregroundcolor magenta
    $msg.IsBodyHTML = $true
	
	Write-host "Error count = " $Global:ErrorCount	-foregroundcolor RED
	$DisplaySizeDeleted = ($global:SizeDeleted/1024) -as [int]
	
	if($Global:ErrorCount -ne 0)
		{
			#We have errors, & will send the transcript with a different subject and body
			if ($TestingMode) 
			{
				$msg.Subject = "TESTING - Logging Cleaned on " + $env:computername + " - ERRORS FOUND, see attached"
				$msg.Body = '<font face="Calibri"><br>Please see the attachment showing deleted log files for today.<p> ' + $DisplaySizeDeleted + ' MB would have been cleaned out, if only you commented out the variable TestingMode!</font>'				
				$msg.Attachments.Add($att)
			}
			else
			{
				$msg.Subject = "Logging Cleaned on " + $env:computername + " - ERRORS FOUND, see attached"
				$msg.Body = '<font face="Calibri"><br>Please see the attachment showing deleted log files AND ERRORS for today - that a file did not delete might be a real problem, so check it out!<p>' + $DisplaySizeDeleted + ' MB cleaned out.</font>'
				$msg.Attachments.Add($att)
			}
			
			
			
		} 
		else 
		{	
			#No errors, so not attaching the transcript
			if ($TestingMode) 
			{
				$msg.Subject = "TESTING - Logging Cleaned on " + $env:computername + " "
			}
			else
			{
				$msg.Subject = "Logging Cleaned on " + $env:computername + " "
			}
			$msg.Body = '<font face="Calibri"><br>Log files cleaned successfully!<p> ' + $DisplaySizeDeleted + ' MB cleaned out.</font>'
			#$att = new-object Net.Mail.Attachment($global:LogFile)
		}
	
    $msg.To.Add($CurrentEmailAddress)
    $msg.From = $FromAddress
    $smtp.Send($msg)
    if($Global:ErrorCount -ne 0){$att.Dispose()}
    }

#################################
#   VARIABLE CLEANUP FUNCTION   #
#################################

Function Clean-Memory {
Get-Variable |
 Where-Object { $startupVariables -notcontains $_.Name } |
 ForEach-Object {
  try { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue}
  catch { }
 }
}

#####################
#   MAIN PROCESS    #
#####################

#OK the script has everything defined at this point. NOW the real order of operations begins...

#Loop through the LogAreasToCleanArray variable and call the LogDelete function for each line there. 
foreach ($CurrentLogArea in $LogAreasToCleanArray)
{
	#Call the LogDelete function to remove files. Parameters: Path,#OfDaysToKeep,FilesToInclude,FilesToExclude
	$CurrentPath = $CurrentLogArea[0] 
	$CurrentNoDaysToKeep = $CurrentLogArea[1]  
	$CurrentFilesToInclude = $CurrentLogArea[2] 
	$CurrentFilesToExclude = $CurrentLogArea[3] 
	
	Write-output ""
	Write-output "#####################"
	Write-output ""
	Write-output "Cleaning the path:"
	Write-output $CurrentPath
	Write-output ""
	Write-output "#####################"
	Write-output ""
	
	#If exclusions don't exist in the array, don't call the function with a null value
	if (-not ($CurrentFilesToExclude)) 
		{ 
			LogDelete $CurrentPath $CurrentNoDaysToKeep $CurrentFilesToInclude 
		}
		else
		{
			LogDelete $CurrentPath $CurrentNoDaysToKeep $CurrentFilesToInclude $CurrentFilesToExclude
		}
}

#Stop logging
stop-transcript

#Emails the transcript out
Foreach ($mailaddress in $DestMailboxes){EmailReport $mailaddress}

#Might as well have this script clean up after itself!
LogDelete $LogOutputPath "30" "*"


#Remove the variable data from memory. Ending the powershell session should kill it anyway, but in case it doesn't...
Clean-Memory