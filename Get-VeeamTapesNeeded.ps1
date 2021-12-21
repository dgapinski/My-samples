<#
.NAME
    Get-VeeamTapesNeeded
.SYNOPSIS
    Figure out what tapes are needed this week and send that list out
.DESCRIPTION
    Figure out what tapes are needed this week and send that list out 
   
.PARAMETER <paramName>
   None (edit the variables in the script)
.EXAMPLE
   None (just run it)
#>

$now = get-date
$nowformatted = "{0:yyyy-MM-dd}" -f (get-date)

[int]$Script:WeeklyTapeNumberLimit = '12'
$Script:JobNameForOffsiteTapes = "Tape - Offsite - WKLY - 10am - 12 RP"
$Script:JobNameForOnsiteTapes = "Tape - Tosa - WKLY and DLY 7am"
$Script:LastWeeklyTapeUsed = ""
$sendEmail = $true
$global:smtpServer = "mail.contoso.com"
$emailUser = ""
$emailPass = ""

$Script:VaultTheseTapesOffsite = $null
$Script:VaultTheseTapesOnsite = $null
$Script:AddTheseDailyTapesToLibrary = $null
$Script:VaultTheseWeeklyTapesToLibrary = $null
$Script:CurrentOffsiteTapeInLibrary = $null

$global:FromAddress = "VeeamTapesNeeded@contoso.com"
$global:ToAddress = @("firstperson@contoso.com","secondperson@contoso.com")
$global:Subject = "Tape Swap for " + $nowformatted
$Global:EmailBody = ""

$VeeamServers= @("VeeamServerName")
Import-Module Veeam.Backup.PowerShell

################
#  FUNCTIONS   #
################

Function EmailReport (){
    
    Param
    (
    [String]$CurrentEmailAddress,
	[String]$Global:EmailBody
    )
    $msg = new-object Net.Mail.MailMessage
    $smtp = new-object Net.Mail.SmtpClient($global:smtpServer)
	$msg.IsBodyHTML = $true
	$msg.Subject = $global:Subject 
	$msg.Body = $Global:HTMLBody
	$msg.To.Add($CurrentEmailAddress)
    $msg.From = $global:FromAddress
    $smtp.Send($msg)
    Send-MailMessage -From
    }
    
Function VaultTheseTapesOffsite(){
    
    $emailmessage = $null
	foreach ($tape in $script:Tapes){
			if (($tape.location -notlike "none") -and ($tape.name -like "B*")){
				$Script:CurrentOffsiteTapeInLibrary = $tape.name 
                #if (($tape.lastwritetime -gt (get-date).AddDays(-14)) -and ($tape.lastwritetime -lt (get-date))) {
                if ($tape.lastwritetime -lt (get-date)) {
                    #Tape needs to be vaulted
                    $Script:VaultTheseTapesOffsite += $tape.name 
					$Message = $tape.name + " was written on " + $tape.lastwritetime + " and goes to the offsite vault"
                    $emailmessage += $tape.name + " : written on " + $tape.lastwritetime
                    
                }else{
                    #No change needed
					$Message = $tape.name + " was written on " + $tape.lastwritetime + " and can stay where it is at" 
                }
                
                write-host $message -ForegroundColor Yellow		
            }
            
	}	
    
    if ($emailmessage -eq $null){$emailmessage = "<B>No tapes to move to offsite vault</B>" + "<P>"} 
    else {$emailmessage = "<B>Please take out this tape and bring it to Brookfield:</B><BR>"+$emailmessage + "<P>"}

    return $emailmessage
}

Function AddThisOffsiteTapeToLibrary(){
    $emailmessage = $null
    $ThisWeeksNeededOffsiteTapes = get-vbrtapemedium | where {$_.Location -like 'none' -and ($_.name -like 'B*1L') -and ($_.ExpirationDate -gt $now) -and ($_.ExpirationDate -lt (get-date).AddDays(8))} | select name, Location, lastwritetime, ExpirationDate | sort ExpirationDate
    $emailmessage = "<B>Then move this weekly tape from the Brookfield safe to the library:</B>" 
    foreach ($ThisWeeksNeededOffsiteTape in $ThisWeeksNeededOffsiteTapes ){
        write-host  $ThisWeeksNeededOffsiteTape.name " expires on " $ThisWeeksNeededOffsiteTape.ExpirationDate " and moves from the Brookfield safe to the library." -ForegroundColor Yellow
        $emailmessage += "<BR>" + $ThisWeeksNeededOffsiteTape.name + " : expires on " +  $ThisWeeksNeededOffsiteTape.ExpirationDate  
    } 
    $emailmessage += "<P>"
    return $emailmessage
}

Function VaultTheseTapesOnsite(){
    $emailmessage = $null
    
    $InsertedTapes = get-vbrtapemedium | where {$_.Location -notlike 'none' -and ($_.name -like 'D*' -or $_.name -like 'W*') -and ($_.lastwritetime -gt (get-date).AddDays(-7)) -and ($_.lastwritetime -lt $now)} | select name, Location, lastwritetime, ExpirationDate | sort lastwritetime
    $emailmessage = "<B>Then move these tapes from the library to the Wauwatosa safe:</B>" 
    foreach ($InsertedTape in $InsertedTapes ){
        if ($InsertedTape.name -like 'W*'){$Script:LastWeeklyTapeUsed = $InsertedTape.name}
        write-host  $InsertedTape.name " was written on " $InsertedTape.lastwritetime " and moves from the library to the Tosa safe." -ForegroundColor Yellow
        $emailmessage += "<BR>" + $InsertedTape.name + " : written on " +  $InsertedTape.lastwritetime  
    } 
    $emailmessage += "<P>"
    return $emailmessage
}

Function AddTheseDailyTapesToLibrary(){
    $emailmessage = $null
    $ThisWeeksNeededDailyTapes =  get-vbrtapemedium | where {($_.Location -notlike 'Drive*' -or $_.Location -notlike 'Slot*')  -and ($_.name -like 'D*') -and ($_.ExpirationDate -lt (get-date).AddDays(8)) } | select name, Location, lastwritetime, ExpirationDate | sort ExpirationDate
    #$ThisWeeksNeededDailyTapes =  get-vbrtapemedium | where {(($_.Location -like 'Tosa*') -or ($_.Location -like 'none') -or ($_.Location -like 'Vault')) -and ($_.name -like 'D*') -and ($_.ExpirationDate -lt (get-date).AddDays(8)) } | select name, Location, lastwritetime, ExpirationDate | sort ExpirationDate
    $emailmessage = "<B>Then move these daily tapes from the Wauwatosa safe to the library:</B>" 
    foreach ($ThisWeeksNeededDailyTape in $ThisWeeksNeededDailyTapes ){
        write-host  $ThisWeeksNeededDailyTape.name " expires on " $ThisWeeksNeededDailyTape.ExpirationDate " and moves from the Tosa safe to the library." -ForegroundColor Yellow
        $emailmessage += "<BR>" + $ThisWeeksNeededDailyTape.name + " : expires on " +  $ThisWeeksNeededDailyTape.ExpirationDate  
    } 
    $emailmessage += "<P>"
    return $emailmessage

}

Function AddTheseWeeklyTapesToLibrary(){
    $emailmessage = $null
    $ThisWeeksNeededWeeklyTapes = get-vbrtapemedium | where {($_.Location -notlike 'Drive*' -or $_.Location -notlike 'Slot*') -and ($_.name -like 'W*1L7') -and ($_.ExpirationDate -gt (get-date).AddDays(6))-and ($_.ExpirationDate -lt (get-date).AddDays(12))} | select name, Location, lastwritetime, ExpirationDate | sort ExpirationDate
    #$ThisWeeksNeededWeeklyTapes = get-vbrtapemedium | where {($_.Location -like 'Tosa*' -or $_.Location -like 'none' -or $_.Location -like 'Vault') -and ($_.name -like 'W*1L7') -and ($_.ExpirationDate -gt (get-date).AddDays(6))-and ($_.ExpirationDate -lt (get-date).AddDays(12))} | select name, Location, lastwritetime, ExpirationDate | sort ExpirationDate
    $emailmessage = "<B>Then move this weekly tape from the Wauwatosa safe to the library:</B>" 
    foreach ($ThisWeeksNeededWeeklyTape in $ThisWeeksNeededWeeklyTapes ){
        write-host  $ThisWeeksNeededWeeklyTape.name " expires on " $ThisWeeksNeededWeeklyTape.ExpirationDate " and moves from the Tosa safe to the library." -ForegroundColor Yellow
        $emailmessage += "<BR>" + $ThisWeeksNeededWeeklyTape.name + " : expires on " +  $ThisWeeksNeededWeeklyTape.ExpirationDate  
    } 
    $emailmessage += "<P>"
    return $emailmessage

}


##################
#  ENTRY POINT   #
##################

$Global:HTMLBody = "<HTML><BODY>"
$Global:HTMLBody += "<face=""calibri"">"

foreach ($VeeamServer in $VeeamServers){
write-host $VeeamServer -ForegroundColor green
disconnect-VBRServer | out-null
connect-vbrserver -server $VeeamServer #-Credential $cred

$Script:Tapes = Get-VBRTapeMedium 

write-host  "From library to offsite vault" -ForegroundColor Cyan
$Global:HTMLBody += VaultTheseTapesOffsite + "<P>"
write-host  "From offsite vault to library" -ForegroundColor Cyan
$Global:HTMLBody += AddThisOffsiteTapeToLibrary + "<P>"
write-host  "From library to onsite vault" -ForegroundColor Cyan
$Global:HTMLBody += VaultTheseTapesOnsite + "<P>"
write-host  "From onsite vault to library - daily" -ForegroundColor Cyan
$Global:HTMLBody += AddTheseDailyTapesToLibrary + "<P>"
write-host  "From onsite vault to library - weekly" -ForegroundColor Cyan
$Global:HTMLBody += AddTheseWeeklyTapesToLibrary + "<P>"
}


$Global:HTMLBody += "<B>Then bask in the glory of fighting the unheroic fight, until someone has a disaster and you're a BIG hero!</B></font></body></html>"


#Let's email that report now too...
write-host "Sending report mail..." -ForegroundColor DarkGreen
Foreach ($mailaddress in $global:ToAddress ){

    Send-MailMessage -From $global:FromAddress -To $mailaddress -subject $global:Subject -BodyAsHtml $Global:HTMLBody -SmtpServer $global:smtpServer
}
write-host "All done!" -ForegroundColor Green
