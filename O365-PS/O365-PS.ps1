#
# Script.ps1

function Get-ExtendedProperty
{
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Exchange.WebServices.Data.Item]
        $Item,

        # Param2 help descriptio
        [ValidateNotNullOrEmpty()]
        [Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition]
        $Property
    )

    $Value = $null
    $Succeeded = $Item.TryGetProperty($Property, [ref]$Value)

    if($Succeeded)
    {
        $Value
    }
    else
    {
        Write-Warning ("Could not get value for " + [System.Convert]::ToString($Property.Tag,16))
    }
}

function Check-Action($bytearray, [REF] $isPotentiallyMalicious, [REF] $actionType, [REF] $actionCommand)
{


    $isPotentiallyMalicious.Value = $false
    $actionType.Value = ""
    $actionCommand.Value = ""


    if ($bytearray.Length -gt 30 )
    {
      if ($bytearray[2] -eq 1 -and $bytearray[14] -eq 5 )
      {
        if ($bytearray[29] -eq 0x14) 
        {
         $actionType.Value = "ID_ACTION_CUSTOM"
         $isPotentiallyMalicious.Value = $true

        } 
        elseif ($bytearray[29] -eq 0x1e) 
        {
         $actionType.Value = "ID_ACTION_EXECUTE"
         $isPotentiallyMalicious.Value = $true

        }
         elseif ($bytearray[29] -eq 0x20) 
        {
         $actionType.Value = "ID_ACTION_RUN_MACRO"
         $isPotentiallyMalicious.Value = $true
        }
      }
      
      if ($isPotentiallyMalicious.Value -eq $true)
      {
        foreach ($byte in $bytearray) 
        {
            if ($byte -gt 31 -and $byte -lt 127)
            {
                $returnstring += [char]$byte
            }
        }

        $actionCommand.Value  = $returnstring
      }
    }
}

function Convert-ByteArrays ($bytearray)
{
    return [Convert]::ToBase64String($bytearray)
}

#Connecting with MFA
#Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter #Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName | ?{ $_ -notmatch "_none_" } | #select -First 1)
#Import-PSSession $Session -AllowClobber
	
Connect-ExchangeOnline
$Session = New-PSSessionOption

While (1) {

    Write-Host ""
    Write-Host "This script is used to run various tenant searches.  Only one can be done at a time."
    Write-Host "Select one of the following:"
    Write-Host “     multi         - Run multi-month front end logs in bulk with logging and geo” 
    Write-Host "     disable       - Disable IMAP and POP on all existing and future mailboxes”    
    Write-Host "     rules         - Run tenant search for newly created mailbox rules"
    Write-Host "     check         - Check mailbox audit log settings"
    Write-Host "     trace         - Recover Forwarded Messages with Impersonation using Message Trace Logs "
    Write-Host "     forwarding    - Run tenant search for newly created SMTP forwarding"
    Write-Host "     ips           - Run tenant search against list of IP Addresses"
    Write-Host "     freetext      - Run tenant search using the FreeText parameter.  Can be used to search for CBAInPROD or other text strings."
    Write-Host "     legacy        - Search tenant for mailboxes with IMAP and/or POP3 enabled"
    Write-Host "     quit          - Quit"
    Write-Host ""
    $selection = Read-Host "Enter selection"
    

if ($selection -ne “trace” -and $selection -ne “disable” -and $selection -ne “check” -and $selection -ne “multi” -and $selection -ne "ips" -and $selection -ne "rules" -and $selection -ne "forwarding" -and $selection -ne "freetext" -and $selection -ne "legacy" -and $selection -ne "quit")
    {
        Write-Host "You didn't enter a valid selection.  Try again."
        break
    }

    if ($selection -eq "quit") {
        Exit
        Stop-Transcript
        break
    }

if ($selection -eq "disable") {

Write-Host "Disabling IMAP and POP for any future mailboxes" -ForegroundColor DarkYellow
Get-CASMailboxPlan -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" } | set-CASMailboxPlan -ImapEnabled $false -PopEnabled $false
$confirmPlans = Get-CASMailboxPlan -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" }
if (!$confirmPlans) {
    Write-Host "IMAP and POP disabled for any future mailboxes" -ForegroundColor Green
}
else {
    Write-Host "IMAP and POP not disabled for any future mailboxes" -ForegroundColor Red
}
 
Write-Host "Disabling IMAP and POP on all existing mailboxes" -ForegroundColor DarkYellow
Get-CASMailbox -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" } | Select-Object @{n = "Identity"; e = {$_.primarysmtpaddress}} | Set-CASMailbox -ImapEnabled $false -PopEnabled $false
$confirmMailboxes = Get-CASMailbox -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" }
if (!$confirmMailboxes) {
    Write-Host "IMAP and POP disabled on all existing mailboxes`n" -ForegroundColor Green
}
else {
   Write-Host "IMAP and POP not disabled for all existing mailboxes" -ForegroundColor Red
} 
}


if ($selection -eq "check") {

While (1) {

    Write-Host
    $custodian = Read-Host ‘Enter custodian email address or Q to quit’

    if ($custodian -eq "Q") {
        
        break  
    } 

    Write-Host
    Write-Host $custodian “Account - Audited Admin Actions"
    Get-Mailbox $custodian | Select-Object -ExpandProperty AuditAdmin
    Write-Host " "
    Write-Host $custodian “Account - Audited Delegate Actions"
    Get-Mailbox $custodian | Select-Object -ExpandProperty AuditDelegate
    Write-Host " "
    Write-Host $custodian “Account - Audited Owner Actions"
    Get-Mailbox $custodian | Select-Object -ExpandProperty AuditOwner

   } 
}


if ($selection -eq "multi") {

Write-Host ""
$FilePath = Read-Host 'Enter the filename and location for the custodian list spreadsheet (C:\custodian_list.xlsx)'
$pathToGeoScript = Read-Host ‘Enter the name of the folder where the geoip.py python script is located (‪D:\Scripts\Misc)’‬‬‬‬‬‬‬‬‬‬‬‬
$driveletter = Read-Host ‘Enter the drive letter where you want the frontend logs to be stored (‪D:)’‬‬‬‬‬‬‬
$pathToFolder = $driveletter + "\O365_Frontend_Log_Pulls"
New-Item -ItemType Directory -Force -path $pathToFolder | out-null
$transcriptfile = $driveletter + "\O365_Frontend_Log_Pulls\log.txt"
Start-Transcript -Path $transcriptfile

$intervalMinutes = Read-Host "Enter the interval in minutes to want to use.  The smaller the interval the longer the log pull will take.  Unless you are getting the message 'Consider reducing the time interval', specify an interval of 1440"

#Open the excel spreadsheet
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false
$WorkBook = $objExcel.Workbooks.Open($FilePath)
$SheetName = "Sheet1"
$WorkSheet = $WorkBook.sheets.item($SheetName)

$Range = $Worksheet.UsedRange
$Rows = $Range.Rows

for ($inputcount = 1; $inputcount -le $Rows.Count; $inputcount = $inputcount + 1)
{
    $custodian = $Worksheet.Cells.Item($inputcount,1).Value()
	[DateTime]$start = $Worksheet.Cells.Item($inputcount,2).Value()
	[DateTime]$end = $Worksheet.Cells.Item($inputcount,3).Value()

    $startdate = Get-Date $start -Format MM-dd-yyyy
    $enddate = Get-Date $end -Format MM-dd-yyyy
    $outputFile = $pathToFolder +"\" + $custodian + "_" + $startdate + "_" + $enddate + ".csv"
    $xlsx = $pathToFolder +"\" + $custodian + "_" + $startdate + "_" + $enddate + ".xlsx"
    $end = $end.AddDays(1)
    
    Write-Host "Pulling frontend logs from " $start " to " $end " for " $custodian

    # 5000 max
    $resultSize = 5000
    #$intervalMinutes = 1440 
    # change to 3 if issues
    $retryCount = 0

    [DateTime]$currentStart = $start
    [DateTime]$currentEnd = $start
    $currentTries = 0
    $resultsTotal = 0
 
    while ($true)
    {
        $currentEnd = $currentStart.AddMinutes($intervalMinutes)
        if ($currentEnd -gt $end)
        {
            break
        }
        $currentTries = 0
        $sessionID = [DateTime]::Now.ToString().Replace('/', '_')
        Write-Host "INFO: Retrieving audit logs between $($currentStart) and $($currentEnd)"
        $currentCount = 0
        while ($true)
        {
            [Array]$results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -UserIds $custodian -SessionId $sessionID -SessionCommand ReturnNextPreviewPage -ResultSize $resultSize
            $resultsTotal = $resultsTotal + $results.Count

            if ($results -eq $null -or $results.Count -eq 0)
            {
                #Retry if needed. This may be due to a temporary network glitch
                if ($currentTries -lt $retryCount)
                {
                    $currentTries = $currentTries + 1
                    continue
                }
                else
                {
                    Write-Host "WARNING: Empty data set returned between $($currentStart) and $($currentEnd)."
                    break
                }
            }
            $currentTotal = $results[0].ResultCount
            if ($currentTotal -gt 5000)
            {
                Write-Host "WARNING: $($currentTotal) total records match the search criteria. Some records may get missed. Consider reducing the time interval."
                #return
                $end = 0
		        break
            }
            $currentCount = $currentCount + $results.Count
            Write-Host "INFO: Retrieved $($currentCount) records out of the total $($currentTotal)"
            $results | epcsv $outputFile -NoTypeInformation -Append
            if ($currentTotal -eq $results[$results.Count - 1].ResultIndex)
            {
                $message = "INFO: Successfully retrieved $($currentTotal) records for the current time range for " + $custodian + “ “ + $inputcount + ” of ” + $Rows.Count
		        Write-Host $message
                break
            }
        }
        $currentStart = $currentEnd
    }

    #Create a new Excel workbook with one empty sheet
    $objExcel1 = New-Object -ComObject Excel.Application
    $objExcel1.Visible = $false
    $WorkBook1 = $objExcel1.workbooks.Add(1)
    $WorkSheet1 = $WorkBook1.worksheets.Item(1)
    $WorkSheet1.name = "Analysis"

    If ($resultsTotal -ne 0) {

    $TxtConnector = ("TEXT;" + $outputFile)
    $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
    $query = $WorkSheet1.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = ","
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
    $query.AdjustColumnWidth = 1

    # Execute & delete the import query
    $query.Refresh() | out-null #writes True to console
    $query.Delete()

    $WorkSheet1.activate()
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161) 
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $objRange = $objExcel1.Range("H1").EntireColumn 
    [void] $objRange.Insert(-4161)
    $Worksheet1.Cells.Item(1,8).Value() = "ResultStatus"
    $Worksheet1.Cells.Item(1,9).Value() = "LogonError"
    $Worksheet1.Cells.Item(1,10).Value() = "KeepMeSignedIn"
    $Worksheet1.Cells.Item(1,11).Value() = "IP Address"
    $Worksheet1.Cells.Item(1,12).Value() = "City"
    $Worksheet1.Cells.Item(1,13).Value() = "State"
    $Worksheet1.Cells.Item(1,14).Value() = "Country Code"
    $Worksheet1.Cells.Item(1,15).Value() = "Country"
    $Worksheet1.Cells.Item(1,16).Value() = "ISP"
    $Worksheet1.Cells.Item(1,17).Value() = "Client/Agent"
    $Worksheet1.Cells.Item(1,18).Value() = "Inbox Rule"
    $Worksheet1.Cells.Item(1,19).Value() = "ForwardingSmtpAddress"

    $Range1 = $Worksheet1.UsedRange
    $Rows1 = $Range1.Rows
    $ipv4regex = ':"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
    $ipv4regexbracket = ':"\[\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
    #$ipv4regex = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"
    #$ipv6regex = "(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))"
    $ipv6regex = ":(?::[a-f\d]{1,4}){0,5}(?:(?::[a-f\d]{1,4}){1,2}|:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})))|[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}|:)|(?::(?:[a-f\d]{1,4})?|(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))))|:(?:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|[a-f\d]{1,4}(?::[a-f\d]{1,4})?|))|(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|:[a-f\d]{1,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){0,2})|:))|(?:(?::[a-f\d]{1,4}){0,2}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,3}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))"
    $resultstatusregex = '"ResultStatus":"[\w+\-=@.]*"'
    $logonerrorregex = '"LogonError":"[\w+\-=@.]*"'
    $clientagentregex = '"UserAgent","Value":".*"'
    $clientagentregex1 = '"UserAgent":".*",'
    $clientagentregex2 = '"ClientInfoString":".*"'
    $inboxruleregex = '"Parameters":.*"'
    $forwardingregex = '"ForwardingSmtpAddress","Value":"smtp:.*"}'
    $keepmesignedinregex = '"KeepMeSignedIn","Value":".*"'

    for ($count = 2; $count -le $Rows1.Count; $count = $count + 1)
    {
        $rslabel = ""
        $rsvalue = ""
        $lelabel = ""
        $levalue = ""
        $calabel = ""
        $cavalue = ""
        $kslabel = ""
        $ksvalue = ""
        $irlabel = ""
        $irvalue = ""
        $flabel = ""
        $fvalue = ""
        $AuditData = $Worksheet1.Cells.Item($count,20).Value()
        if ($AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }) {
            $resultstatus = $AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }
            $rslabel, $rsvalue = $resultstatus.split(':')
            $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
        }
        else {
            $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
        }
        if ($AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }) {
            $logonerror = $AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }
            $lelabel, $levalue = $logonerror.split(':')
            $Worksheet1.Cells.Item($count,9).Value() = $levalue
        }
        else {
            $Worksheet1.Cells.Item($count,9).Value() = $levalue
        }
        if ($AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }) {
            $clientagent = $AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }
            $calabel, $cavalue = $clientagent -split '":'
            $cavalue = $cavalue.split('}')
            $Worksheet1.Cells.Item($count,17).Value() = $cavalue
        }
        else {
            if ($AuditData | Select-String -Pattern $clientagentregex1 | % { $_.Matches } | % { $_.Value }) {
                $clientagent = $AuditData | Select-String -Pattern $clientagentregex1 | % { $_.Matches } | % { $_.Value }
                $calabel, $cavalue = $clientagent -split '":'
                $cavalue = $cavalue.split(',')
                $Worksheet1.Cells.Item($count,17).Value() = $cavalue
            }
            else {
                if ($AuditData | Select-String -Pattern $clientagentregex2 | % { $_.Matches } | % { $_.Value }) {
                    $clientagent = $AuditData | Select-String -Pattern $clientagentregex2 | % { $_.Matches } | % { $_.Value }
                    $calabel, $cavalue = $clientagent -split '":'
                    $cavalue = $cavalue.split(',')
                    $Worksheet1.Cells.Item($count,17).Value() = $cavalue
                }
                else {
                    $Worksheet1.Cells.Item($count,17).Value() = $cavalue
                }
            }
        }
        if ($AuditData | Select-String -Pattern $inboxruleregex | % { $_.Matches } | % { $_.Value }) {
                $inboxrule = $AuditData | Select-String -Pattern $inboxruleregex | % { $_.Matches } | % { $_.Value }
                $irlabel, $irvalue = $inboxrule -split '"Parameters":'
                $irvalue = $irvalue -split ',"SessionId"'
                $Worksheet1.Cells.Item($count,18).Value() = $irvalue[0]
            }
            else {
                $Worksheet1.Cells.Item($count,18).Value() = $irvalue
            }
        if ($AuditData | Select-String -Pattern $forwardingregex | % { $_.Matches } | % { $_.Value }) {
                $forwarding = $AuditData | Select-String -Pattern $forwardingregex | % { $_.Matches } | % { $_.Value }
                $flabel, $fvalue = $forwarding -split 'smtp:'
                $fvalue = $fvalue.split('"}')
                $Worksheet1.Cells.Item($count,19).Value() = $fvalue
            }
            else {
                $Worksheet1.Cells.Item($count,19).Value() = $fvalue
            }

        if ($AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }) {
            $ipv4string = $AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }
            $Worksheet1.Cells.Item($count,11).Value() = $ipv4string.Substring(2)
        }
        else {
            # No IPv4 found, so checking for IPv6
            if ($AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }) {
                $Worksheet1.Cells.Item($count,11).Value() = $AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }
            }
            else {
                # IP may have brackets
                if ($AuditData | Select-String -Pattern $ipv4regexbracket | % { $_.Matches } | % { $_.Value }) {
                    $bracket = $AuditData | Select-String -Pattern $ipv4regexbracket | % { $_.Matches } | % { $_.Value }
                    $blabel, $bvalue  = $bracket -split ':"\['
                    $Worksheet1.Cells.Item($count,11).Value() = $bvalue
                }
                else {
                    # IP must be blank
                    $Worksheet1.Cells.Item($count,11).Value() = "N/A"
                }
            }
        }
        if ($AuditData | Select-String -Pattern $keepmesignedinregex | % { $_.Matches } | % { $_.Value }) {
            $keepmesignedin = $AuditData | Select-String -Pattern $keepmesignedinregex | % { $_.Matches } | % { $_.Value }
            $kslabel, $ksvalue = $keepmesignedin.split(':')
            $ksvalue = $ksvalue.split('}')
            $Worksheet1.Cells.Item($count,10).Value() = $ksvalue
        }
        else {
            $Worksheet1.Cells.Item($count,10).Value() = $ksvalue
        }   
    }

    $WorkSheet1 = $WorkBook1.worksheets.add()
    $WorkSheet1.name = "IPs"
    $WorkSheet1 = $WorkBook1.sheets.item("IPs")
    $WorkSheet1 = $WorkBook1.sheets.item("Analysis")

    $Range2 = $WorkSheet1.Range(“K1”).EntireColumn
    $Range2.Copy() | out-null
    $WorkSheet1 = $WorkBook1.worksheets.item("IPs")
    $Range2 = $WorkSheet1.Range(“A1”)
    $WorkSheet1.Paste($Range2) 

    $WorkSheet1.activate()
    $WorkSheet1.UsedRange.RemoveDuplicates(1)

    Write-Host "Geolocating IP addresses.  Please be patient."

    $IpListFile = $pathToFolder +"\" + $custodian + "_IPs.txt"
    $Range2 = $WorkSheet1.Range("A1").EntireColumn
    $Range2.Copy() | out-null
    $ips = Get-Clipboard -TextFormatType Text

    for ($count = 0; $ips[$count] -ne ""; $count = $count + 1)
    {
        if (($ips[$count]  -ne "::1") -and ($ips[$count]  -ne "127.0.0.1") -and ($ips[$count]  -ne "IP Address") -and ($ips[$count]  -ne "N/A")) {
            Add-Content -Path $IpListFile -Value $ips[$count]
        }
    }

    $IpListFileGeo = $pathToFolder +"\" + $custodian + "_IPs_geo.txt"
    $arg1 = $pathToGeoScript + "\geoip.py"
    $parms = $arg1, $IpListFile, $IpListFileGeo
    & python.exe @parms

    $WorkSheet1 = $WorkBook1.worksheets.add()
    $WorkSheet1.name = "IPs_Geolocated"
    $WorkSheet1 = $WorkBook1.sheets.item("IPs_Geolocated")

    $TxtConnector = ("TEXT;" + $IpListFileGeo)
    $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
    $query = $WorkSheet1.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = ","
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
    $query.AdjustColumnWidth = 1

    # Execute & delete the import query
    $query.Refresh() | out-null #writes True to console
    $query.Delete()

    $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
    $Rows2 = $WorkSheet1.range("A1").currentregion.rows.count
    $WorkSheet1.range("L1:L$Rows2").formula = "=INDEX(IPs_Geolocated!B:B,(MATCH(K:K,IPs_Geolocated!A:A,0)))"
    $WorkSheet1.range("M1:M$Rows2").formula = "=INDEX(IPs_Geolocated!C:C,(MATCH(K:K,IPs_Geolocated!A:A,0)))"
    $WorkSheet1.range("N1:N$Rows2").formula = "=INDEX(IPs_Geolocated!D:D,(MATCH(K:K,IPs_Geolocated!A:A,0)))"
    $WorkSheet1.range("O1:O$Rows2").formula = "=INDEX(IPs_Geolocated!E:E,(MATCH(K:K,IPs_Geolocated!A:A,0)))"
    $WorkSheet1.range("P1:P$Rows2").formula = "=INDEX(IPs_Geolocated!F:F,(MATCH(K:K,IPs_Geolocated!A:A,0)))"

    }

    $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
    $WorkSheet1.activate()
    $WorkBook1.SaveAs($xlsx,51)

    $WorkBook1.Close()
    $objExcel1.Quit()

    If ($resultsTotal -ne 0) {

        Remove-Item $IpListFile
        Remove-Item $IpListFileGeo
        Remove-Item $outputFile
    }
}

$WorkBook.Save()
$WorkBook.Close()
$objExcel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel) | Out-Null 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel1) | Out-Null 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook1) | Out-Null 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet1) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range1) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows1) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objRange) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Connector) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($query) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range2) | Out-Null

# no $ needed on variable name in Remove-Variable call
Remove-Variable objexcel
Remove-Variable WorkBook
Remove-Variable WorkSheet
Remove-Variable Range
Remove-Variable Rows
Remove-Variable objexcel1
Remove-Variable WorkBook1
Remove-Variable WorkSheet1
Remove-Variable Range1
Remove-Variable Rows1
Remove-Variable objRange
Remove-Variable Connector
Remove-Variable query
Remove-Variable Range2
Remove-Variable Rows2

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

#Remove-PSSession $Session
Stop-Transcript
}

if ($selection -eq "legacy") {

        $PathToFolder = Read-Host "Enter the location to store output file (C:\Users\username\Desktop)"
        $transcriptfile = $PathToFolder + "\log.txt"
        Start-Transcript -Path $transcriptfile -Append
        $FileName = "\MailboxPlans_LegacyProtocolsEnabled.csv"
        $outputFile = $PathToFolder + $FileName
        [Array]$results1 = Get-CASMailboxPlan -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" }
        $results1 | export-csv -force $outputFile
        $FileName = "\Mailboxes_LegacyProtocalsEnabled.csv"
        $outputFile = $PathToFolder + $FileName
        [Array]$results1 = Get- Get-EXOCASMailbox -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" }
        $results1 | export-csv -force $outputFile

        #Remove-PSSession $Session
    }

    if ($selection -eq "rules")
    {
        
        Write-Host ""
        $PathToFolder = Read-Host "Enter the location to store output file (C:\Users\username\Desktop)"
        $transcriptfile = $PathToFolder + "\log.txt"
        Start-Transcript -Path $transcriptfile -Append
        [DateTime]$start = Read-Host "Enter start date.  Cannot be more than 90 days in the past (1/01/20)"
        [DateTime]$end = Read-Host "Enter end date.  Date cannot be the same as start date (1/20/20)"
        $pathToGeoScript = Read-Host "Enter the name of the folder where the geoip python script is located (D:\Scripts)"
        $intervalMinutes = Read-Host "Enter the interval in minutes to want to use.  The smaller the interval the longer the log pull will take.  Unless you are getting the message 'Consider reducing the time interval', specify an interval of 1440"
        $end = $end.AddDays(1)
        $resultSize = 5000

        $FileName = "\NewInboxRules_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".csv"
        $outputFile = $PathToFolder + $FileName
        $xlsx = $pathToFolder +"\NewInboxRules_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".xlsx"
        Write-Host "INFO: Running tenant search for newly created inbox rules between $($start) and $($end). Search may take a few minutes.  Please be patient."
        # Shouldn't have more than 5000 results so just running the normal command
        #Search-UnifiedAuditLog -StartDate $start -EndDate $end -Operations *new-inbox* -ResultSize $resultSize | export-csv -force $outputFile
    
        [Array]$results1 = Search-UnifiedAuditLog -StartDate $start -EndDate $end -Operations *new-inbox* -ResultSize $resultSize
        $results1 | export-csv -force $outputFile

        if ($results1 -ne $null -and $results1.Count -ne 0) {

            #Create a new Excel workbook with one empty sheet
            $objExcel1 = New-Object -ComObject Excel.Application
            $objExcel1.Visible = $false
            $WorkBook1 = $objExcel1.workbooks.Add(1)
            $WorkSheet1 = $WorkBook1.worksheets.Item(1)
            $WorkSheet1.name = "Analysis"

            $TxtConnector = ("TEXT;" + $outputFile)
            $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
            $query = $WorkSheet1.QueryTables.item($Connector.name)
            $query.TextFileOtherDelimiter = ","
            $query.TextFileParseType  = 1
            $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
            $query.AdjustColumnWidth = 1

            # Execute & delete the import query
            $query.Refresh() | out-null #writes True to console
            $query.Delete()

            $WorkSheet1.activate()
            [void]$Worksheet1.Cells.Item(1,1).EntireRow.Delete() # Delete the first row
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161) 
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $Worksheet1.Cells.Item(1,8).Value() = "ResultStatus"
            $Worksheet1.Cells.Item(1,9).Value() = "LogonError"
            $Worksheet1.Cells.Item(1,10).Value() = "IP Address"
            $Worksheet1.Cells.Item(1,11).Value() = "City"
            $Worksheet1.Cells.Item(1,12).Value() = "State"
            $Worksheet1.Cells.Item(1,13).Value() = "Country Code"
            $Worksheet1.Cells.Item(1,14).Value() = "Country"
            $Worksheet1.Cells.Item(1,15).Value() = "ISP"
            $Worksheet1.Cells.Item(1,16).Value() = "Client/Agent"
            $Worksheet1.Cells.Item(1,17).Value() = "Inbox Rule"

            $Range1 = $Worksheet1.UsedRange
            $Rows1 = $Range1.Rows

            $ipv4regex = ':"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
            #$ipv4regex = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"
            #$ipv6regex = "(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))"
            $ipv6regex = ":(?::[a-f\d]{1,4}){0,5}(?:(?::[a-f\d]{1,4}){1,2}|:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})))|[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}|:)|(?::(?:[a-f\d]{1,4})?|(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))))|:(?:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|[a-f\d]{1,4}(?::[a-f\d]{1,4})?|))|(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|:[a-f\d]{1,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){0,2})|:))|(?:(?::[a-f\d]{1,4}){0,2}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,3}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))"
            $resultstatusregex = '"ResultStatus":"[\w+\-=@.]*"'
            $logonerrorregex = '"LogonError":"[\w+\-=@.]*"'
            $clientagentregex = '"UserAgent","Value":".*"'
            #$inboxruleregex = '"Parameters":[{'
            $inboxruleregex = '"Parameters":.*"'

            for ($count = 2; $count -le $Rows1.Count; $count = $count + 1)
            {
                $rslabel = ""
                $rsvalue = ""
                $lelabel = ""
                $levalue = ""
                $calabel = ""
                $cavalue = ""
                $irlabel = ""
                $irvalue = ""
                $AuditData = $Worksheet1.Cells.Item($count,18).Value()
                if ($AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }) {
                    $resultstatus = $AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }
                    $rslabel, $rsvalue = $resultstatus.split(':')
                    $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
                }
                else {
                    $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
                }
                if ($AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }) {
                    $logonerror = $AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }
                    $lelabel, $levalue = $logonerror.split(':')
                    $Worksheet1.Cells.Item($count,9).Value() = $levalue
                }
                else {
                    $Worksheet1.Cells.Item($count,9).Value() = $levalue
                }
                if ($AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }) {
                    $clientagent = $AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }
                    $calabel, $cavalue = $clientagent -split '":'
                    $cavalue = $cavalue.split('}')
                    $Worksheet1.Cells.Item($count,16).Value() = $cavalue
                }
                else {
                    $Worksheet1.Cells.Item($count,16).Value() = $cavalue
                }

                if ($AuditData | Select-String -Pattern $inboxruleregex | % { $_.Matches } | % { $_.Value }) {
                    $inboxrule = $AuditData | Select-String -Pattern $inboxruleregex | % { $_.Matches } | % { $_.Value }
                    $irlabel, $irvalue = $inboxrule -split '"Parameters":'
                    $irvalue = $irvalue -split ',"SessionId"'
                    $Worksheet1.Cells.Item($count,17).Value() = $irvalue[0]
                }
                else {
                    $Worksheet1.Cells.Item($count,17).Value() = $irvalue
                }

                if ($AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }) {
                    $ipv4string = $AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }
                    $Worksheet1.Cells.Item($count,10).Value() = $ipv4string.Substring(2)
                }
                else {
                    # No IPv4 found, so checking for IPv6
                    if ($AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }) {
                        $Worksheet1.Cells.Item($count,10).Value() = $AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }
                    }
                    else {
                        # IP must be blank
                        $Worksheet1.Cells.Item($count,10).Value() = "N/A"
                    }
                }   
            }

            $WorkSheet1 = $WorkBook1.worksheets.add()
            $WorkSheet1.name = "IPs"
            $WorkSheet1 = $WorkBook1.sheets.item("IPs")
            $WorkSheet1 = $WorkBook1.sheets.item("Analysis")

            $Range2 = $WorkSheet1.Range(“J1”).EntireColumn
            $Range2.Copy() | out-null
            $WorkSheet1 = $WorkBook1.worksheets.item("IPs")
            $Range2 = $WorkSheet1.Range(“A1”)
            $WorkSheet1.Paste($Range2) 

            $WorkSheet1.activate()
            $WorkSheet1.UsedRange.RemoveDuplicates(1)

            Write-Host "Geolocating IP addresses.  Please be patient."

            $IpListFile = $pathToFolder +"\" + $selection + "_IPs.txt"
            $Range2 = $WorkSheet1.Range("A1").EntireColumn
            $Range2.Copy() | out-null
            $ips = Get-Clipboard -TextFormatType Text

            for ($count = 0; $ips[$count] -ne ""; $count = $count + 1)
            {
                if (($ips[$count]  -ne "::1") -and ($ips[$count]  -ne "127.0.0.1") -and ($ips[$count]  -ne "IP Address") -and ($ips[$count]  -ne "N/A")) {
                    Add-Content -Path $IpListFile -Value $ips[$count]
                }
            }

            $IpListFileGeo = $pathToFolder +"\" + $selection + "_IPs_geo.txt"
            $arg1 = $pathToGeoScript + "\geoip.py"
            $parms = $arg1, $IpListFile, $IpListFileGeo
            & python.exe @parms

            $WorkSheet1 = $WorkBook1.worksheets.add()
            $WorkSheet1.name = "IPs_Geolocated"
            $WorkSheet1 = $WorkBook1.sheets.item("IPs_Geolocated")

            $TxtConnector = ("TEXT;" + $IpListFileGeo)
            $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
            $query = $WorkSheet1.QueryTables.item($Connector.name)
            $query.TextFileOtherDelimiter = ","
            $query.TextFileParseType  = 1
            $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
            $query.AdjustColumnWidth = 1

            # Execute & delete the import query
            $query.Refresh() | out-null #writes True to console
            $query.Delete()

            $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
            $Rows2 = $WorkSheet1.range("A1").currentregion.rows.count
            $WorkSheet1.range("K1:K$Rows2").formula = "=INDEX(IPs_Geolocated!B:B,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("L1:L$Rows2").formula = "=INDEX(IPs_Geolocated!C:C,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("M1:M$Rows2").formula = "=INDEX(IPs_Geolocated!D:D,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("N1:N$Rows2").formula = "=INDEX(IPs_Geolocated!E:E,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("O1:O$Rows2").formula = "=INDEX(IPs_Geolocated!F:F,(MATCH(J:J,IPs_Geolocated!A:A,0)))"

            $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
            $WorkSheet1.activate()
            $WorkBook1.SaveAs($xlsx,51)

            $WorkBook1.Close()
            $objExcel1.Quit()

            Remove-Item $IpListFile
            Remove-Item $IpListFileGeo
            Remove-Item $outputFile

            Write-Host "INFO: Please check results.  If they equal 5000, some records were not collected. Re-run the search with a reduced time frame."

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel1) | Out-Null 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook1) | Out-Null 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objRange) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Connector) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($query) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range2) | Out-Null

            # no $ needed on variable name in Remove-Variable call
            Remove-Variable objexcel1
            Remove-Variable WorkBook1
            Remove-Variable WorkSheet1
            Remove-Variable Range1
            Remove-Variable Rows1
            Remove-Variable objRange
            Remove-Variable Connector
            Remove-Variable query
            Remove-Variable Range2
            Remove-Variable Rows2

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        else {
            Write-Host "INFO: No results found for the time frame specified."
        }

    }

    if ($selection -eq "forwarding")
    {
        Write-Host ""
        $PathToFolder = Read-Host "Enter the location to store output file (C:\Users\username\Desktop)"
        $transcriptfile = $PathToFolder + "\log.txt"
        Start-Transcript -Path $transcriptfile -Append
        [DateTime]$start = Read-Host "Enter start date.  Can not be more than 90 days in the past (6/01/20)"
        [DateTime]$end = Read-Host "Enter end date.  Date can not be the same as start date (6/01/20)"
        $pathToGeoScript = Read-Host "Enter the name of the folder where the geoip python script is located (D:\Scripts)"
        $intervalMinutes = Read-Host "Enter the interval in minutes to want to use.  The smaller the interval the longer the log pull will take.  Unless you are getting the message 'Consider reducing the time interval', specify an interval of 1440"
        $end = $end.AddDays(1)
        $resultSize = 5000

        $FileName = "\SetMailboxActions_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".csv"
        $outputFile = $PathToFolder + $FileName
        $xlsx = $pathToFolder +"\SetMailBoxActions_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".xlsx"
        #Write-Host "Output File:" $outputFile
        Write-Host "INFO: Running tenant search for forwarding rules between $($start) and $($end). Search may take a few minutes.  Please be patient."
        # Shouldn't have more than 5000 results so just running the normal command
        #Search-UnifiedAuditLog -StartDate $start -EndDate $end -Operations *set-mailbox* -ResultSize $resultSize | export-csv -force $outputFile

        [Array]$results1 = Search-UnifiedAuditLog -StartDate $start -EndDate $end -Operations *set-mailbox* -ResultSize $resultSize
        $results1 | export-csv -force $outputFile

        if ($results1 -ne $null -and $results1.Count -ne 0) {

            #Create a new Excel workbook with one empty sheet
            $objExcel1 = New-Object -ComObject Excel.Application
            $objExcel1.Visible = $false
            $WorkBook1 = $objExcel1.workbooks.Add(1)
            $WorkSheet1 = $WorkBook1.worksheets.Item(1)
            $WorkSheet1.name = "Analysis"

            $TxtConnector = ("TEXT;" + $outputFile)
            $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
            $query = $WorkSheet1.QueryTables.item($Connector.name)
            $query.TextFileOtherDelimiter = ","
            $query.TextFileParseType  = 1
            $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
            $query.AdjustColumnWidth = 1

            # Execute & delete the import query
            $query.Refresh() | out-null #writes True to console
            $query.Delete()

            $WorkSheet1.activate()
            [void]$Worksheet1.Cells.Item(1,1).EntireRow.Delete() # Delete the first row
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161) 
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $objRange = $objExcel1.Range("H1").EntireColumn 
            [void] $objRange.Insert(-4161)
            $Worksheet1.Cells.Item(1,8).Value() = "ResultStatus"
            $Worksheet1.Cells.Item(1,9).Value() = "LogonError"
            $Worksheet1.Cells.Item(1,10).Value() = "IP Address"
            $Worksheet1.Cells.Item(1,11).Value() = "City"
            $Worksheet1.Cells.Item(1,12).Value() = "State"
            $Worksheet1.Cells.Item(1,13).Value() = "Country Code"
            $Worksheet1.Cells.Item(1,14).Value() = "Country"
            $Worksheet1.Cells.Item(1,15).Value() = "ISP"
            $Worksheet1.Cells.Item(1,16).Value() = "Client/Agent"
            $Worksheet1.Cells.Item(1,17).Value() = "ForwardingSmtpAddress"

            $Range1 = $Worksheet1.UsedRange
            $Rows1 = $Range1.Rows
            $ipv4regex = ':"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
            #$ipv4regex = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"
            #$ipv6regex = "(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))"
            $ipv6regex = ":(?::[a-f\d]{1,4}){0,5}(?:(?::[a-f\d]{1,4}){1,2}|:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})))|[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}|:)|(?::(?:[a-f\d]{1,4})?|(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))))|:(?:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|[a-f\d]{1,4}(?::[a-f\d]{1,4})?|))|(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|:[a-f\d]{1,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){0,2})|:))|(?:(?::[a-f\d]{1,4}){0,2}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,3}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))"
            $resultstatusregex = '"ResultStatus":"[\w+\-=@.]*"'
            $logonerrorregex = '"LogonError":"[\w+\-=@.]*"'
            $clientagentregex = '"UserAgent","Value":".*"'
            $forwardingregex = '"ForwardingSmtpAddress","Value":"smtp:.*"}'
            #$forwardingregex = '"ForwardingSmtpAddress"'

            for ($count = 2; $count -le $Rows1.Count; $count = $count + 1)
            {
                $rslabel = ""
                $rsvalue = ""
                $lelabel = ""
                $levalue = ""
                $calabel = ""
                $cavalue = ""
                $flabel = ""
                $fvalue = ""
                $AuditData = $Worksheet1.Cells.Item($count,18).Value()
                if ($AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }) {
                    $resultstatus = $AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }
                    $rslabel, $rsvalue = $resultstatus.split(':')
                    $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
                }
                else {
                    $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
                }
                if ($AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }) {
                    $logonerror = $AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }
                    $lelabel, $levalue = $logonerror.split(':')
                    $Worksheet1.Cells.Item($count,9).Value() = $levalue
                }
                else {
                    $Worksheet1.Cells.Item($count,9).Value() = $levalue
                }
                if ($AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }) {
                    $clientagent = $AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }
                    $calabel, $cavalue = $clientagent -split '":'
                    $cavalue = $cavalue.split('}')
                    $Worksheet1.Cells.Item($count,16).Value() = $cavalue
                }
                else {
                    $Worksheet1.Cells.Item($count,16).Value() = $cavalue
                }

                if ($AuditData | Select-String -Pattern $forwardingregex | % { $_.Matches } | % { $_.Value }) {
                    $forwarding = $AuditData | Select-String -Pattern $forwardingregex | % { $_.Matches } | % { $_.Value }
                    $flabel, $fvalue = $forwarding -split 'smtp:'
                    $fvalue = $fvalue.split('"}')
                    $Worksheet1.Cells.Item($count,17).Value() = $fvalue
                }
                else {
                    $Worksheet1.Cells.Item($count,17).Value() = $fvalue
                }

                if ($AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }) {
                    $ipv4string = $AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }
                    $Worksheet1.Cells.Item($count,10).Value() = $ipv4string.Substring(2)
                }
                else {
                    # No IPv4 found, so checking for IPv6
                    if ($AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }) {
                        $Worksheet1.Cells.Item($count,10).Value() = $AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }
                    }
                    else {
                        # IP must be blank
                        $Worksheet1.Cells.Item($count,10).Value() = "N/A"
                    }
                }   
            }

            $WorkSheet1 = $WorkBook1.worksheets.add()
            $WorkSheet1.name = "IPs"
            $WorkSheet1 = $WorkBook1.sheets.item("IPs")
            $WorkSheet1 = $WorkBook1.sheets.item("Analysis")

            $Range2 = $WorkSheet1.Range(“J1”).EntireColumn
            $Range2.Copy() | out-null
            $WorkSheet1 = $WorkBook1.worksheets.item("IPs")
            $Range2 = $WorkSheet1.Range(“A1”)
            $WorkSheet1.Paste($Range2) 

            $WorkSheet1.activate()
            $WorkSheet1.UsedRange.RemoveDuplicates(1)

            Write-Host "Geolocating IP addresses.  Please be patient."

            $IpListFile = $pathToFolder +"\" + $selection + "_IPs.txt"
            $Range2 = $WorkSheet1.Range("A1").EntireColumn
            $Range2.Copy() | out-null
            $ips = Get-Clipboard -TextFormatType Text

            for ($count = 0; $ips[$count] -ne ""; $count = $count + 1)
            {
                if (($ips[$count]  -ne "::1") -and ($ips[$count]  -ne "127.0.0.1") -and ($ips[$count]  -ne "IP Address") -and ($ips[$count]  -ne "N/A")) {
                    Add-Content -Path $IpListFile -Value $ips[$count]
                }
            }

            $IpListFileGeo = $pathToFolder +"\" + $selection + "_IPs_geo.txt"
            $arg1 = $pathToGeoScript + "\geoip.py"
            $parms = $arg1, $IpListFile, $IpListFileGeo
            & python.exe @parms

            $WorkSheet1 = $WorkBook1.worksheets.add()
            $WorkSheet1.name = "IPs_Geolocated"
            $WorkSheet1 = $WorkBook1.sheets.item("IPs_Geolocated")

            $TxtConnector = ("TEXT;" + $IpListFileGeo)
            $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
            $query = $WorkSheet1.QueryTables.item($Connector.name)
            $query.TextFileOtherDelimiter = ","
            $query.TextFileParseType  = 1
            $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
            $query.AdjustColumnWidth = 1

            # Execute & delete the import query
            $query.Refresh() | out-null #writes True to console
            $query.Delete()

            $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
            $Rows2 = $WorkSheet1.range("A1").currentregion.rows.count
            $WorkSheet1.range("K1:K$Rows2").formula = "=INDEX(IPs_Geolocated!B:B,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("L1:L$Rows2").formula = "=INDEX(IPs_Geolocated!C:C,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("M1:M$Rows2").formula = "=INDEX(IPs_Geolocated!D:D,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("N1:N$Rows2").formula = "=INDEX(IPs_Geolocated!E:E,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
            $WorkSheet1.range("O1:O$Rows2").formula = "=INDEX(IPs_Geolocated!F:F,(MATCH(J:J,IPs_Geolocated!A:A,0)))"

            $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
            $WorkSheet1.activate()
            $WorkBook1.SaveAs($xlsx,51)

            $WorkBook1.Close()
            $objExcel1.Quit()

            Remove-Item $IpListFile
            Remove-Item $IpListFileGeo
            Remove-Item $outputFile

            Write-Host "INFO: Please check results.  If they equal 5000, some records were not collected. Re-run the search with a reduced time frame."

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel1) | Out-Null 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook1) | Out-Null 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows1) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objRange) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Connector) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($query) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range2) | Out-Null

            # no $ needed on variable name in Remove-Variable call
            Remove-Variable objexcel1
            Remove-Variable WorkBook1
            Remove-Variable WorkSheet1
            Remove-Variable Range1
            Remove-Variable Rows1
            Remove-Variable objRange
            Remove-Variable Connector
            Remove-Variable query
            Remove-Variable Range2
            Remove-Variable Rows2

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        else {
            Write-Host "INFO: No results found for the time frame specified."
        }
    }

    if ($selection -eq "ips")
    {
        Write-Host ""
        $PathToFolder = Read-Host "Enter the location to store output file (C:\Users\username\Desktop)"
        $transcriptfile = $PathToFolder + "\log.txt"
        Start-Transcript -Path $transcriptfile -Append
        [DateTime]$start = Read-Host "Enter start date.  Cannot be more than 90 days in the past (6/01/20)"
        [DateTime]$end = Read-Host "Enter end date.  Date cannot be the same as start date (6/01/20)"
        $pathToGeoScript = Read-Host "Enter the name of the folder where the geoip python script is located (D:\Scripts)"
        $intervalMinutes = Read-Host "Enter the interval in minutes to want to use.  The smaller the interval the longer the log pull will take.  Unless you are getting the message 'Consider reducing the time interval', specify an interval of 1440"
        $end = $end.AddDays(1)
        $resultSize = 5000

        $ipsfile = Read-Host 'Enter the location and filename that contains the list of IP addresses to search (C:\custodian_list.xlsx)'
        $SheetName = Read-Host 'Enter the name of the worksheet that contains the list of IP addresses to search (Bad_IPs)'

        #Open the excel spreadsheet
        $objExcel = New-Object -ComObject Excel.Application
        $objExcel.Visible = $false
        $WorkBook = $objExcel.Workbooks.Open($ipsfile)
        #$SheetName = "Sheet1"
        $WorkSheet = $WorkBook.sheets.item($SheetName)

        $Range = $Worksheet.UsedRange
        $Rows = $Range.Rows

        $iplist = ""

        for ($inputcount = 1; $inputcount -le $Rows.Count; $inputcount = $inputcount + 1)
        {
            $ip = $Worksheet.Cells.Item($inputcount,1).Value()
            $ipsplit = $ip.Split(".")
            $ip = $ipsplit[0] + "." + $ipsplit[1] + "." + $ipsplit[2] + ".*"
            if ($inputcount -eq $Rows.Count) {
                $iplist = $iplist + $ip
            }
            else {
                $iplist = $iplist + $ip + ","
            }
        }

        $FileName = "Bad_IPs_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".csv"
        $xlsx = $pathToFolder +"\Bad_IPs_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".xlsx"

        # 5000 max
        $resultSize = 5000
        # 1440 = 1 day, 10080 = 1 week, 43800 = 1 month
        #$intervalMinutes = 1440 
        # change to 3 if issues
        $retryCount = 0

        [DateTime]$currentStart = $start
        [DateTime]$currentEnd = $start
        $currentTries = 0
        $foundflag = ""
 
        while ($true)
        {
            $currentEnd = $currentStart.AddMinutes($intervalMinutes)
            if ($currentEnd -gt $end)
            {
                break
            }
            $currentTries = 0
            $sessionID = [DateTime]::Now.ToString().Replace('/', '_')
            Write-Host "INFO: Running tenant search for IP Addresses between $($currentStart) and $($currentEnd). Search may take a few minutes.  Please be patient."
            $currentCount = 0
            while ($true)
            {
                # -UserIds
                # -IPAddresses
                # -FreeText CBAInPROD
                # -Operations *new-inbox*
                # -Operations *-inbox*
                # -Operations *set-mailbox*
            
                [Array]$results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -IPAddresses $iplist -SessionId $sessionID -SessionCommand ReturnNextPreviewPage -ResultSize $resultSize

                if ($results -eq $null -or $results.Count -eq 0)
                {
                    #Retry if needed. This may be due to a temporary network glitch
                    if ($currentTries -lt $retryCount)
                    {
                        $currentTries = $currentTries + 1
                        continue
                    }
                    else
                    {
                        Write-Host "WARNING: Empty data set returned between $($currentStart) and $($currentEnd)."
                        break
                    }
                }
                $currentTotal = $results[0].ResultCount
                #if ($currentTotal -gt 5000)
                if ($currentTotal -gt $resultSize)
                {
                    Write-Host "WARNING: $($currentTotal) total records match the search criteria. Some records may get missed. Consider reducing the time interval."
                    return
                }
                $currentCount = $currentCount + $results.Count
                Write-Host "INFO: Retrieved $($currentCount) records out of the total $($currentTotal)"
                $outputFile = $PathToFolder + "\" + $FileName
                $results | epcsv $outputFile -NoTypeInformation -Append
                if ($currentTotal -eq $results[$results.Count - 1].ResultIndex)
                {
                    $message = "INFO: Successfully retrieved $($currentTotal) records for the current time range."
                    Write-Host $message
                    $foundflag = "found"
                    break
                }
            }
            $currentStart = $currentEnd
        }

        if ($foundflag -eq "found") {
        #Create a new Excel workbook with one empty sheet
        $objExcel1 = New-Object -ComObject Excel.Application
        $objExcel1.Visible = $false
        $WorkBook1 = $objExcel1.workbooks.Add(1)
        $WorkSheet1 = $WorkBook1.worksheets.Item(1)
        $WorkSheet1.name = "Analysis"

        $TxtConnector = ("TEXT;" + $outputFile)
        $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
        $query = $WorkSheet1.QueryTables.item($Connector.name)
        $query.TextFileOtherDelimiter = ","
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        # Execute & delete the import query
        $query.Refresh() | out-null #writes True to console
        $query.Delete()

        $WorkSheet1.activate()
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161) 
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $Worksheet1.Cells.Item(1,8).Value() = "ResultStatus"
        $Worksheet1.Cells.Item(1,9).Value() = "LogonError"
        $Worksheet1.Cells.Item(1,10).Value() = "IP Address"
        $Worksheet1.Cells.Item(1,11).Value() = "City"
        $Worksheet1.Cells.Item(1,12).Value() = "State"
        $Worksheet1.Cells.Item(1,13).Value() = "Country Code"
        $Worksheet1.Cells.Item(1,14).Value() = "Country"
        $Worksheet1.Cells.Item(1,15).Value() = "ISP"
        $Worksheet1.Cells.Item(1,16).Value() = "Client/Agent"

        $Range1 = $Worksheet1.UsedRange
        $Rows1 = $Range1.Rows
        $ipv4regex = ':"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
        #$ipv4regex = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"
        #$ipv6regex = "(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))"
        $ipv6regex = ":(?::[a-f\d]{1,4}){0,5}(?:(?::[a-f\d]{1,4}){1,2}|:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})))|[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}|:)|(?::(?:[a-f\d]{1,4})?|(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))))|:(?:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|[a-f\d]{1,4}(?::[a-f\d]{1,4})?|))|(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|:[a-f\d]{1,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){0,2})|:))|(?:(?::[a-f\d]{1,4}){0,2}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,3}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))"
        $resultstatusregex = '"ResultStatus":"[\w+\-=@.]*"'
        $logonerrorregex = '"LogonError":"[\w+\-=@.]*"'
        $clientagentregex = '"UserAgent","Value":".*"'

        for ($count = 2; $count -le $Rows1.Count; $count = $count + 1)
        {
            $rslabel = ""
            $rsvalue = ""
            $lelabel = ""
            $levalue = ""
            $calabel = ""
            $cavalue = ""
            $AuditData = $Worksheet1.Cells.Item($count,17).Value()
            if ($AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }) {
                $resultstatus = $AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }
                $rslabel, $rsvalue = $resultstatus.split(':')
                $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
            }
            else {
                $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
            }
            if ($AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }) {
                $logonerror = $AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }
                $lelabel, $levalue = $logonerror.split(':')
                $Worksheet1.Cells.Item($count,9).Value() = $levalue
            }
            else {
                $Worksheet1.Cells.Item($count,9).Value() = $levalue
            }
            if ($AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }) {
                $clientagent = $AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }
                $calabel, $cavalue = $clientagent -split '":'
                $cavalue = $cavalue.split('}')
                $Worksheet1.Cells.Item($count,16).Value() = $cavalue
            }
            else {
                $Worksheet1.Cells.Item($count,16).Value() = $cavalue
            }
            if ($AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }) {
                $ipv4string = $AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }
                $Worksheet1.Cells.Item($count,10).Value() = $ipv4string.Substring(2)
            }
            else {
                # No IPv4 found, so checking for IPv6
                if ($AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }) {
                    $Worksheet1.Cells.Item($count,10).Value() = $AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }
                }
                else {
                    # IP must be blank
                    $Worksheet1.Cells.Item($count,10).Value() = "N/A"
                }
            }   
        }

        $WorkSheet1 = $WorkBook1.worksheets.add()
        $WorkSheet1.name = "IPs"
        $WorkSheet1 = $WorkBook1.sheets.item("IPs")
        $WorkSheet1 = $WorkBook1.sheets.item("Analysis")

        $Range2 = $WorkSheet1.Range(“J1”).EntireColumn
        $Range2.Copy() | out-null
        $WorkSheet1 = $WorkBook1.worksheets.item("IPs")
        $Range2 = $WorkSheet1.Range(“A1”)
        $WorkSheet1.Paste($Range2) 

        $WorkSheet1.activate()
        $WorkSheet1.UsedRange.RemoveDuplicates(1)

        Write-Host "Geolocating IP addresses.  Please be patient."

        $IpListFile = $pathToFolder +"\" + $custodian + "_IPs.txt"
        $Range2 = $WorkSheet1.Range("A1").EntireColumn
        $Range2.Copy() | out-null
        $ips = Get-Clipboard -TextFormatType Text

        for ($count = 0; $ips[$count] -ne ""; $count = $count + 1)
        {
            if (($ips[$count]  -ne "::1") -and ($ips[$count]  -ne "127.0.0.1") -and ($ips[$count]  -ne "IP Address") -and ($ips[$count]  -ne "N/A")) {
                Add-Content -Path $IpListFile -Value $ips[$count]
            }
        }

        $IpListFileGeo = $pathToFolder +"\" + $custodian + "_IPs_geo.txt"
        $arg1 = $pathToGeoScript + "\geoip.py"
        $parms = $arg1, $IpListFile, $IpListFileGeo
        & python.exe @parms

        $WorkSheet1 = $WorkBook1.worksheets.add()
        $WorkSheet1.name = "IPs_Geolocated"
        $WorkSheet1 = $WorkBook1.sheets.item("IPs_Geolocated")

        $TxtConnector = ("TEXT;" + $IpListFileGeo)
        $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
        $query = $WorkSheet1.QueryTables.item($Connector.name)
        $query.TextFileOtherDelimiter = ","
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        # Execute & delete the import query
        $query.Refresh() | out-null #writes True to console
        $query.Delete()

        $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
        $Rows2 = $WorkSheet1.range("A1").currentregion.rows.count
        $WorkSheet1.range("K1:K$Rows2").formula = "=INDEX(IPs_Geolocated!B:B,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("L1:L$Rows2").formula = "=INDEX(IPs_Geolocated!C:C,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("M1:M$Rows2").formula = "=INDEX(IPs_Geolocated!D:D,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("N1:N$Rows2").formula = "=INDEX(IPs_Geolocated!E:E,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("O1:O$Rows2").formula = "=INDEX(IPs_Geolocated!F:F,(MATCH(J:J,IPs_Geolocated!A:A,0)))"

        $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
        $WorkSheet1.activate()
        $WorkBook1.SaveAs($xlsx,51)

        $WorkBook1.Close()
        $objExcel1.Quit()

        Remove-Item $IpListFile
        Remove-Item $IpListFileGeo
        Remove-Item $outputFile

        $WorkBook.Save()
        $WorkBook.Close()
        $objExcel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel) | Out-Null 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel1) | Out-Null 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook1) | Out-Null 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet1) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range1) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows1) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objRange) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Connector) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($query) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range2) | Out-Null

        # no $ needed on variable name in Remove-Variable call
        Remove-Variable objexcel
        Remove-Variable WorkBook
        Remove-Variable WorkSheet
        Remove-Variable Range
        Remove-Variable Rows
        Remove-Variable objexcel1
        Remove-Variable WorkBook1
        Remove-Variable WorkSheet1
        Remove-Variable Range1
        Remove-Variable Rows1
        Remove-Variable objRange
        Remove-Variable Connector
        Remove-Variable query
        Remove-Variable Range2
        Remove-Variable Rows2

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        }
        else {
            $results1 = ""
            $outputFile = $PathToFolder + "\" + $FileName
            $results1 >> $outputFile
            Write-Host "INFO: No results found for the time frame specified."
            $WorkBook.Save()
            $WorkBook.Close()
            $objExcel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel) | Out-Null 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null 
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows) | Out-Null

            # no $ needed on variable name in Remove-Variable call
            Remove-Variable objexcel
            Remove-Variable WorkBook
            Remove-Variable WorkSheet
            Remove-Variable Range
            Remove-Variable Rows

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }

        #Remove-PSSession $Session
    }

    if ($selection -eq "freetext")
    {
        Write-Host ""
        $PathToFolder = Read-Host "Enter the location to store output file (C:\Users\username\Desktop)"
        $transcriptfile = $PathToFolder + "\log.txt"
        Start-Transcript -Path $transcriptfile -Append
        [DateTime]$start = Read-Host "Enter start date.  Cannot be more than 90 days in the past (6/01/20)"
        [DateTime]$end = Read-Host "Enter end date.  Date cannot be the same as start date (6/01/20)"
        $pathToGeoScript = Read-Host "Enter the name of the folder where the geoip python script is located (D:\Scripts)"
        $intervalMinutes = Read-Host "Enter the interval in minutes to want to use.  The smaller the interval the longer the log pull will take.  Unless you are getting the message 'Consider reducing the time interval', specify an interval of 1440"
        $end = $end.AddDays(1)
        $resultSize = 5000

        $unformattedFreetext = Read-Host "Enter text to search"
        #$freetext = """ + $unformattedFreetext + """
        $freetext = $unformattedFreetext

        $FileName = $freetext + "_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".csv"
        $xlsx = $pathToFolder +"\" + $freetext + "_" + $start.tostring(“MM-dd-yyyy”) + "_" + $end.tostring(“MM-dd-yyyy”) + ".xlsx"

        # 5000 max
        $resultSize = 5000
        # 1440 = 1 day, 10080 = 1 week, 43800 = 1 month
        #$intervalMinutes = 1440 
        # change to 3 if issues
        $retryCount = 0

        [DateTime]$currentStart = $start
        [DateTime]$currentEnd = $start
        $currentTries = 0
        $foundflag = ""
 
        while ($true)
        {
            $currentEnd = $currentStart.AddMinutes($intervalMinutes)
            if ($currentEnd -gt $end)
            {
                break
            }
            $currentTries = 0
            $sessionID = [DateTime]::Now.ToString().Replace('/', '_')
            Write-Host "INFO: Running tenant search for $freetext between $($currentStart) and $($currentEnd). Search may take a few minutes.  Please be patient."
            $currentCount = 0
            while ($true)
            {
                # -UserIds
                # -IPAddresses
                # -FreeText CBAInPROD
                # -Operations *new-inbox*
                # -Operations *-inbox*
                # -Operations *set-mailbox*
            
                [Array]$results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -FreeText $freetext -SessionId $sessionID -SessionCommand ReturnNextPreviewPage -ResultSize $resultSize

                if ($results -eq $null -or $results.Count -eq 0)
                {
                    #Retry if needed. This may be due to a temporary network glitch
                    if ($currentTries -lt $retryCount)
                    {
                        $currentTries = $currentTries + 1
                        continue
                    }
                    else
                    {
                        Write-Host "WARNING: Empty data set returned between $($currentStart) and $($currentEnd)."
                        break
                    }
                }
                $currentTotal = $results[0].ResultCount
                #if ($currentTotal -gt 5000)
                if ($currentTotal -gt $resultSize)
                {
                    Write-Host "WARNING: $($currentTotal) total records match the search criteria. Some records may get missed. Consider reducing the time interval."
                    return
                }
                $currentCount = $currentCount + $results.Count
                Write-Host "INFO: Retrieved $($currentCount) records out of the total $($currentTotal)"
                $outputFile = $PathToFolder + "\" + $FileName
                $results | epcsv $outputFile -NoTypeInformation -Append
                if ($currentTotal -eq $results[$results.Count - 1].ResultIndex)
                {
                    $message = "INFO: Successfully retrieved $($currentTotal) records for the current time range."
                    Write-Host $message
                    $foundflag = "found"
                    break
                }
            }
            $currentStart = $currentEnd
        }

        if ($foundflag -eq "found") {
        #Create a new Excel workbook with one empty sheet
        $objExcel1 = New-Object -ComObject Excel.Application
        $objExcel1.Visible = $false
        $WorkBook1 = $objExcel1.workbooks.Add(1)
        $WorkSheet1 = $WorkBook1.worksheets.Item(1)
        $WorkSheet1.name = "Analysis"

        $TxtConnector = ("TEXT;" + $outputFile)
        $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
        $query = $WorkSheet1.QueryTables.item($Connector.name)
        $query.TextFileOtherDelimiter = ","
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        # Execute & delete the import query
        $query.Refresh() | out-null #writes True to console
        $query.Delete()

        $WorkSheet1.activate()
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161) 
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $objRange = $objExcel1.Range("H1").EntireColumn 
        [void] $objRange.Insert(-4161)
        $Worksheet1.Cells.Item(1,8).Value() = "ResultStatus"
        $Worksheet1.Cells.Item(1,9).Value() = "LogonError"
        $Worksheet1.Cells.Item(1,10).Value() = "IP Address"
        $Worksheet1.Cells.Item(1,11).Value() = "City"
        $Worksheet1.Cells.Item(1,12).Value() = "State"
        $Worksheet1.Cells.Item(1,13).Value() = "Country Code"
        $Worksheet1.Cells.Item(1,14).Value() = "Country"
        $Worksheet1.Cells.Item(1,15).Value() = "ISP"
        $Worksheet1.Cells.Item(1,16).Value() = "Client/Agent"

        $Range1 = $Worksheet1.UsedRange
        $Rows1 = $Range1.Rows
        $ipv4regex = ':"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
        #$ipv4regex = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"
        #$ipv6regex = "(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))"
        $ipv6regex = ":(?::[a-f\d]{1,4}){0,5}(?:(?::[a-f\d]{1,4}){1,2}|:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})))|[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}:(?:[a-f\d]{1,4}|:)|(?::(?:[a-f\d]{1,4})?|(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))))|:(?:(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|[a-f\d]{1,4}(?::[a-f\d]{1,4})?|))|(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|:[a-f\d]{1,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){0,2})|:))|(?:(?::[a-f\d]{1,4}){0,2}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,3}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))|(?:(?::[a-f\d]{1,4}){0,4}(?::(?:(?:(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(?:25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))|(?::[a-f\d]{1,4}){1,2})|:))"
        $resultstatusregex = '"ResultStatus":"[\w+\-=@.]*"'
        $logonerrorregex = '"LogonError":"[\w+\-=@.]*"'
        $clientagentregex = '"UserAgent","Value":".*"'

        for ($count = 2; $count -le $Rows1.Count; $count = $count + 1)
        {
            $rslabel = ""
            $rsvalue = ""
            $lelabel = ""
            $levalue = ""
            $calabel = ""
            $cavalue = ""
            $AuditData = $Worksheet1.Cells.Item($count,17).Value()
            if ($AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }) {
                $resultstatus = $AuditData | Select-String -Pattern $resultstatusregex | % { $_.Matches } | % { $_.Value }
                $rslabel, $rsvalue = $resultstatus.split(':')
                $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
            }
            else {
                $Worksheet1.Cells.Item($count,8).Value() = $rsvalue
            }
            if ($AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }) {
                $logonerror = $AuditData | Select-String -Pattern $logonerrorregex | % { $_.Matches } | % { $_.Value }
                $lelabel, $levalue = $logonerror.split(':')
                $Worksheet1.Cells.Item($count,9).Value() = $levalue
            }
            else {
                $Worksheet1.Cells.Item($count,9).Value() = $levalue
            }
            if ($AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }) {
                $clientagent = $AuditData | Select-String -Pattern $clientagentregex | % { $_.Matches } | % { $_.Value }
                $calabel, $cavalue = $clientagent -split '":'
                $cavalue = $cavalue.split('}')
                $Worksheet1.Cells.Item($count,16).Value() = $cavalue
            }
            else {
                $Worksheet1.Cells.Item($count,16).Value() = $cavalue
            }
            if ($AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }) {
                $ipv4string = $AuditData | Select-String -Pattern $ipv4regex | % { $_.Matches } | % { $_.Value }
                $Worksheet1.Cells.Item($count,10).Value() = $ipv4string.Substring(2)
            }
            else {
                # No IPv4 found, so checking for IPv6
                if ($AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }) {
                    $Worksheet1.Cells.Item($count,10).Value() = $AuditData | Select-String -Pattern $ipv6regex | % { $_.Matches } | % { $_.Value }
                }
                else {
                    # IP must be blank
                    $Worksheet1.Cells.Item($count,10).Value() = "N/A"
                }
            }   
        }

        $WorkSheet1 = $WorkBook1.worksheets.add()
        $WorkSheet1.name = "IPs"
        $WorkSheet1 = $WorkBook1.sheets.item("IPs")
        $WorkSheet1 = $WorkBook1.sheets.item("Analysis")

        $Range2 = $WorkSheet1.Range(“J1”).EntireColumn
        $Range2.Copy() | out-null
        $WorkSheet1 = $WorkBook1.worksheets.item("IPs")
        $Range2 = $WorkSheet1.Range(“A1”)
        $WorkSheet1.Paste($Range2) 

        $WorkSheet1.activate()
        $WorkSheet1.UsedRange.RemoveDuplicates(1)

        Write-Host "Geolocating IP addresses.  Please be patient."

        $IpListFile = $pathToFolder +"\" + $custodian + "_IPs.txt"
        $Range2 = $WorkSheet1.Range("A1").EntireColumn
        $Range2.Copy() | out-null
        $ips = Get-Clipboard -TextFormatType Text

        for ($count = 0; $ips[$count] -ne ""; $count = $count + 1)
        {
            if (($ips[$count]  -ne "::1") -and ($ips[$count]  -ne "127.0.0.1") -and ($ips[$count]  -ne "IP Address") -and ($ips[$count]  -ne "N/A")) {
                Add-Content -Path $IpListFile -Value $ips[$count]
            }
        }

        $IpListFileGeo = $pathToFolder +"\" + $custodian + "_IPs_geo.txt"
        $arg1 = $pathToGeoScript + "\geoip.py"
        $parms = $arg1, $IpListFile, $IpListFileGeo
        & python.exe @parms

        $WorkSheet1 = $WorkBook1.worksheets.add()
        $WorkSheet1.name = "IPs_Geolocated"
        $WorkSheet1 = $WorkBook1.sheets.item("IPs_Geolocated")

        $TxtConnector = ("TEXT;" + $IpListFileGeo)
        $Connector = $WorkSheet1.QueryTables.add($TxtConnector,$WorkSheet1.Range("A1"))
        $query = $WorkSheet1.QueryTables.item($Connector.name)
        $query.TextFileOtherDelimiter = ","
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,1 * $WorkSheet1.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        # Execute & delete the import query
        $query.Refresh() | out-null #writes True to console
        $query.Delete()

        $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
        $Rows2 = $WorkSheet1.range("A1").currentregion.rows.count
        $WorkSheet1.range("K1:K$Rows2").formula = "=INDEX(IPs_Geolocated!B:B,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("L1:L$Rows2").formula = "=INDEX(IPs_Geolocated!C:C,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("M1:M$Rows2").formula = "=INDEX(IPs_Geolocated!D:D,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("N1:N$Rows2").formula = "=INDEX(IPs_Geolocated!E:E,(MATCH(J:J,IPs_Geolocated!A:A,0)))"
        $WorkSheet1.range("O1:O$Rows2").formula = "=INDEX(IPs_Geolocated!F:F,(MATCH(J:J,IPs_Geolocated!A:A,0)))"

        $WorkSheet1 = $WorkBook1.worksheets.item("Analysis")
        $WorkSheet1.activate()
        $WorkBook1.SaveAs($xlsx,51)

        $WorkBook1.Close()
        $objExcel1.Quit()

        Remove-Item $IpListFile
        Remove-Item $IpListFileGeo
        Remove-Item $outputFile

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel1) | Out-Null 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook1) | Out-Null 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet1) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range1) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Rows1) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objRange) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Connector) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($query) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range2) | Out-Null

        # no $ needed on variable name in Remove-Variable call
        Remove-Variable objexcel1
        Remove-Variable WorkBook1
        Remove-Variable WorkSheet1
        Remove-Variable Range1
        Remove-Variable Rows1
        Remove-Variable objRange
        Remove-Variable Connector
        Remove-Variable query
        Remove-Variable Range2
        Remove-Variable Rows2

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        }
        else {
            $results1 = ""
            $outputFile = $PathToFolder + "\" + $FileName
            $results1 >> $outputFile
            Write-Host "INFO: No results found for the time frame specified."
        }

    }
}

