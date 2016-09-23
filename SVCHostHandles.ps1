<#------------------------------------------------------------------------------
    Jason McClary
    mcclarj@mail.amc.edu
    19 Jul 2016
    20 Jul 2016 - Added email formating
    29 Aug 2016 - Changed the Warn and Alert values:
                          OLD         NEW
                  WARN    8000        11000
                  ALERT   11000       20000
                - Fixed formatting issue with could not connect
                - Formatted the list with color rows
                - Added Intro text about "could not connect" and thresholds
    08 Sep 2016 - Added Free Memory and Uptime columns
    09 Sep 2016 - Added Total System handle count
    
    
    Description:
    Query SVCHost Handle counts
    
    Arguments:
    If blank script runs against local computer
    Multiple computer names can be passed as a list separated by spaces:
        SVCHostHandles.ps1 computer1 computer2 anotherComputer
    A text file with a list of computer names can also be passed
        SVCHostHandles.ps1 comp.txt
        
    Tasks:
    - Create a file that lists highest handle count per server

        
--------------------------------------------------------------------------------
                                CONSTANTS
------------------------------------------------------------------------------#>

set-variable procToCheck -option Constant -value "svchost"
set-variable alertValue -option Constant -value "20000"
set-variable warnValue -option Constant -value "11000"
set-variable runAsTask -option Constant -value $TRUE
set-variable sendEmail -option Constant -value $TRUE


$fromemail = "AMC_EDM_Checks@mail.amc.edu"
$smtpServer = "smtp.amc.edu"
#$Emailuser = "ISEDMNotification@mail.amc.edu", "RohrwaR@mail.amc.edu"
$Emailuser = "RohrwaR@mail.amc.edu", "mcclarj@mail.amc.edu"
$bkgndRowColor = "#F1F1F1"
$alertColor = "Red"
$warnColor = "DarkOrange"


<#------------------------------------------------------------------------------
                                FUNCTIONS
------------------------------------------------------------------------------#>

    
<#------------------------------------------------------------------------------
                                    MAIN
------------------------------------------------------------------------------#>

## Format arguments from none, list or text file
IF ($runAsTask) {
    $compNames = get-content "D:\Server_Checks\EDM_Servers.txt"
    #$compNames = get-content "C:\Users\mcclarj\Desktop\Server_Info\EDM_Servers.txt"
}
ELSE {
    IF (!$args){
        $compNames = $env:computername # Get the local computer name
    } ELSE {
        $passFile = Test-Path $args

        IF ($passFile -eq $True) {
            $compNames = get-content $args
        } ELSE {
            $compNames = $args
        }
    }
}

## Initialize Variables
$redAlert = $FALSE
$warnAlert = $FALSE
$mailMessage = ""
$screenDisplay = ""



## Format Powershell Header
$server = "Server Name"
$server = $server.PadRight(20)
$handles = "Handles"
$handles = $handles.PadRight(10)
$id = "ID"
$id = $id.PadRight(10)
$procName = "Process Name"
$procName = $procName.PadRight(10)
$header = "$server`t$handles`t$id`t$procName"
IF ($runAsTask) {
    $mailMessage += "<html>
    <head>
    </head>
    <body style='font-family:`"Courier New`"'>
        <table border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'>
            <tr>
                <td valign=center colspan=7>
                    <p><i>This check only tries once to connect to a server. `"Could not connect`" errors are most likely false positives unless there were previous handle warnings.</i><br><br> </p>
                </td>
            </tr>
            <tr>
                <td valign=center colspan=7>
                    <p>&emsp;Current thresholds:<br>
                    <code>&emsp;&emsp;&emsp;&emsp;<font color = `"$warnColor`">Warning = &emsp;$warnValue</font><br>
                          &emsp;&emsp;&emsp;&emsp;<font color = `"$alertColor`">Alert &emsp;&emsp;= &emsp;$alertValue</font><br><br></code> </p>
                </td>
            </tr>
            <tr><u><b>
                <td width=20% valign=center >
                    <p>Server Name</p>
                </td>
                <td width=10%>
                    <p>Handles</p>
                </td>
                <td width=10%>
                    <p>ID</p>
                </td>
                <td width=10%>
                    <p>Process Name</p>
                </td>
                <td width=10%>
                    <p>Free Mem</p>
                </td>
                <td width=10%>
                    <p>Total Handles</p>
                </td>
                <td width=30%>
                    <p>Uptime</p>
                </td>
            </b></u></tr>"}
ELSE {write-host $header}

$server = "-----------"
$server = $server.PadRight(20)
$handles = "-------"
$handles = $handles.PadRight(10)
$id = "--"
$id = $id.PadRight(10)
$procName = "------------"
$procName = $procName.PadRight(10)
$header = "$server`t$handles`t$id`t$procName"

IF (!$runAsTask) {write-host $header}



FOREACH ($compName in $compNames) {
    IF(Test-Connection -BufferSize 16 -count 1 -quiet $compName){  # Check for valid connection to computer

        $array = Get-Process $procToCheck -ComputerName $compName | select @{LABEL='Server';EXPRESSION={$compName}}, @{LABEL='Handles';EXPRESSION={$_.handles}}, @{LABEL='ID';EXPRESSION={$_.Id}}, @{LABEL='Process';EXPRESSION={$_.ProcessName}}
        $max = 0
        # If multiple services with same name (ex. SVCHost) find the one with the highest count
        FOREACH ($i in $array) 
        { 
            if($max -le $i.handles)
            { 
                $output = $i 
                $max = $i.handles
                IF ($max -ge $alertValue) {
                    $txtColor = $alertColor
                    $redAlert = $TRUE
                }
                ELSEIF ($max -ge $warnValue){
                    $warnAlert = $TRUE
                    IF ($runAsTask) {$txtColor = $warnColor}
                    ELSE {$txtColor = "Yellow"}
                }
                ELSE {
                    IF ($runAsTask) {$txtColor = "Black"}
                    ELSE {$txtColor = "White"}
                }
            } 
        }
        $server = $output.Server.PadRight(20)
        $handles = ($output.handles -as [string]).PadRight(10)
        $id = ($output.Id -as [string]).PadRight(10)
        $procName = $output.Process.PadRight(10)

        $upTime = (Get-WmiObject -ComputerName $compName win32_operatingsystem | select CSName, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}})
        $upTime =  NEW-TIMESPAN -Start $upTime.lastbootuptime -End ([DateTime]::Now)      
        $uptime = "$($upTime.Days)d $($uptime.Hours)hrs $($uptime.Minutes)min"

        $totalHandles = (Get-Counter -Counter "\\$compName\Process(_total)\Handle Count").CounterSamples
        $totalHandles = $totalHandles[0].CookedValue

        $freeMem = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $compName
        $freeMem = ([math]::round(100 - (((($freemem.TotalVisibleMemorySize - $freemem.FreePhysicalMemory) / $freemem.TotalVisibleMemorySize)) * 100), 0))


        $line = "$server`t$handles`t$id`t$procName"
        IF ($runAsTask) {
            $mailMessage += "<tr  style=`"color: $txtColor;background-color:$bkgndRowColor;`">
                <td width=15% valign=center >
                    <p>$server</p>
                </td>
                <td width=10%>
                    <p>$handles</p>
                </td>
                <td width=10%>
                    <p>$id</p>
                </td>
                <td width=10%>
                    <p>$procName</p>
                </td>
                <td width=10%>
                    <p>$freeMem%</p>
                </td>
                <td width=10%>
                    <p>$totalHandles</p>
                </td>
                <td width=35%>
                    <p>$upTime</p>
                </td>
            </tr>"}
        ELSE {write-host $line -foregroundcolor $txtColor}
    } ELSE { # If no connection
        IF ($runAsTask) {
            $mailMessage += "<tr  style=`"color: $alertColor;background-color:$bkgndRowColor;`">
                <td width=15% valign=center>
                    <p>$compName</p>
                </td>
                <td width=85% valign=center colspan=6>
                    <p>****** Could not connect ******</p>
                </td>
            </tr>"}
        ELSE {
            $line = "$compName        ****** Could not connect ******"
            $line = $line.PadRight(70)
            Write-Host $line -foregroundcolor "Red" -BackgroundColor "Black"
        }
        $redAlert = $TRUE # send email if no connection.  ACUSISFAX1 sends false positives turning this off
    }
    # Alternate the row background color
    IF ($bkgndRowColor -eq "#FFF") {
        $bkgndRowColor = "#F1F1F1"
    } ELSE {
        $bkgndRowColor = "#FFF"
    }
}

IF ($runAsTask) { # when run as task send the email
    $mailMessage += "        </table>
    </body>
</html>"

    IF ($redAlert) {
        $Subject_Text = "Handle Count Alert!"
        foreach ($users in $EmailUser) {
            send-mailmessage -from $fromemail -to $users -subject $Subject_Text -BodyAsHTML -body $mailMessage -priority High -smtpServer $smtpServer
        }
    } ELSEIF ($warnAlert){
        $Subject_Text = "Handle Count Warning!"
        foreach ($users in $EmailUser) {
            send-mailmessage -from $fromemail -to $users -subject $Subject_Text -BodyAsHTML -body $mailMessage -priority Normal -smtpServer $smtpServer
        }
    } ELSEIF ($sendEmail){
        $Subject_Text = "Handle Count Info"
        foreach ($users in $EmailUser) {
            send-mailmessage -from $fromemail -to $users -subject $Subject_Text -BodyAsHTML -body $mailMessage -priority Normal -smtpServer $smtpServer
        }
    }
} ELSE {
    $line = "`n***************************  END OF LINE  ***************************`n`n"
    #$line = $line.PadRight(70)
    Write-Host $line -foregroundcolor "Blue" -BackgroundColor "Black"
}
