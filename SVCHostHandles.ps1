<#------------------------------------------------------------------------------
    Jason McClary
    mcclarj@mail.amc.edu
    19 Jul 2016
    20 Jul 2016 - Added email formating
    
    
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
set-variable alertValue -option Constant -value "17000"
set-variable warnValue -option Constant -value "10000"
set-variable runAsTask -option Constant -value $FALSE


$fromemail = "AMC_EDM_Checks@mail.amc.edu"
$smtpServer = "smtp.amc.edu"
$Emailuser = "ISEDMNotification@mail.amc.edu"


<#------------------------------------------------------------------------------
                                FUNCTIONS
------------------------------------------------------------------------------#>

    
<#------------------------------------------------------------------------------
                                    MAIN
------------------------------------------------------------------------------#>

## Format arguments from none, list or text file
IF ($runAsTask) {
    $compNames = get-content "E:\Server_Checks\EDM_Servers.txt"
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
                <td width=40% valign=center >
                    <p>Server Name</p>
                </td>
                <td width=20%>
                    <p>Handles</p>
                </td>
                <td width=20%>
                    <p>ID</p>
                </td>
                <td width=20%>
                    <p>Process Name</p>
                </td>
            </tr>"}
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
        FOREACH ($i in $array) 
        { 
            if($max -le $i.handles)
            { 
                $output = $i 
                $max = $i.handles
                IF ($max -ge $alertValue) {
                    $txtColor = "Red"
                    $redAlert = $TRUE
                }
                ELSEIF ($max -ge $warnValue){
                    $warnAlert = $TRUE
                    IF ($runAsTask) {$txtColor = "DarkOrange"}
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
        $line = "$server`t$handles`t$id`t$procName"
        IF ($runAsTask) {
            $mailMessage += "<tr  style=`"color: $txtColor`">
                <td width=40% valign=center >
                    <p>$server</p>
                </td>
                <td width=20%>
                    <p>$handles</p>
                </td>
                <td width=20%>
                    <p>$id</p>
                </td>
                <td width=20%>
                    <p>$procName</p>
                </td>
            </tr>"}
        ELSE {write-host $line -foregroundcolor $txtColor}
    } ELSE { # If no connection
        IF ($runAsTask) {
            $mailMessage += "<tr>
                <td width=40% valign=center span=4>
                    <p>$compName        ****** Could not connect ******</p>
                </td>
            </tr>"}
        ELSE {
            $line = "$compName        ****** Could not connect ******"
            $line = $line.PadRight(70)
            Write-Host $line -foregroundcolor "Red" -BackgroundColor "Black"
        }
        # $redAlert = $TRUE # send email if no connection.  ACUSISFAX1 sends false positives turning this off
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
    }
} ELSE {
    $line = "`n***************************  END OF LINE  ***************************`n`n"
    #$line = $line.PadRight(70)
    Write-Host $line -foregroundcolor "Blue" -BackgroundColor "Black"
}
