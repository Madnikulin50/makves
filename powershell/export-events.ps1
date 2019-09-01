
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] $computers = ("acme.local"),
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $outfilename = "events",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $Count = 3000,
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $user = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $pwd = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $start = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $fwd = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $syslog = "",
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $syslog_port = 514,
    [Parameter(Mandatory = $False, Position = 10, ParameterSetName = "NormalRun")] [ValidateSet("All","Logon","Service","User","Computer", "Clean", "File", "MSSQL", "RAS", "USB", "Printer", "Sysmon", "TS")] [array]$target="All"
)


$EventLevel="All"
$NumberOfLastEventsToGet = $Count
$EventLogName = ("Security")



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"

Write-Host "computers: " $computers
Write-Host "outfilename: " $server
Write-Host "target: " $target
Write-Host "EventSource: " $EventSource
Write-Host "EventLevel: " $EventLevel
Write-Host "EventLogName: " $EventLogName

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}


if ($user -ne "") {
    $pass = ConvertTo-SecureString -AsPlainText $pwd -Force    
    $GetAdminact = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pass    
} else {
    $GetAdminact = Get-Credential
}


$stopwatch = [system.diagnostics.stopwatch]::StartNew()
$DebugPreference = "Continue"
$ErrorActionPreference = "SilentlyContinue"

$FilterHashProperties = $null
$Answer = ""



if ($syslog -ne "") {
    Write-Host "syslog: " $syslog
    Import-Module (Join-Path $PSScriptRoot "Send-SyslogMessage.ps1")   
   
}

function SendSyslog($event) {
    try {
        Send-SyslogMessage -Server $syslog -Message $event.Message -Severity "Informational" -Facility "logaudit" -Hostname $env:COMPUTERNAME -ApplicationName "EventLog" -MessageID $event.Id    
        # -Timestamp $event.TimeGenerated
    } catch {
        Write-Host ($event | Co)
        $msg = "Error send to syslog $PSItem.Exception.InnerExceptionMessage"
        Write-Host $msg
    }
}

function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}


Write-Host "Starting script..."
function ExportFor($eid, $ln, $type) {

    if ($fwd -ne "") {
        $ln = $fwd
    }        

    Write-Host "logname:" $ln
    Write-Host "type: " $type
   
    

    $FilterHashProperties = @{
        LogName = $ln
    }
    
    if ($start -ne "") {
        Write-Host "start: " $start
        $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
        $FilterHashProperties.Add("startTime", $starttime)
    }

    If (!(IsEmpty $eid)){
        Write-Host "eid: " $eid
        $FilterHashProperties.Add("ID",$eid)
    }
 
      
    $msg = ("About to collect events on $($computers.count)") + $(If ($($computers.count) -gt 1){" machines"}Else{" machine"})
    Write-host $msg
    
    Foreach ($computer in $computers)
    {
        $msg = "Checking Computer $Computer"
        Write-host $msg -BackgroundColor yellow -ForegroundColor Blue
        
        try
        {
            $Events = Get-WinEvent -Credential $GetAdminact -FilterHashtable $FilterHashProperties -Computer $Computer -ErrorAction SilentlyContinue 
            $Events | Select-Object -first $NumberOfLastEventsToGet
            $Events | Foreach-Object {
                $cur = $_ 
                try {
                    $xml = $_.ToXml()
                    $cur | Add-Member -MemberType NoteProperty -Name XML -Value $xml -Force
                } Catch {
                    Write-Host "error create xml" -ForegroundColor Red
                }
                
                $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
                if ($syslog -ne "") {
                    SendSyslog($cur)
                } 
            }
            Write-host "Found at least $($Events.count) events ! Here are the $NumberOfLastEventsToGet last ones"
                
        }
        Catch {
            $msg = "Error accessing Event Logs of $computer by Get-WinEvent + $PSItem.Exception.InnerExceptionMessage"
            Write-Host $msg -ForegroundColor Red
            try {    
                $Events = get-eventlog -logname $EventLogName -newest 10000 -Computer $Computer
                $Events | Where-Object {$eid -contains $_.EventID}
                $Events | Select-Object -first $NumberOfLastEventsToGet
                $Events | Foreach-Object {
                    $cur = $_ 
                    try {
                        $xml = $_.ToXml()
                        $cur | Add-Member -MemberType NoteProperty -Name XML -Value $xml -Force
                    } Catch {
                        Write-Host "error create xml" -ForegroundColor Red
                    }
                    
                    $cur | Add-Member -MemberType NoteProperty -Name XML -Value $xml -Force
                    $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append 
                    if ($syslog -ne "") {
                        SendSyslog($cur)
                    } 
                }
                
            }
            Catch {
                $msg = "Error accessing Event Logs of $computer by get-eventlog + $PSItem.Exception.InnerExceptionMessage"
                Write-Host $msg -ForegroundColor Red
            }

        }
        Finally {
            Write-Host "OK_"
        }
    }    
}


Foreach ($i in $target)
{
    if ($i -eq "Logon" -or $i -eq "All") {
        ExportFor ("4776","4672", "4624", "4634", "4800", "4801") "Security" "logon"
    }

    if ($i -eq "Service" -or $i -eq "All") {
        ExportFor ("7036","7031") "System" "service"
    }

    if ($i -eq "User" -or $i -eq "All") {
        ExportFor ("4720", "4722", "4723", "4724", "4725", "4726", "4738", "4740", "4767", "4780", "4794", "5376", "5377") "Security" "user"
    }

    if ($i -eq "Computer" -or $i -eq "All") {
        ExportFor ("4720", "4722", "4725", "4726", "4738", "4740", "4767") "Security" "user"
    }

    if ($i -eq "Clean" -or $i -eq "All") {
        
        Write-Host "EventID: " $id_clean
        ExportFor ("1102") "Security" "clean"
    }

    if ($i -eq "File" -or $i -eq "All") {
        ExportFor("4656", "4663", "4660", "4658") "Security" "file"
    }
    if ($i -eq "Printer" -or $i -eq "All") {
        ExportFor ("307")  ("Microsoft-Windows-PrintService/Operational") "printer"
    }

    if ($i -eq "MSSQL" -or $i -eq "All") {
        ExportFor ("18456")  "Application" "mssql"
    }

    if ($i -eq "RAS" -or $i -eq "All") {
        ExportFor ("20249", "20250", "20253", "20255", "20258", "20266", "20271", "20272") "RemoteAccess/Operational" "ras"
    }

    if ($i -eq "USB" -or $i -eq "All") {
        ExportFor ("2003") "Microsoft-Windows-DriverFrameworks-UserMode/Operational" "usb"
    }
    if ($i -eq "Sysmon" -or $i -eq "All") {
        ExportFor ("1", "3", "5", "11", "12", "13", "14") "Microsoft-Windows-Sysmon/Operational" "sysmon"
    }    
    if ($i -eq "TS" -or $i -eq "All") {
        ExportFor ("21", "24") "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" "ts"
    }
}

Write-host "I'm done. I exported the results to a file located on the same directory as the Script"

Write-Host "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
