[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] [string]$url = "http://10.0.0.10:8000",
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")] [string]$user = "admin",
    [Parameter(Mandatory = $False, Position = 3, ParameterSetName = "NormalRun")] [string]$pwd = "admin"
)

$currentComputer = $Env:Computername
$sig = @'
[DllImport("advapi32.dll", SetLastError = true)]
public static extern bool GetUserName(System.Text.StringBuilder sb, ref Int32 length);
'@

Add-Type -MemberDefinition $sig -Namespace Advapi32 -Name Util

$size = 64
$str = New-Object System.Text.StringBuilder -ArgumentList $size

[Advapi32.util]::GetUserName($str, [ref]$size) |Out-Null
$currentUser = $str.ToString()



Register-CimIndicationEvent -ClassName Win32_ProcessStartTrace -SourceIdentifier "ProcessStarted"
function store($data) {
    $JSON = $data | ConvertTo-Json
    $response = Invoke-WebRequest -Uri $uri -Method Post -Body $JSON -ContentType "application/json" -Headers $headers
}

while ($true)
{
    Start-Sleep -Milliseconds 1000
	if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character))
    {
        Write-Host "You pressed CTRL-C. Do you want to continue doing this and that?" 
        $key = $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
        if ($key.Character -eq "N") { break; }
	}
	$events = Get-Event | Select-Object @{L='start-time'; E ={timegenerated}}, @{L='process'; E ={$_.sourceeventargs.newevent.processname}}
	if ($events.length -ne 0) {
		$data = @{ 
			data = $events
			type = "process-start"
			user = $currentUser
			computer = $currentComputer
			time = Get-Date -Format "dd.MM.yyyy HH:mm:ss"}	
		store($data)
	}
	
}




Get-Event | Remove-Event
Get-EventSubscriber | Unregister-Event