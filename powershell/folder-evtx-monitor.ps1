param (
    [string]$folder = "C:\Windows\System32\winevt\Logs",
    [string]$url = "http://10.0.0.10:8000",
    [string]$user = "admin",
    [string]$pass = "admin"
 )

$uri = $url + "/data/upload"
$pair = "${user}:${pass}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue }
function inspectFile($fn) {
    $response = Invoke-WebRequest -Uri $uri -Method Post -InFile $uploadPath -Headers $headers
}

$filter = '*.evtx'  # You can enter a wildcard filter here. 


$fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $false;NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'} 
 
# Here, all three events are registerd.  You need only subscribe to events that you need: 
 
Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action { 
$name = $Event.SourceEventArgs.Name 
$changeType = $Event.SourceEventArgs.ChangeType 
$timeStamp = $Event.TimeGenerated 
Write-Host "The file '$name' was $changeType at $timeStamp" -fore green 
inspectFile $name
} 
 
Register-ObjectEvent $fsw Deleted -SourceIdentifier FileDeleted -Action { 
$name = $Event.SourceEventArgs.Name 
$changeType = $Event.SourceEventArgs.ChangeType 
$timeStamp = $Event.TimeGenerated 
Write-Host "The file '$name' was $changeType at $timeStamp" -fore red 
} 
 
Register-ObjectEvent $fsw Changed -SourceIdentifier FileChanged -Action { 
$name = $Event.SourceEventArgs.Name 
$changeType = $Event.SourceEventArgs.ChangeType 
$timeStamp = $Event.TimeGenerated 
Write-Host "The file '$name' was $changeType at $timeStamp" -fore white 
inspectFile $name } 


while ($true)
{
    Start-Sleep -Milliseconds 1000
	if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character))
    {
        Write-Host "You pressed CTRL-C. Do you want to continue doing this and that?" 
        $key = $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
        if ($key.Character -eq "N") { break; }
    }
}

Unregister-Event FileDeleted 
Unregister-Event FileCreated 
Unregister-Event FileChanged

