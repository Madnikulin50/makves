param (
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] [string]$folder = "C:\Windows\System32\winevt\Logs",
    [string]$url = "http://10.0.0.10:8000",
    [string]$user = "admin",
    [string]$pwd = "admin",
    [Parameter(Mandatory = $False, Position = 10, ParameterSetName = "NormalRun")] [ValidateSet("Security","Application","System","All")] [array]$target="Security"

 )

$uri = $url + "/data/upload"
$pair = "${user}:${pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue 
"X-Forwarder" = "evtx-forwarder"}
Add-Type -AssemblyName 'System.Net.Http'

function isMyFileName($entry) {
	$filepath = Get-ChildItem $entry
	$fn = $filepath.BaseName
	Foreach ($i in $target) {
	    if ($i -eq "All") {
			return $True
		}
		
		if ($i -eq $fn) {
			return $True
		}
		
	}
	return $False
}

function inspectFile($entry) {
    Try
    {
		$isMy = isMyFileName($entry)
		if ($isMy -eq $false) {
			return
		}

        $tempFile = $entry + "copy"
        Copy-Item $entry $tempFile
        Try
        {
            #$response = Invoke-WebRequest -Uri $uri -Method Post -InFile $tempFile -Headers $headers -ContentType "multipart/form-data"
            $fileName = "events.evtx"
            $fileBin = [IO.File]::ReadAllBytes($tempFile)
            $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
            $fileEnc = $enc.GetString($fileBin)

            $boundary = [System.Guid]::NewGuid().ToString()
            $LF = "`r`n"
            $bodyLines = (
                "--$boundary",
                "Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"",
                "Content-Type: application/octet-stream$LF",
                $fileEnc,
                "--$boundary--$LF"
            ) -join $LF

            Invoke-RestMethod -Uri $uri -Method Post -ContentType "multipart/form-data; boundary=`"$boundary`"" -Body $bodyLines -Headers $headers
        }
        Catch {
            Write-Host "send file error:" + $PSItem.Exception.Message
        }
        Remove-Item -path $tempFile
    }
    Catch {
        Write-Host "inspectFile error:" + $PSItem.Exception.Message
    }
}

$filter = '*.evtx'  # You can enter a wildcard filter here. 


$fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $false;NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'} 
 
# Here, all three events are registerd.  You need only subscribe to events that you need: 
 
Register-ObjectEvent $fsw Created -SourceIdentifier FileCreated -Action { 
    $name = $Event.SourceEventArgs.Name
    $changeType = $Event.SourceEventArgs.ChangeType
    $timeStamp = $Event.TimeGenerated
    Write-Host "The file '$name' was $changeType at $timeStamp" -fore green
    $fullname = Join-Path -Path $folder -ChildPath $name
    inspectFile $fullname
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
    $fullname = Join-Path -Path $folder -ChildPath $name
    inspectFile $fullname
}


Get-ChildItem $folder -Recurse | Foreach-Object {
    Try
    {
        if ($_.Extension -eq ".evtx") {
            inspectFile $_.FullName
        }
        
    }
    Catch {
        Write-Host "inspectFile error:" + $PSItem.Exception.Message
    }
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
}

Unregister-Event FileDeleted 
Unregister-Event FileCreated 
Unregister-Event FileChanged

