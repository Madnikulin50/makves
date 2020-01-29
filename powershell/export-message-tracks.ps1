param (
    [string]$outfilename = 'export_message_tracks',
    [string]$start = "",
    
    [string]$makves_url = "",##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"

 )


$LogDate = get-date -f yyyyMMddhhmm


## Init web server 
$uri = $makves_url + "/data/upload/agent"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

if ($makves_url -eq "") {
    $uri = ""
    Add-Type -AssemblyName 'System.Net.Http'
}

$outfile = ""

if ($outfilename -ne "") {
    $outfile = "$($outfilename)_$LogDate.json"
    if (Test-Path $outfile) 
    {
        Remove-Item $outfile
    }
}

Write-Host "outfile: " $outfile
$starttime = ""

if ($start -ne "") {
  Write-Host "start: " $start
  $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
}


function store($item) {
    if ($outfile -ne "") {
        $item | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
    }
    $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "message-tracks-forwarder" -Force
    $JSON = $data | ConvertTo-Json
    $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
    $response = Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
    return $response  
}


if ($starttime -ne "") {

    Get-MessageTrackingLog -EventId Send -Start $starttime | ForEach-Object {
        store($_)
    }

    Get-MessageTrackingLog -EventId Receive -Start $starttime | ForEach-Object {
        store($_)
    }
} else {
    Get-MessageTrackingLog -EventId Send -Start | ForEach-Object {
        store($_)
    }

    Get-MessageTrackingLog -EventId Receive | ForEach-Object {
        store($_)
    }
}




Write-Host "MessageTracksExport finished export finished to: " $outfile