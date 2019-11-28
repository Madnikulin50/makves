param (
    [string]$connection = 'server=gamma;user id=sa;password=P@ssw0rd;',
    [string]$outfilename = 'rusguard',
    [string]$start = "",
    [string]$makves_url = "",##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
 )

[string] $query= "SELECT DateTime, LogMessageSubType, DrvName, LastName, 
FirstName, SecondName, TableName, DepartmentName, Position
FROM  [RusGuardDB].[dbo].[EmployeesNLMK]";


Write-Host "connection: " $connection

#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm

Import-Module ActiveDirectory

$SearchBase = $base 

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



function ExecuteSqlQuery ($connectionString, $query) {
    $Datatable = New-Object System.Data.DataTable
    
    $Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = $connectionString
    $Connection.Open()
    $Command = New-Object System.Data.SQLClient.SQLCommand
    $Command.Connection = $Connection
    $Command.CommandText = $query
    $Reader = $Command.ExecuteReader()
    $Datatable.Load($Reader)
    $Connection.Close()
    
    return $Datatable
}



function store($data) {
    $data | Add-Member -MemberType NoteProperty -Name Forwarder -Value "event-forwarder" -Force
    $JSON = $data | ConvertTo-Json
    Try
    {
        Invoke-WebRequest -Uri $uri -Method Post -Body $JSON -ContentType "application/json" -Headers $headers
        Write-Host  "Send data to server:" + $data.Name
    }
    Catch {
        Write-Host $PSItem.Exception.Message
    }
}

$global:ErrorlastTime = ""

while ($true)
{
    Start-Sleep -Milliseconds 1000
	if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character))
    {
        Write-Host "You pressed CTRL-C. Do you want to continue doing this and that?" 
        $key = $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
        if ($key.Character -eq "N") { break; }
	}
	$resultsDataTable = New-Object System.Data.DataTable
	$q = $query
	if ($global:lastTime -ne "") {
		$q += " where [DataTime] > '" + $global:lastTime + "'"
	}

	$q += " order by [DataTime] DESC";

	Write-Host $q

	$resultsDataTable = ExecuteSqlQuery $connection $q

	if ($resultsDataTable.Rows.Count -ne 0) {
		Write-Host ("The table contains: " + $resultsDataTable.Rows.Count + " rows")
		$res = $resultsDataTable | Select-Object @{L='time'; E ={$_.ItemArray[0]}}, @{L='login'; E ={$_.ItemArray[1]}}, @{L='query'; E ={$_.ItemArray[2]}}, @{L='program'; E ={$_.ItemArray[3]}}, @{L='host'; E ={$_.ItemArray[4]}}
		$global:lastTime = $res[0].time
		$data = @{ 
			data = $res
			type = "event"
			user = $currentUser
			computer = $currentComputer
			time = Get-Date -Format "dd.MM.yyyy HH:mm:ss"}	
		store($data)
	}
	
}
