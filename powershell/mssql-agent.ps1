[string] $connection= "server=gamma;user id=sa;password=P@ssw0rd;";

[string] $query= "SELECT      last_request_start_time,
login_name, text, program_name, host_name
FROM        sys.dm_exec_connections c
INNER JOIN  sys.dm_exec_sessions s ON c.session_id = s.session_id
CROSS APPLY sys.dm_exec_sql_text(most_recent_sql_handle) AS st";


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
	$JSON = $data | ConvertTo-Json
	$body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());
    $response = Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
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
		$q += " where last_request_start_time > '" + $global:lastTime + "'"
	}

	$q += " order by last_request_start_time DESC";

	Write-Host $q

	$resultsDataTable = ExecuteSqlQuery $connection $q

	if ($resultsDataTable.Rows.Count -ne 0) {
		Write-Host ("The table contains: " + $resultsDataTable.Rows.Count + " rows")
		$res = $resultsDataTable | Select-Object @{L='time'; E ={$_.ItemArray[0]}}, @{L='login'; E ={$_.ItemArray[1]}}, @{L='query'; E ={$_.ItemArray[2]}}, @{L='program'; E ={$_.ItemArray[3]}}, @{L='host'; E ={$_.ItemArray[4]}}
		$global:lastTime = $res[0].time
		$data = @{ 
			data = $res
			type = "mssql"
			user = $currentUser
			computer = $currentComputer
			time = Get-Date -Format "dd.MM.yyyy HH:mm:ss"}	
		store($data)
	}
	
}
