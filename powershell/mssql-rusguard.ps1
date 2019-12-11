param (
    [string]$connection = 'server=localhost;user id=sa;password=Zse45rdx;',
    [string]$outfilename = '', ##'rusguard',
    [string]$start = "",
    [string]$startfn = "", ##".rusguard-monitor.time_mark",
    [string]$makves_url = "http://127.0.0.1:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
 )

[string] $query= "SELECT DateTime, LogMessageSubType, [DrvName], LastName, 
FirstName, SecondName, TableNumber, DepartmentName, Position
FROM  [RusGuardDB].[dbo].[EmployeesNLMK]";


Write-Host "connection: " $connection

#Create a variable for the date stamp in the log file

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

$markTime = Get-Date -format "yyyyMMddHHmmss"

if ($startfn -ne "") {
    Try
    {
        $start = Get-Content $startfn
    }
    Catch {
        Write-Host "Error read time mark:" + $PSItem.Exception.Message
        $start = ""
    }
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

    $data | Add-Member -MemberType NoteProperty -Name Forwarder -Value "event-direct" -Force
    $JSON = $data | ConvertTo-Json
    Try
    {
        if ($outfile -ne "") {
            $JSON | Out-File -FilePath $outfile -Encoding UTF8 -Append
        }
       
        if ($uri -ne "") {
            Try
            {
                $body = [System.Text.Encoding]::UTF8.GetBytes($JSON.ToString());

                Invoke-WebRequest -Uri $uri -Method Post -Body $body -ContentType "application/json" -Headers $headers
                Write-Host  "Send data to server:" + $cur.Name
            }
            Catch {
                Write-Host "Error send data to server:" + $PSItem.Exception.Message
            }
        }
    }
    Catch {
        Write-Host $PSItem.Exception.Message
    }
}



$resultsDataTable = New-Object System.Data.DataTable
$q = $query
if ($start -ne "") {
    $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
    $q += " where [DateTime] > '" + $starttime.ToString("yyyy-MM-dd HH:mm:ss") + "'"
}

$q += " order by [DateTime] ASC";

Write-Host $q

$resultsDataTable = ExecuteSqlQuery $connection $q


if ($resultsDataTable.Rows.Count -ne 0) {

    Write-Host ("The table contains: " + $resultsDataTable.Rows.Count + " rows")

    $resultsDataTable | Select-Object @{L='DateTime'; E ={$_.ItemArray[0]}},
    @{L='LogMessageSubType'; E ={$_.ItemArray[1]}},
    @{L='DrvName'; E ={$_.ItemArray[2]}},
    @{L='LastName'; E ={$_.ItemArray[3]}},
    @{L='FirstName'; E ={$_.ItemArray[4]}},
    @{L='SecondName'; E ={$_.ItemArray[5]}},
    @{L='TableNumber'; E ={$_.ItemArray[6]}},
    @{L='DepartmentName'; E ={$_.ItemArray[7]}},
    @{L='Position'; E ={$_.ItemArray[8]}} | Foreach-Object {
        $action = "entry"
        if ($_.LogMessageSubType -eq 67) {
            $action = "exit"
        }

        $content = "Employee: " + $_.LastName + " " + $_.FirstName + " " + $_.SecondName + "`n"
        $content += "Department: " + $_.DepartmentName + "`n"
        $content += "Position: " + $_.Position + "`n"
        $content += "TableNumber: " + $_.TableNumber + "`n"

        $data = @{
            time = $_.DateTime.ToString("dd.MM.yyyy HH:mm:ss")
            type = "direct-event"
            category = "RusGuard"
            object_type = "employee"
            who = $_.LastName
            where = $_.DrvName
            action = $action
            contents = $content
        }

        store($data)
    }
}

if ($startfn -ne "") {
    $markTime | Out-File -FilePath $startfn -Encoding UTF8
    Write-Host "Store new mark: " $markTime
}