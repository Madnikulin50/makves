param (
    [string]$base = 'DC=testdomain,DC=local',
    [string]$server = 'kappa.testdomain.local',
    [string]$outfilename = 'shares',
    [switch]$force = $false
 )



Write-Host "base: " $base
Write-Host "server: " $server
Write-Host "outfile: " $outfile
#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm


$outfile = "$($outfilename)_$LogDate.json"

Import-Module ActiveDirectory

$GetAdminact = Get-Credential

$outfile = "$($outfilename).json"

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}


function Export-Shares($machine) {
    Write-Host "export shares from machine: " $machine
    net view $machine | Select-Object -Skip  7 | Select-Object -SkipLast 2|
    ForEach-Object -Process {[regex]::replace($_.trim(),'\s+',' ')} |
    ConvertFrom-Csv -delimiter ' ' -Header 'sharename', 'type', 'usedas', 'comment' |
    foreach-object {
        Try {
            Get-Acl "\\$($machine)\$($_.sharename)"
        } Catch {
            # Do nothing
        }
    } | Select-Object -Property "Path", "Owner", "Group", "Access" |
    ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}



$computers = Get-ADComputer -Filter * -server $server `
-Credential $GetAdminact -searchbase $base ` |
Select-Object "Name"
$computers | ForEach {Export-Shares $_.Name}