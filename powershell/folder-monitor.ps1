param (
    [string]$folder = "C:\work\",
    [string]$url = "http://localhost:8000",
    [string]$user = "admin",
    [string]$pwd = "admin"
 )

$uri = $url + "/data/upload/file-info"
$pair = "${user}:${pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

Add-Type -AssemblyName 'System.Net.Http'


Function Get-MKVS-FileHash([String] $FileName,$HashName = "SHA1") 
{
    if ($hashlen -eq 0) {
        $FileStream = New-Object System.IO.FileStream($FileName,"Open", "Read") 
        $StringBuilder = New-Object System.Text.StringBuilder 
        [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($FileStream)|%{[Void]$StringBuilder.Append($_.ToString("x2"))} 
        $FileStream.Close() 
        $StringBuilder.ToString()
    } else {
        $StringBuilder = New-Object System.Text.StringBuilder 
        $binaryReader = New-Object System.IO.BinaryReader(New-Object System.IO.FileStream($FileName,"Open", "Read"))
       
        $bytes = $binaryReader.ReadBytes($hashlen)
        $binaryReader.Close() 
        if ($bytes -ne 0) {
            [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($bytes)| ForEach-Object { [Void]$StringBuilder.Append($_.ToString("x2")) }
        }
        $StringBuilder.ToString()
    }
}

function Get-MKVS-DocText([String] $FileName) {
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0
    $text = ""
    Try
    {
        $catch = $false
        Try{
            $Document = $Word.Documents.Open($FileName, $null, $null, $null, "")
        }
        Catch {
            Write-Host 'Doc is password protected.'
            $catch = $true
        }
        if ($catch -eq $false) {
            $Document.Paragraphs | ForEach-Object {
                $text += $_.Range.Text
            }
            
        }
    }
    Catch {
        Write-Host $PSItem.Exception.Message
        $Document.Close()
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
        Remove-Variable Word
    }
    $Document.Close()
    $Word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
    Remove-Variable Word        
    return $text
}

function Get-MKVS-XlsText([String] $FileName) {
    $excel = New-Object -ComObject excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = 0
    $text = ""
    $password
    Try    
    {
        $catch = $false
        Try{
            $wb =$excel.Workbooks.open($path, 0, 0, 5, "")
        }
        Catch{
            Write-Host 'Book is password protected.'
            $catch = $true
        }
        if ($catch -eq $false) {
            foreach ($sh in $wb.Worksheets) {
                #Write-Host "sheet: " $sh.Name            
                $endRow = $sh.UsedRange.SpecialCells(11).Row
                $endCol = $sh.UsedRange.SpecialCells(11).Column
                Write-Host "dim: " $endRow $endCol
                for ($r = 1; $r -le $endRow; $r++) {
                    for ($c = 1; $c -le $endCol; $c++) {
                        $t = $sh.Cells($r, $c).Text
                        $text += $t
                        #Write-Host "text cel: " $r $c $t
                    }
                }
            }
        }
    }
    Catch {
        Write-Host $PSItem.Exception.Message
    }
    #Write-Host "text: " $text
    $excel.Workbooks.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Remove-Variable excel
    return $text
}

function Get-MKVS-FileText([String] $FileName, [String] $Extension) {
    Write-Host "filename: " $FileName
    Write-Host "ext: " $Extension

    switch ($Extension) {
        ".doc" {
            return Get-MKVS-DocText $FileName
        }
        ".docx" {
            return Get-MKVS-DocText $FileName
        }
        ".xls" {
            return Get-MKVS-XlsText $FileName
        }
        ".xlsx" {
            return Get-MKVS-XlsText $FileName
        }
    }
    return ""    
}

function inspectFile($fullpath) {
    Write-Host $fullpath
    Try
    {
        $cur =  Get-ChildItem $fullpath

        $cur = $cur | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"

        
        $acl = Get-Acl $cur.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
        $path = $cur.FullName
        $ext = $cur.Extension
        
        if ($cur.PSIsContainer -eq $false) {
            Try
            {
                $hash = Get-MKVS-FileHash $path
            }
            Catch {
                Write-Host $PSItem.Exception.Message
                Try
                {
                    $hash = Get-FileHash $path | Select-Object -Property "Hash"
                }
                Catch {
                    Write-Host $PSItem.Exception.Message
                }
            }

            if ($extruct -eq $true)
            {
                Try
                {
                    $text =  Get-MKVS-FileText $path $ext
                    $cur | Add-Member -MemberType NoteProperty -Name Text -Value $text -Force
                }
                Catch {
                    Write-Host "Get-MKVS-FileText error:" + $PSItem.Exception.Message
                }    
            }
            $cur | Add-Member -MemberType NoteProperty -Name Hash -Value $hash -Force
        }
        
        $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
        Try
        {
            store($cur)
        }
        Catch {
            Write-Host "ConvertTo-Json error:" + $PSItem.Exception.Message
        }
    }
    Catch {
        Write-Host $PSItem.Exception.Message
    }
}

function store($data) {
    $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "folder-forwarder" -Force
    $JSON = $data | ConvertTo-Json
    $response = Invoke-WebRequest -Uri $uri -Method Post -Body $JSON -ContentType "application/json" -Headers $headers
}



$filter = '*.*'  # You can enter a wildcard filter here. 


$fsw = New-Object IO.FileSystemWatcher $folder, $filter -Property @{IncludeSubdirectories = $true; NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'} 
 
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

