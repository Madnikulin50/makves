param (
    [string]$folder = "C:\work",
    [string]$outfilename =  "",
    [string]$base = "",
    [string]$server = "",
    [int]$hashlen = 1048576,
    [switch]$no_hash = $false,
    [switch]$extruct = $false,
    [switch]$compliance = $true,
    [string]$start = "",
    [string]$startfn = "", ##".file-monitor.time_mark",
    [string]$makves_url = "http://192.168.2.22:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
)

## Init web server 
$uri = $makves_url + "/data/upload/file-info"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

if ($makves_url -eq "") {
    $uri = ""
    Add-Type -AssemblyName 'System.Net.Http'
}


if ($compliance -eq $true) {
    Import-Module "./compliance.dll" -Verbose
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



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = ""

if ($outfilename -ne "") {
    $outfile = "$($outfilename)_$LogDate.json"
    if (Test-Path $outfile) 
    {
        Remove-Item $outfile
    }
}

Write-Host "base: " $folder
Write-Host "outfile: " $outfile



Function Get-MKVS-FileHash([String] $FileName,$HashName = "SHA1") 
{
    if ($hashlen -eq 0) {
        $FileStream = New-Object System.IO.FileStream($FileName,"Open", "Read") 
        $StringBuilder = New-Object System.Text.StringBuilder 
        [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($FileStream)| ForEach-Object {[Void]$StringBuilder.Append($_.ToString("x2"))} 
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

function inspectFile($cur) {
    $cur = $cur | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"
	Write-Host $cur.FullName
	Try
	{
		$acl = Get-Acl $cur.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
	}
	Catch {
		Write-Host "Get-Acl error:" + $PSItem.Exception.Message
	}
    $path = $cur.FullName
    $ext = $cur.Extension
        
    if ($cur.PSIsContainer -eq $false) {
        if ($no_hash -eq $false) {
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

        if ($compliance -eq $true)
        {
            Try
            {
                $compliance =  Get-Compliance -File $path
                $cur | Add-Member -MemberType NoteProperty -Name Compliance -Value $compliance -Force
            }
            Catch {
				Write-Host "Get-Compliance error:" + $PSItem.Exception.Message
            }
        }

        $cur | Add-Member -MemberType NoteProperty -Name Hash -Value $hash -Force
    }
        
    $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
    $cur | Add-Member -MemberType NoteProperty -Name Computer -Value $env:computername -Force
    Try
    {
		if ($outfile -ne "") {
            $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
        }
           
        if ($uri -ne "") {
            $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "folder-forwarder" -Force
            $JSON = $cur | ConvertTo-Json
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
        Write-Host "Store error:" + $PSItem.Exception.Message
    }
}

function inspectFolder($f) {
    Try
    {
        $cur  = Get-Item $f  
    } Catch {
        Write-Host "Error Get-Item:" + $f + ":" $PSItem.Exception.Message
        return;
    }

    if ($cur -eq $null) {
        return;
    }

    $cur  = $cur | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"
    
    Write-Host $cur.FullName
    Try
    {
        $acl = Get-Acl $cur.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"  
    } Catch {
        Write-Host "Error Get-Acl:" + $f + ":" $PSItem.Exception.Message
        return;
    }
    $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
    $cur | Add-Member -MemberType NoteProperty -Name RootAudit -Value $true -Force
    $cur | Add-Member -MemberType NoteProperty -Name Computer -Value $env:computername -Force
  
    if ($outfile -ne "") {
        $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
    }
    
    if ($uri -ne "") {
        $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "folder-forwarder" -Force
        $JSON = $cur | ConvertTo-Json
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

    
    if ($start -ne "") {
        Write-Host "start: " $start
        $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)

        Get-ChildItem $f -Recurse | Where-Object { $_.LastWriteTime -gt $starttime } | Foreach-Object {
            Try
            {
                inspectFile $_
            }
            Catch {
                Write-Host "inspectFile error:" + $PSItem.Exception.Message
            }
        }
    } else {
        Get-ChildItem $f -Recurse | Foreach-Object {
            Try
            {
                inspectFile $_
            }
            Catch {
                Write-Host "inspectFile error:" + $PSItem.Exception.Message
            }
        }
    }
}


if ($base -eq "" ) {
    inspectFolder $folder
} else {
    Import-Module ActiveDirectory
    $GetAdminact = Get-Credential
    $computers = Get-ADComputer -Filter * -server $server -Credential $GetAdminact -searchbase $base | Select-Object "Name"    
    $computers | ForEach-Object {
        $machine = $_.Name
        Write-Host "export shares from machine: " $machine
        net view $machine | Select-Object -Skip  7 | Select-Object -SkipLast 2|
        ForEach-Object -Process {[regex]::replace($_.trim(),'\s+',' ')} |
        ConvertFrom-Csv -delimiter ' ' -Header 'sharename', 'type', 'usedas', 'comment' |
        foreach-object {
            inspectFolder "\\$($machine)\$($_.sharename)"
        }
    }

}

if ($startfn -ne "") {
    $markTime | Out-File -FilePath $startfn -Encoding UTF8
    Write-Host "Store new mark: " $markTime
}
