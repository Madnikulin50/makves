param (
    [string]$folder = "z:\test\Files_DB",
    [string]$outfilename = "explore-folder",
    [string]$base = "",
    [string]$server = "",
    [int]$hashlen = 1048576,
    [switch]$no_hash = $false,
    [switch]$extruct = $false,
    [string]$start = ""
 )
 



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"

Write-Host "base: " $folder
Write-Host "outfile: " $outfile

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}

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

function inspectFile($cur) {
    $cur = $cur | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"
    Write-Host $cur.FullName
        $acl = Get-Acl $cur.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
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
            $cur | Add-Member -MemberType NoteProperty -Name Hash -Value $hash -Force
        }
        
        $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
        Try
        {
            $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
        }
        Catch {
            Write-Host "ConvertTo-Json error:" + $PSItem.Exception.Message
        }
}

function inspectFolder($f) {
    $cur  = Get-Item $f | 
    Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"
    
    Write-Host $cur.FullName
    $acl = Get-Acl $cur.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
    $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
    $cur | Add-Member -MemberType NoteProperty -Name RootAudit -Value $true -Force
    $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append

    
    if ($start -ne "") {
        Write-Host "start: " $start
        $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)

        Get-ChildItem $f -Recurse | ? { $_.LastWriteTime -gt $starttime } | Foreach-Object {
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
    $computers | ForEach {
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
