param (
    [string]$outfilename = ".\explore-mailboxes"    
 )
 



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"
Write-Host "outfile: " $outfile

if (Test-Path $outfile) 
{
    Remove-Item $outfile
}


Get-Mailbox -resultSize Unlimited | Select-Object Alias, Name, DisplayName, ServerName |
foreach-object {
    try
    {
        $cur = $_
        $fstat = Get-MailboxFolderStatistics -Identity $cur.alias
        $s = Get-MailboxStatistics -Identity $cur.alias | Select-Object DisplayName, LastLogonTime, ItemCount, LastLogoffTime, LegacyDN, LastLoggedOnUserAccount, ObjectClass
        $s | ForEach-Object {
            $cs = $_ | ConvertTo-Json
            Write-Host "Statistic: " $cs
        }
        $cur | Add-Member -MemberType NoteProperty -Name Statistic -Value $s -Force
        
        try{
            $uc= Get-MailboxUserConfiguration -Identity ($cur.alias)
            if ($uc -ne $null) {

              Write-Host "UserConfiguration: " $uc
              $cur | Add-Member -MemberType NoteProperty -Name UserConfiguration -Value $uc -Force
            }
        }
        Catch {
            $msg = "Error + $PSItem.Exception.InnerExceptionMessage"
            Write-Host $msg -ForegroundColor Red
        }

        $p = Get-MailboxPermission -Identity ($cur.alias) | Select-Object Identity, User, AccessRights
        if ($p -ne $null) {

          [string]$t = "["
          $p| foreach-object {
            if ($t.length -ne 1) {
                $t += ","
            }
            $t += "{`"Identity`": `"$($_.Identity)`","
            $t += "`"User`": `"$($_.User)`","
            $t += "`"AccessRights`": `"$($_.AccessRights)`"}"
          }
          $t += "]"
          Write-Host "permision: " $t
          $cur | Add-Member -MemberType NoteProperty -Name Permissions -Value $t -Force
        }


        $p = Get-MailboxFolderPermission -Identity ($cur.alias + ":\") -ErrorAction SilentlyContinue | Select-Object Identity, FolderName, User, AccessRights
        if ($p -ne $null) {

          [string]$t = "["
          $p| foreach-object {
            if ($t.length -ne 1) {
                $t += ","
            }
            $t += "{`"Identity`": `"$($_.Identity)`","
            $t += "`"FolderName`": `"$($_.FolderName)`","
            $t += "`"User`": `"$($_.User)`","
            $t += "`"AccessRights`": `"$($_.AccessRights)`"}"
          }
          $t += "]"
          Write-Host "folder permision: " $t
          $cur | Add-Member -MemberType NoteProperty -Name FolderPermissions -Value $t -Force
        }
    
        $cur | Add-Member -MemberType NoteProperty -Name Folders -Value $fstat -Force
        $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
   }
   Catch {
            $msg = "Error + $PSItem.Exception.InnerExceptionMessage"
            Write-Host $msg -ForegroundColor Red
  }
}
