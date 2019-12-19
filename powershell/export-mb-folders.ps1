param (
    [string]$outfilename = ".\explore-mailboxes",
    [string]$makves_url = "", ##"http://10.0.0.10:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"    
 )
 
 ## Init web server 
 $uri = $makves_url + "/agent/push"
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
 
 Write-Host "outfile: " $outfile

 function inspect($item) {
  $ t = $item 
  $t | Add-Member -MemberType NoteProperty -Name Type -Value "exchange-mailbox" -Force        
  $t | Add-Member -MemberType NoteProperty -Name Forwarder -Value "exchange-mailboxes-forwarder" -Force
  


  $JSON = $t | ConvertTo-Json
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
        inspect($cur)
   }
   Catch {
            $msg = "Error + $PSItem.Exception.InnerExceptionMessage"
            Write-Host $msg -ForegroundColor Red
  }
}
