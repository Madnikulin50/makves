param (
    [string]$base = 'DC=acme,DC=local',
    [string]$server = 'acme.local',
    [string]$timeout = 3600,
    [string]$outfile = "",
    [string]$user = "",
    [string]$pwd = "",
    [string]$start = "",
    [string]$makves_url = "http://localhost:8000",
    [string]$makves_user = "admin",
    [string]$makves_pwd = "admin"
 )

Write-Host "base: " $base
Write-Host "server: " $server
Write-Host "outfile: " $outfile
Write-Host "user: " $user
Write-Host "pwd: " $pwd

Write-Host "makves_url: " $makves_url
Write-Host "makves_user: " $makves_user
Write-Host "makves_pwd: " $makves_pwd


## Init web server 
$uri = $urmakves_url + "/data/upload/ad"
$pair = "${makves_user}:${makves_pwd}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue}

Add-Type -AssemblyName 'System.Net.Http'
## End web server init


Import-Module ActiveDirectory

$SearchBase = $base 

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}


if ($user -ne "") {
    $pass = ConvertTo-SecureString -AsPlainText $pwd -Force    
    $GetAdminact = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pass    
} else {
    $GetAdminact = Get-Credential
}

$domain = Get-ADDomain -server $server -Credential $GetAdminact


if ($start -ne "") {
  Write-Host "start: " $start
  $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
}

Write-Host "domain: " $domain.NetBIOSName

func store($item) {
    if ($outfile -ne "") {
        $item | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
    }
    $cur | Add-Member -MemberType NoteProperty -Name Forwarder -Value "ad-forwarder" -Force
    $JSON = $data | ConvertTo-Json
    $response = Invoke-WebRequest -Uri $uri -Method Post -Body $JSON -ContentType "application/json" -Headers $headers
    return $response
}
  


function Get-ADPrincipalGroupMembershipRecursive() {
  Param(
      [string] $dsn,
      [array]$groups = @()
  )

  $obj = Get-ADObject -server $server  -Credential $GetAdminact $dsn -Properties memberOf

  foreach( $groupDsn in $obj.memberOf ) {

      $tmpGrp = Get-ADObject -server $server  -Credential $GetAdminact $groupDsn -Properties * | Select-Object "Name", "cn", "distinguishedName", "objectSid", "DisplayName", "memberOf"

      if( ($groups | Where-Object { $_.DistinguishedName -eq $groupDsn }).Count -eq 0 ) {
          $groups +=  $tmpGrp           
          $groups = Get-ADPrincipalGroupMembershipRecursive $groupDsn $groups
      }
  }

  return $groups
}

function inspectComputers() {
    Get-ADComputer -Filter * -Properties * -server $server  -Credential $GetAdminact -searchbase $SearchBase |
    Select-Object "Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
   "sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
   "PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "lastlogondate", "ObjectClass", "DNSHostName", "ObjectCategory", "LastBadPasswordAttempt" |
   Foreach-Object {
     $cur = $_
     if ($start -ne "") {
       if (($cur.whenChanged -lt $starttime) -and ($cur.lastlogondate -lt $starttime)) {
         Write-Host "skip " $cur.Name
         return
       }
       Write-Host "write " $cur.Name
   
     }
   
     $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
     $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force
     $licensies = Get-WmiObject SoftwareLicensingProduct -Credential $GetAdminact -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object Description, LicenseStatus
     if ($Null -ne $licensies) {
       Write-Host $cur.DNSHostName " : " $($licensies)
       $cur | Add-Member -MemberType NoteProperty -Name OperatingSystemLicensies -Value $licensies -Force
     }
   
     Try {
       $userprofiles = Get-WmiObject -Credential $GetAdminact -Class win32_userprofile -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object sid, localpath 
       if ($null -ne $userprofiles) {
         Write-Host $cur.DNSHostName  " : " $userprofiles
         $cur | Add-Member -MemberType NoteProperty -Name UserProfiles -Value $userprofiles -Force
       }    
     } Catch {
       Write-Host $cur.DNSHostName  " : " "$($_.Exception.Message)"
     }
   
     Try {
       $apps = Get-WMIObject -Class win32_product -Credential $GetAdminact -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object Name, Version
       if ($Null -ne $apps) {
         Write-Host $cur.DNSHostName " : " $apps
         $cur | Add-Member -MemberType NoteProperty -Name Applications -Value $apps -Force
       }
       
     }
     Catch {
         Write-Host $cur.DNSHostName " win32_product Offline "
         try {
          
   
           $Registry = $Null;
           Try{$Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $cur.DNSHostName);}
           Catch{Write-Host -ForegroundColor Red "$($_.Exception.Message)";}
           
           If ($Registry){
             $apps =  New-Object System.Collections.Generic.List[System.Object];
             $UninstallKeys = $Null;
             $SubKey = $Null;
             $UninstallKeys = $Registry.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall",$False);
             $UninstallKeys.GetSubKeyNames()| ForEach-Object {
               $SubKey = $UninstallKeys.OpenSubKey($_,$False);
               $DisplayName = $SubKey.GetValue("DisplayName");
               If ($DisplayName.Length -gt 0){
                 $Entry = $Base | Select-Object *
                 $Entry.ComputerName = $ComputerName;
                 $Entry.Name = $DisplayName.Trim(); 
                 $Entry.Publisher = $SubKey.GetValue("Publisher"); 
                 [ref]$ParsedInstallDate = Get-Date
                 If ([DateTime]::TryParseExact($SubKey.GetValue("InstallDate"),"yyyyMMdd",$Null,[System.Globalization.DateTimeStyles]::None,$ParsedInstallDate)){					
                 $Entry.InstallDate = $ParsedInstallDate.Value
                 }
                 $Entry.EstimatedSize = [Math]::Round($SubKey.GetValue("EstimatedSize")/1KB,1);
                 $Entry.Version = $SubKey.GetValue("DisplayVersion");
                 [Void]$apps.Add($Entry);
               }
             }
             
               If ([IntPtr]::Size -eq 8){
                       $UninstallKeysWow6432Node = $Null;
                       $SubKeyWow6432Node = $Null;
                       $UninstallKeysWow6432Node = $Registry.OpenSubKey("Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",$False);
                           If ($UninstallKeysWow6432Node) {
                               $UninstallKeysWow6432Node.GetSubKeyNames()| ForEach-Object {
                               $SubKeyWow6432Node = $UninstallKeysWow6432Node.OpenSubKey($_,$False);
                               $DisplayName = $SubKeyWow6432Node.GetValue("DisplayName");
                               If ($DisplayName.Length -gt 0){
                                 $Entry = $Base | Select-Object *
                                   $Entry.ComputerName = $ComputerName;
                                   $Entry.Name = $DisplayName.Trim(); 
                                   $Entry.Publisher = $SubKeyWow6432Node.GetValue("Publisher"); 
                                   [ref]$ParsedInstallDate = Get-Date
                                   If ([DateTime]::TryParseExact($SubKeyWow6432Node.GetValue("InstallDate"),"yyyyMMdd",$Null,[System.Globalization.DateTimeStyles]::None,$ParsedInstallDate)){                     
                                   $Entry.InstallDate = $ParsedInstallDate.Value
                                   }
                                   $Entry.EstimatedSize = [Math]::Round($SubKeyWow6432Node.GetValue("EstimatedSize")/1KB,1);
                                   $Entry.Version = $SubKeyWow6432Node.GetValue("DisplayVersion");
                                   $Entry.Wow6432Node = $True;
                                   [Void]$apps.Add($Entry);
                                 }
                               }
                         }
                       }
              Write-Host $cur.DNSHostName + " : " $apps
              $cur | Add-Member -MemberType NoteProperty -Name Applications -Value $apps -Force
           }
         } Catch {
           Write-Host $cur.DNSHostName " error apps" "$($_.Exception.Message)"
         }
      }
   
   
    $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
    $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force
    store $cur
   }
   Write-Host "computers inspect finished"   
}

function inspectComputers() {
    Get-ADGroup -server $server `
    -Credential $GetAdminact -searchbase $SearchBase `
    -Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} | Select-Object "Name", "GivenName", "Surname", "sn", "cn", "distinguishedName",
    "whenCreated", "whenChanged", "memberOf", "objectSid", "DisplayName", 
    "sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
    "Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
    "Mail", "userAccountControl", "Manager", "ObjectClass", "logonCount", "UserPrincipalName"| Foreach-Object {
      $cur = $_ 
      if ($start -ne "") {
        if ($cur.whenChanged -lt $starttime) {
          Write-Host "skip " $cur.Name
          return
        }
    
      }
    
      $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
      $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force
      
      $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
      $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force
    
      store $cur
    }
    
    Write-Host "groups inspect finished"
    
}


function inspectUsers() {
    Get-ADUser -server $server `
    -Credential $GetAdminact -searchbase $SearchBase `
    -Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} | Select-Object "Name", "GivenName", "Surname", "sn", "cn", "distinguishedName",
    "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
    "sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
    "Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
    "Mail", "userAccountControl", "PasswordNeverExpires", "PasswordExpired", "DoesNotRequirePreAuth",
    "CannotChangePassword", "PasswordNotRequired", "TrustedForDelegation", "TrustedToAuthForDelegation",
    "Manager", "Enabled", "lastlogondate", "ObjectClass", "logonCount", "LogonHours", "UserPrincipalName" | Foreach-Object {
        $cur = $_  
        if ($start -ne "") {
            if (($cur.whenChanged -lt $starttime) -and ($cur.lastlogondate -lt $starttime)){
            Write-Host "skip " $cur.Name
            return
            }
            Write-Host "write " $cur.Name

        }

        $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"

        if ($null -ne $cur.thumbnailPhoto) {
            $cur.thumbnailPhoto =[Convert]::ToBase64String($cur.thumbnailPhoto)
        }

        $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force

        $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
        $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

        store $cur
    }

    Write-Host "users inspect finished"
}


while ($true)
{
	if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character))
    {
        Write-Host "You pressed CTRL-C. Do you want to continue doing this and that?" 
        $key = $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
        if ($key.Character -eq "N") { break; }
    }
    $nextStart = Get-Date
    inspectComputers
    inspectComputers
    inspectUsers
    $start = $nextStart
    Start-Sleep -Seconds $timeout
}

