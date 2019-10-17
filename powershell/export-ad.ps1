param (
    [string]$base = 'DC=acme,DC=local',
    [string]$server = 'acme.local',
    [string]$outfilename = 'export_ad',
    [string]$user = "",
    [string]$pwd = "",
    [switch]$force = $false,
    [string]$start = ""
 )

Write-Host "base: " $base
Write-Host "server: " $server
Write-Host "outfile: " $outfile
Write-Host "user: " $user
Write-Host "pwd: " $pwd
#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm

Import-Module ActiveDirectory

$SearchBase = $base 
$outfile = "$($outfilename)_$LogDate.json"

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

Write-Host "domain: " $domain.NetBIOSName

$domain | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append


if ($start -ne "") {
  Write-Host "start: " $start
  $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
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
          $add = $tmpGrp 
          $groups +=  $tmpGrp           
          $groups = Get-ADPrincipalGroupMembershipRecursive $groupDsn $groups
      }
  }

  return $groups
}

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
  if ($licensies -ne $Null) {
    Write-Host $cur.DNSHostName " : " $($licensies)
    $cur | Add-Member -MemberType NoteProperty -Name OperatingSystemLicensies -Value $licensies -Force
  }

  Try {
    $userprofiles = Get-WmiObject -Credential $GetAdminact -Class win32_userprofile -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object sid, localpath 
    if ($userprofiles -ne $null) {
      Write-Host $cur.DNSHostName  " : " $userprofiles
      $cur | Add-Member -MemberType NoteProperty -Name UserProfiles -Value $userprofiles -Force
    }    
  } Catch {
    Write-Host $cur.DNSHostName  " : " "$($_.Exception.Message)"
  }

  Try {
    $apps = Get-WMIObject -Class win32_product -Credential $GetAdminact -ComputerName $cur.DNSHostName -ErrorAction SilentlyContinue | Select-Object Name, Version
    if ($apps -ne $Null) {
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
          $UninstallKeys.GetSubKeyNames()|%{
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
                            $UninstallKeysWow6432Node.GetSubKeyNames()|%{
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

  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}


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

  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "groups export finished to: " $outfile

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

  if ($cur.thumbnailPhoto -ne $null) {
    $cur.thumbnailPhoto =[Convert]::ToBase64String($cur.thumbnailPhoto)
  }

  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force

  $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
  $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "users export finished to: " $outfile


# | Select-Object "Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
#"sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
#"PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "lastlogondate", "ObjectClass"



Write-Host "computers export finished to: " $outfile
