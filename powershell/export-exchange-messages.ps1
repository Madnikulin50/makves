  param (
    [string]$outfilename = '.\export-exchange-messages',
    [string]$user = "admin",
    [string]$pwd = "flvw003",
    [string]$domain   = "URD",
    [switch]$save_body = $false,
    [switch]$compliance = $false,
    [string]$start = "",
    [string]$startfn = "", ##".file-monitor.time_mark",
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
    $t = $item | Select-Object -Property * -ExcludeProperty Schema, InstanceKey, Service, EntityExtractionResult
    
    $t | Add-Member -MemberType NoteProperty -Name Type -Value "exchange-message" -Force        
    $t | Add-Member -MemberType NoteProperty -Name Forwarder -Value "exchange-items-forwarder" -Force
    
    if ($compliance -eq $true)
    {
        Try
        {
            $compliance =  Get-Compliance -File $t.Body.Text
            $t | Add-Member -MemberType NoteProperty -Name Compliance -Value $compliance -Force
        }
        Catch {
            Write-Host "Get-Compliance error:" + $PSItem.Exception.Message
        }
    }

    if ($save_body -eq $false) {
      $t = $t | Select-Object -Property * -ExcludeProperty Body
  }


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

  
  Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $email    = $_.PrimarySmtpAddress

    # load the assembly
    [void] [Reflection.Assembly]::LoadFile("Microsoft.Exchange.WebServices.dll")

    # set ref to exchange, first references 2007, 2nd is 2010 (default)
    $s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    #$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

    # use first option if you want to impersonate, otherwise, grab your own credentials
    $s.Credentials = New-Object Net.NetworkCredential($user, $pwd, $domain)
    #$s.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    #$s.UseDefaultCredentials = $true

    # discover the url from your email address
    $s.AutodiscoverUrl($email)

    Write-host $email


    # get a handle to the inbox
    $inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

    #create a property set (to let us access the body & other details not available from the FindItems call)
    $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;

    $items = $inbox.FindItems(100)

    # output unread count
    Write-Host -Text "Unread count: ",$inbox.UnreadCount

    foreach ($item in $items.Items)
    {
      # load the property set to allow us to get to the body
      $item.load($psPropertySet)
      if ($item -eq $null){
        return;
      }
      inspect($item)
    }
}