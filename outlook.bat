<# : choosing screen
@echo off
setlocal

:: 1. Open Outlook in Edge in App Mode with GPU disabled
start msedge --disable-gpu --app=https://outlook.office365.com/mail/
 
:: 2. Wait 3 seconds to ensure Edge completely loads
timeout /t 3 /nobreak >nul

:: 3. Run this exact file inside PowerShell cleanly
powershell -NoExit -Command "Invoke-Expression (Get-Content '%~f0' -Raw)"
exit /b
#>

# --- EVERYTHING BELOW THIS LINE IS PURE POWERSHELL ---
try {
    Add-Type -AssemblyName System.Windows.Forms
    $screens = [System.Windows.Forms.Screen]::AllScreens
    
    Write-Host "--- DETECTED MONITORS ---" -ForegroundColor Cyan
    for ($i = 0; $i -lt $screens.Count; $i++) {
        $scr = $screens[$i]
        $primary = if ($scr.Primary) { "(Primary)" } else { "" }
        Write-Host ("[{0}] {1} - Resolution: {2}x{3} {4}" -f ($i+1), $scr.DeviceName, $scr.Bounds.Width, $scr.Bounds.Height, $primary)
    }
    
    $choice = Read-Host "Select the monitor number to send Outlook to"
    $index = [int]$choice - 1
    
    if ($index -ge 0 -and $index -lt $screens.Count) {
        $targetScreen = $screens[$index]
        $posX = $targetScreen.Bounds.X
        $posY = $targetScreen.Bounds.Y
        $width = $targetScreen.Bounds.Width
        $height = $targetScreen.Bounds.Height
        
        $wshell = New-Object -ComObject WScript.Shell
        
        $apiCode = @'
            [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
            [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
            [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
            [DllImport("user32.dll")] public static extern bool SetLayeredWindowAttributes(IntPtr hWnd, uint crKey, byte bAlpha, uint dwFlags);
'@
        $User32 = Add-Type -MemberDefinition $apiCode -Name 'User32' -Namespace 'Win32' -PassThru
        
        $proc = Get-Process msedge | Where-Object {$_.MainWindowTitle -like '*Outlook*' -and $_.MainWindowHandle -ne 0} | Select-Object -First 1
        
        if ($proc) {
            $hwnd = $proc.MainWindowHandle
            
            # Move to target screen
            [Win32.User32]::SetWindowPos($hwnd, 0, $posX, $posY, $width, $height, 0x0040)
            
            # Apply transparency (160 = semi-transparent)
            $wl = [Win32.User32]::GetWindowLong($hwnd, -20)
            [void][Win32.User32]::SetWindowLong($hwnd, -20, ($wl -bor 0x80000))
            [void][Win32.User32]::SetLayeredWindowAttributes($hwnd, 0, 220, 2)
            
            # Fullscreen
            Start-Sleep -Milliseconds 500
            $wshell.AppActivate($proc.Id)
            $wshell.SendKeys('{F11}')
            Write-Host "Success!" -ForegroundColor Green
        } else {
            Write-Host "Could not find the Outlook window. Make sure it opened completely." -ForegroundColor Red
        }
    } else {
        Write-Host "Invalid selection." -ForegroundColor Red
    }
} catch {
    Write-Error $_
}

Write-Host "Press any key to exit..." -ForegroundColor Yellow
[void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
exit