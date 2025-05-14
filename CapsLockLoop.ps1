Add-Type @"
using System;
using System.Runtime.InteropServices;

public class UserSession {
    [DllImport("user32.dll")]
    public static extern bool OpenInputDesktop(uint dwFlags, bool fInherit, uint dwDesiredAccess);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr OpenDesktop(string lpszDesktop, uint dwFlags, bool fInherit, uint dwDesiredAccess);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SwitchDesktop(IntPtr hDesktop);
}
"@

function Test-WorkstationLocked {
    try {
        $hDesktop = [UserSession]::OpenDesktop("default", 0, $false, 0x0100)
        if ($hDesktop -eq [IntPtr]::Zero) {
            return $true
        }
        $result = [UserSession]::SwitchDesktop($hDesktop)
        return -not $result
    }
    catch {
        return $true
    }
}

# Disable sleep
powercfg -change -standby-timeout-ac 0

while ($true) {
    if (Test-WorkstationLocked) {
        exit
    }
    $wshell = New-Object -ComObject wscript.shell
    $wshell.SendKeys("{CAPSLOCK}")
    Start-Sleep -Milliseconds 500
    $wshell.SendKeys("{CAPSLOCK}")
    Start-Sleep -Seconds 300
}
