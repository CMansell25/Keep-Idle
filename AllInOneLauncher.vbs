Set objShell = CreateObject("Wscript.Shell")
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command ""powercfg -change -standby-timeout-ac 0; while ($true) { $wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('{CAPSLOCK}'); Start-Sleep -Milliseconds 500; $wshell.SendKeys('{CAPSLOCK}'); Start-Sleep -Seconds 300 }""", 0
