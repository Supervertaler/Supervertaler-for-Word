# Registers (or removes) the Supervertaler for Word spike as a Word developer add-in.
#   powershell -NoProfile -File register.ps1            register
#   powershell -NoProfile -File register.ps1 -Remove    remove
# The registry write runs through Task Scheduler so it lands in the real HKCU
# even when this script is launched from a sandboxed or packaged shell. The
# command is put in a .cmd file first because the repo path contains a space.
param([switch]$Remove)
$manifest = Join-Path $PSScriptRoot "manifest.xml"
$key = "HKCU\Software\Microsoft\Office\16.0\WEF\Developer"
$cmdFile = Join-Path $PSScriptRoot "_register.cmd"
$line = if ($Remove) { "reg delete `"$key`" /v svword /f" } else { "reg add `"$key`" /v svword /t REG_SZ /d `"$manifest`" /f" }
Set-Content -Path $cmdFile -Value $line -Encoding ASCII
schtasks /create /f /tn svword_register /sc once /st 23:59 /tr "`"$cmdFile`"" | Out-Null
schtasks /run /tn svword_register | Out-Null
Start-Sleep -Seconds 3
schtasks /delete /f /tn svword_register | Out-Null
Remove-Item $cmdFile -ErrorAction SilentlyContinue
if ($Remove) { "Removed svword from $key" } else { "Registered $manifest under $key" }
