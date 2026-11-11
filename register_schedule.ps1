$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$python = (Get-Command python -ErrorAction Stop).Source
$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$action = New-ScheduledTaskAction -Execute $python -Argument ('-X utf8 "{0}" --scheduled' -f (Join-Path $root 'daily_send.py')) -WorkingDirectory $root
$trigger = New-ScheduledTaskTrigger -Daily -At '08:50'
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2) -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 5)
Register-ScheduledTask -TaskName 'YidaExcelDailySend' -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description 'Prepare Yida report at 08:50; send Excel and brief to the selected DingTalk group at 09:00 China time.' | Format-List TaskName, State
