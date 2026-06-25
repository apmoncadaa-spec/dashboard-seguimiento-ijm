$dataDir = "C:\Users\FTQ\apoyoconsultoria.com\File Server - 2025-070-O IJM - Linea de base\04. Analisis\03. Programacion\06. Dashboard Seguimiento Web\data"
$batFile = Join-Path $dataDir "auto_push_data.bat"

$action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument ("/c `"" + $batFile + "`"") `
    -WorkingDirectory $dataDir

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel Highest

$trigger = New-ScheduledTaskTrigger -Daily -At "10:00AM"
Register-ScheduledTask -TaskName "Dashboard IJM - Data 10am" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force
Write-Host "Tarea creada: Dashboard IJM - Data 10am"
Write-Host "Listo."
