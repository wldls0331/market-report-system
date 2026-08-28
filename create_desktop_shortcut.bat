@echo off
setlocal EnableExtensions
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$desktop = [Environment]::GetFolderPath('Desktop');" ^
  "$target = Join-Path (Get-Location) 'run.bat';" ^
  "$path = Join-Path $desktop 'Market Report.lnk';" ^
  "$shell = New-Object -ComObject WScript.Shell;" ^
  "$shortcut = $shell.CreateShortcut($path);" ^
  "$shortcut.TargetPath = $target;" ^
  "$shortcut.WorkingDirectory = (Get-Location).Path;" ^
  "$shortcut.WindowStyle = 7;" ^
  "$shortcut.Description = 'Market Report';" ^
  "$py = (Get-Command python -ErrorAction SilentlyContinue).Source;" ^
  "if ($py) { $shortcut.IconLocation = $py + ',0' };" ^
  "$shortcut.Save();" ^
  "Write-Output $path"

if errorlevel 1 (
  echo Failed to create the desktop shortcut.
  exit /b 1
)

echo Desktop shortcut created: Market Report
exit /b 0
