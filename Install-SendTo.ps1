# Install Convert & Merge to PDF in Send To menu
# Run as Administrator

$sendToPath = [Environment]::GetFolderPath('SendTo')
$scriptPath = "C:\Tools\Convert-and-Merge-to-PDF-SendTo.ps1"
$shortcutPath = Join-Path $sendToPath "Convert & Merge to PDF.lnk"

# Check if script exists
if (-not (Test-Path $scriptPath)) {
    Write-Host "Error: Script not found at $scriptPath" -ForegroundColor Red
    Write-Host "Please copy Convert-and-Merge-to-PDF-SendTo.ps1 to C:\Tools\ first" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Create shortcut
$WScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $WScriptShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -File `"$scriptPath`""
$shortcut.IconLocation = "shell32.dll,134"
$shortcut.Description = "Convert and merge selected files to PDF"
$shortcut.Save()

Write-Host "Successfully installed 'Convert & Merge to PDF' in Send To menu" -ForegroundColor Green
Write-Host "`nUsage:" -ForegroundColor Cyan
Write-Host "1. Select files in Windows Explorer" -ForegroundColor White
Write-Host "2. Right-click -> Send to -> Convert & Merge to PDF" -ForegroundColor White

Read-Host "`nPress Enter to exit"
