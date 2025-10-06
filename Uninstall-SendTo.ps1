# Uninstall Convert & Merge to PDF from Send To menu

$sendToPath = [Environment]::GetFolderPath('SendTo')
$shortcutPath = Join-Path $sendToPath "Convert & Merge to PDF.lnk"

if (Test-Path $shortcutPath) {
    Remove-Item $shortcutPath -Force
    Write-Host "Successfully removed 'Convert & Merge to PDF' from Send To menu" -ForegroundColor Green
} else {
    Write-Host "Shortcut not found in Send To menu" -ForegroundColor Yellow
}

Read-Host "Press Enter to exit"
