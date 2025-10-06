# Convert & Merge to PDF - Send To Installation

## Prerequisites

Install these open-source tools:

### 1. ImageMagick
- Download: https://imagemagick.org/script/download.php#windows
- Check "Add application directory to system path" during install
- Verify: `magick --version`

### 2. Ghostscript
- Download: https://ghostscript.com/releases/gsdnld.html
- Install 64-bit version
- Verify: `gswin64c --version`

### 3. PDFtk Free (GPL)
- Download: https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/
- Verify: `pdftk --version`

## Installation

### 1. Copy Script
```powershell
# Create directory
New-Item -ItemType Directory -Path "C:\Tools" -Force

# Copy the script
Copy-Item "Convert-and-Merge-to-PDF-SendTo.ps1" "C:\Tools\"
```

### 2. Install Send To Entry
Right-click `Install-SendTo.ps1` → Run with PowerShell

Or manually:
1. Press Win+R, type: `shell:sendto`
2. Create shortcut in that folder:
   - Target: `powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -File "C:\Tools\Convert-and-Merge-to-PDF-SendTo.ps1"`
   - Name: `Convert & Merge to PDF`

## Usage

1. Select files in Windows Explorer (Ctrl+click for multiple)
2. Right-click → **Send to** → **Convert & Merge to PDF**
3. PowerShell window shows progress
4. Merged PDF opens when complete

## Advantages Over Context Menu

- **Zero timing issues** - Windows passes all files to single instance
- **No race conditions** - Native Windows behavior
- **100% reliable** - No queue management needed
- **Simpler code** - No coordination logic required

## Disadvantages

- One extra click (Send To submenu vs direct context menu)
- Less discoverable than top-level context menu item

## Supported Formats

- Images: JPG, PNG, BMP, GIF, TIFF, WEBP
- PDFs: Merged as-is
- Office: DOCX, XLSX, PPTX (requires MS Office)
- Text: TXT

## Output

File created in same folder as first selected file:
`Merged-YYYYMMDD-HHMMSS.pdf`

## Uninstall

Run `Uninstall-SendTo.ps1`

Or manually delete: `%APPDATA%\Microsoft\Windows\SendTo\Convert & Merge to PDF.lnk`

## Troubleshooting

### Send To not appearing
- Check script exists at C:\Tools\Convert-and-Merge-to-PDF-SendTo.ps1
- Verify shortcut in SendTo folder (Win+R → `shell:sendto`)

### Execution Policy errors
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Office conversion fails
- Requires Microsoft Office installed and activated
- Try opening document manually first

### Ghostscript 32-bit
Edit script, change `gswin64c.exe` to `gswin32c.exe`
