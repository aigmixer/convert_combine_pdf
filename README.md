# Convert & Merge to PDF

A Windows PowerShell tool to convert and merge multiple files (images, PDFs, Office documents) into a single PDF via the "Send To" context menu.

![License](https://img.shields.io/badge/license-GPL%20v3-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)

## Features

- **Convert multiple file types to PDF**: Images (JPG, PNG, BMP, GIF, TIFF, WEBP), Office documents (DOCX, XLSX, PPTX), text files
- **Merge into single PDF**: Combines all selected files in order
- **Windows Explorer integration**: Simple "Send To" menu access
- **No timing issues**: Native Windows file handling
- **Open source**: GPL v3 licensed, uses open-source dependencies

## Screenshots

### Usage
```
Select files → Right-click → Send to → Convert & Merge to PDF
```

### Output
```
Converting and merging 5 file(s) to PDF...
Processing: photo1.jpg
Processing: document.pdf
Processing: scan.png
Processing: report.docx
Processing: data.xlsx
Merging 5 PDF files...
Successfully merged to: D:\Folder\Merged-20251006-143022.pdf
```

## Requirements

### Open Source Dependencies

1. **ImageMagick** - Image conversion
   - Download: https://imagemagick.org/script/download.php#windows
   - License: Apache 2.0

2. **Ghostscript** - PDF processing
   - Download: https://ghostscript.com/releases/gsdnld.html
   - License: AGPL

3. **PDFtk Free** - PDF merging
   - Download: https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/
   - License: GPL v2

4. **Pandoc** - Document conversion
   - Download: https://pandoc.org/installing.html
   - License: GPL v2
   - Requires: LaTeX distribution (MiKTeX or TeX Live) for PDF output

### Operating System

- Windows 10 or Windows 11
- PowerShell 5.1 or later (built-in)

## Installation

### Quick Install

1. **Install dependencies** (ImageMagick, Ghostscript, PDFtk Free)

2. **Download and extract** this repository

3. **Copy script to C:\Tools**
   ```powershell
   New-Item -ItemType Directory -Path "C:\Tools" -Force
   Copy-Item "Convert-and-Merge-to-PDF-SendTo.ps1" "C:\Tools\"
   ```

4. **Run installer**
   - Right-click `Install-SendTo.ps1` → Run with PowerShell

### Manual Install

1. Press `Win+R`, type `shell:sendto`, press Enter
2. Create shortcut with these properties:
   - **Target**: `powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Normal -File "C:\Tools\Convert-and-Merge-to-PDF-SendTo.ps1"`
   - **Name**: `Convert & Merge to PDF`

## Usage

1. **Select files** in Windows Explorer (use Ctrl+Click for multiple files)
2. **Right-click** → **Send to** → **Convert & Merge to PDF**
3. PowerShell window shows progress
4. Merged PDF opens automatically when complete

### Supported File Types

| Type | Extensions | Notes |
|------|-----------|-------|
| Images | JPG, JPEG, PNG, BMP, GIF, TIFF, TIF, WEBP | Converted at 300 DPI |
| PDFs | PDF | Merged as-is |
| Office | DOC, DOCX, XLS, XLSX, PPT, PPTX | Via Pandoc |
| OpenDocument | ODT, ODS, ODP | Via Pandoc |
| Other Documents | RTF, EPUB, HTML, HTM | Via Pandoc |
| Text | TXT | Basic conversion |

### Output

Files are created in the same folder as the first selected file with timestamp:
```
Merged-YYYYMMDD-HHMMSS.pdf
```

## Customization

### Change Output Location

Edit `Convert-and-Merge-to-PDF-SendTo.ps1`:
```powershell
# Change this line:
$outputDir = Split-Path $Files[0] -Parent

# To fixed location:
$outputDir = "C:\Users\YourName\Documents\PDFs"
```

### Change Output Filename

```powershell
# Change this line:
$outputFile = Join-Path $outputDir "Merged-$timestamp.pdf"

# To custom name:
$outputFile = Join-Path $outputDir "Combined-$timestamp.pdf"
```

### Adjust Image Quality

```powershell
# Find this line (around line 42):
& magick convert "$file" -density 300 -quality 95 "$tempPdf"

# Adjust values:
# -density: DPI (150=low, 300=high, 600=very high)
# -quality: JPEG quality (75=good, 95=excellent, 100=maximum)
```

### Use qpdf Instead of PDFtk

Replace merge section:
```powershell
# Replace:
$pdftkArgs = $pdfFiles + @('cat', 'output', $outputFile)
& pdftk @pdftkArgs 2>&1 | Out-Null

# With:
$qpdfArgs = @('--empty', '--pages') + $pdfFiles + @('--', $outputFile)
& qpdf @qpdfArgs 2>&1 | Out-Null
```

Install qpdf: http://qpdf.sourceforge.net/

## Uninstallation

Run `Uninstall-SendTo.ps1` or manually delete:
```
%APPDATA%\Microsoft\Windows\SendTo\Convert & Merge to PDF.lnk
```

## Troubleshooting

### "Cannot find ImageMagick/Ghostscript/PDFtk/Pandoc"
- Verify installation: `magick --version`, `gswin64c --version`, `pdftk --version`, `pandoc --version`
- Restart PowerShell/Explorer after installing tools
- Ensure tools are in system PATH

### Execution Policy Errors
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Document Conversion Fails
- Install LaTeX distribution (required by Pandoc for PDF output)
  - **MiKTeX**: https://miktex.org/download (recommended for Windows)
  - **TeX Live**: https://www.tug.org/texlive/
- Verify Pandoc can access LaTeX: `pandoc --pdf-engine=xelatex --version`
- First conversion may be slow (LaTeX downloads packages)

### 32-bit Ghostscript
Edit script, change `gswin64c.exe` to `gswin32c.exe`

### Send To Menu Item Missing
- Verify script exists: `C:\Tools\Convert-and-Merge-to-PDF-SendTo.ps1`
- Check SendTo folder: Press `Win+R`, type `shell:sendto`
- Verify shortcut target path is correct

## Why "Send To" Instead of Context Menu?

**Advantages:**
- ✅ Zero timing/race conditions
- ✅ Windows natively passes all files to single instance
- ✅ 100% reliable multi-file handling
- ✅ Simpler code (no coordination logic)

**Trade-off:**
- ❌ One extra click (Send To submenu)

The context menu registry approach launches one process per file, requiring complex coordination. Send To is simpler and more reliable.

## Architecture

### How It Works

1. Windows passes all selected files as arguments to single PowerShell process
2. Script converts each file to PDF (temp directory)
3. PDFtk merges all PDFs into one
4. Output saved to source folder with timestamp
5. Temp files cleaned up
6. Result opens in default PDF viewer

### File Flow

```
Input Files → Conversion → Temp PDFs → Merge → Output PDF
  ↓              ↓           ↓           ↓         ↓
.jpg          ImageMagick  file1.pdf  PDFtk   Merged.pdf
.png          ImageMagick  file2.pdf    ↓
.pdf          Copy         file3.pdf    ↓
.docx         MS Word      file4.pdf  ──┘
.xlsx         MS Excel     file5.pdf
```

## Contributing

Contributions welcome! Areas for improvement:
- Additional file format support
- Better error handling
- Progress bar/GUI wrapper
- Batch processing options
- PDF optimization/compression

## License

GPL v3 - See [LICENSE](LICENSE) file

This project uses:
- PDFtk Server (GPL v2)
- Ghostscript (AGPL)
- ImageMagick (Apache 2.0)

## Related Projects

- [PDFtk Server](https://www.pdflabs.com/tools/pdftk-server/)
- [ImageMagick](https://imagemagick.org/)
- [Ghostscript](https://www.ghostscript.com/)

## Support

For issues or questions, please open a GitHub issue.

---

**Note**: This tool requires local installation of open-source dependencies. No data is sent to external servers - all processing happens locally on your machine.
