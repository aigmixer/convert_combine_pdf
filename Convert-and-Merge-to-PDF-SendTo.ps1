# Convert and Merge to PDF (Send To version)
# Requires: ImageMagick, Ghostscript, and PDFtk installed
param(
    [Parameter(Mandatory=$true, ValueFromRemainingArguments=$true)]
    [string[]]$Files
)

# Configuration
$tempDir = Join-Path $env:TEMP "pdf-merge-$(Get-Date -Format 'yyyyMMddHHmmss')"
$outputDir = Split-Path $Files[0] -Parent
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outputFile = Join-Path $outputDir "Merged-$timestamp.pdf"

# Create temp directory
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

Write-Host "Converting and merging $($Files.Count) file(s) to PDF..." -ForegroundColor Cyan

# Array to store temporary PDF paths
$pdfFiles = @()

# Process each file
foreach ($file in $Files) {
    $ext = [System.IO.Path]::GetExtension($file).ToLower()
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $tempPdf = Join-Path $tempDir "$basename.pdf"
    
    Write-Host "Processing: $([System.IO.Path]::GetFileName($file))" -ForegroundColor Yellow
    
    try {
        switch ($ext) {
            { $_ -in '.pdf' } {
                # Already PDF, just copy
                Copy-Item $file $tempPdf
                $pdfFiles += $tempPdf
            }
            { $_ -in '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.webp' } {
                # Convert image to PDF using ImageMagick
                & magick convert "$file" -density 300 -quality 95 "$tempPdf" 2>&1 | Out-Null
                
                if (Test-Path $tempPdf) {
                    $pdfFiles += $tempPdf
                } else {
                    Write-Host "  Failed to convert image" -ForegroundColor Red
                }
            }
            { $_ -in '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx' } {
                # Convert Office documents using COM objects
                Write-Host "  Converting Office document..." -ForegroundColor Yellow
                
                if ($ext -in '.doc', '.docx') {
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $doc = $word.Documents.Open($file)
                    $doc.SaveAs([ref]$tempPdf, [ref]17)
                    $doc.Close()
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                    $pdfFiles += $tempPdf
                }
                elseif ($ext -in '.xls', '.xlsx') {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $wb = $excel.Workbooks.Open($file)
                    $wb.ExportAsFixedFormat(0, $tempPdf)
                    $wb.Close()
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    $pdfFiles += $tempPdf
                }
                elseif ($ext -in '.ppt', '.pptx') {
                    $ppt = New-Object -ComObject PowerPoint.Application
                    $ppt.Visible = $false
                    $presentation = $ppt.Presentations.Open($file)
                    $presentation.SaveAs($tempPdf, 32)
                    $presentation.Close()
                    $ppt.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
                    $pdfFiles += $tempPdf
                }
            }
            { $_ -in '.txt' } {
                # Convert text to PDF using Ghostscript
                Write-Host "  Converting text file..." -ForegroundColor Yellow
                
                $psFile = Join-Path $tempDir "$basename.ps"
                $content = (Get-Content $file -Raw) -replace '[()]', ' '
                
                @"
%!PS-Adobe-3.0
/Courier findfont 10 scalefont setfont
72 720 moveto
($content) show
showpage
"@ | Out-File -FilePath $psFile -Encoding ASCII
                
                & gswin64c.exe -dBATCH -dNOPAUSE -sDEVICE=pdfwrite -sOutputFile="$tempPdf" "$psFile" 2>&1 | Out-Null
                
                if (Test-Path $tempPdf) {
                    $pdfFiles += $tempPdf
                }
            }
            default {
                Write-Host "  Unsupported file type: $ext" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "  Error processing file: $_" -ForegroundColor Red
    }
}

# Merge PDFs using PDFtk
if ($pdfFiles.Count -eq 0) {
    Write-Host "No files were successfully converted to PDF" -ForegroundColor Red
    Remove-Item $tempDir -Recurse -Force
    Read-Host "Press Enter to exit"
    exit 1
}
elseif ($pdfFiles.Count -eq 1) {
    # Only one file, just copy it
    Copy-Item $pdfFiles[0] $outputFile
    Write-Host "Single PDF created: $outputFile" -ForegroundColor Green
}
else {
    Write-Host "Merging $($pdfFiles.Count) PDF files..." -ForegroundColor Cyan
    
    try {
        # Use PDFtk to merge
        $pdftkArgs = $pdfFiles + @('cat', 'output', $outputFile)
        & pdftk @pdftkArgs 2>&1 | Out-Null
        
        if (Test-Path $outputFile) {
            Write-Host "Successfully merged to: $outputFile" -ForegroundColor Green
        } else {
            Write-Host "Failed to create merged PDF" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error merging PDFs: $_" -ForegroundColor Red
    }
}

# Cleanup temp directory
Remove-Item $tempDir -Recurse -Force

# Open the output file
if (Test-Path $outputFile) {
    Start-Process $outputFile
}

Write-Host "`nDone! Press Enter to close..." -ForegroundColor Cyan
Read-Host
