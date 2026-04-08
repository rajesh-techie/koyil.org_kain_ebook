<#
.SYNOPSIS
    Convert DOCX file to PDF using Microsoft Word COM object

.DESCRIPTION
    Converts a .docx file to PDF format using Word's native conversion capabilities.
    Handles error conditions gracefully and releases COM object handles properly.

.PARAMETER InputPath
    Full path to the .docx file to convert (required)

.PARAMETER OutputPath
    Full path for output .pdf file (optional, defaults to same directory/name with .pdf extension)

.PARAMETER OpenPDF
    Boolean to open PDF after conversion (default: $false)

.EXAMPLE
    .\convert-docx-to-pdf.ps1 -InputPath "C:\output\priya_tirumozi.docx"
    .\convert-docx-to-pdf.ps1 -InputPath "C:\output\priya_tirumozi.docx" -OutputPath "C:\pdfs\priya_tirumozi.pdf"
    .\convert-docx-to-pdf.ps1 -InputPath "C:\output\priya_tirumozi.docx" -OpenPDF $true

.NOTES
    Requires Microsoft Word to be installed
    Author: GitHub Copilot
    Date: 2026-04-08
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="Full path to the .docx file")]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "File does not exist: $_"
        }
        if ((Get-Item $_).Extension -ne ".docx") {
            throw "File must have .docx extension"
        }
        return $true
    })]
    [string]$InputPath,

    [Parameter(Mandatory=$false, HelpMessage="Full path for output .pdf file")]
    [string]$OutputPath = "",

    [Parameter(Mandatory=$false, HelpMessage="Open PDF after conversion")]
    [boolean]$OpenPDF = $false
)

# -----------------------------------------------------------------------
# Setup
# -----------------------------------------------------------------------

# Resolve paths to full absolute paths
$InputPath = (Resolve-Path $InputPath).Path
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Default OutputPath: same directory, same name with .pdf extension
if (-not $OutputPath) {
    $inputDir = Split-Path $InputPath
    $inputName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $OutputPath = Join-Path $inputDir ($inputName + ".pdf")
}

# Ensure output path is absolute
$OutputPath = (Join-Path (Get-Location).Path $OutputPath)

# -----------------------------------------------------------------------
# Validation
# -----------------------------------------------------------------------

# Check if Word is installed
$wordInstalled = $null -ne (Get-Command winword.exe -ErrorAction SilentlyContinue)
if (-not $wordInstalled) {
    Write-Host "[$timestamp] ✗ Error: Microsoft Word is not installed." -ForegroundColor Red
    Write-Host "[$timestamp]   Please install Microsoft Office to use this converter." -ForegroundColor Yellow
    exit 1
}

# Check if output directory exists, create if not
$outputDir = Split-Path $OutputPath
if (-not (Test-Path $outputDir)) {
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Host "[$timestamp] ✓ Created output directory: $outputDir" -ForegroundColor Green
    }
    catch {
        Write-Host "[$timestamp] ✗ Error: Cannot create output directory '$outputDir'" -ForegroundColor Red
        Write-Host "[$timestamp]   Details: $_" -ForegroundColor Yellow
        exit 1
    }
}

# Check write permission
if (-not (Test-Path $outputDir -PathType Container)) {
    Write-Host "[$timestamp] ✗ Error: Output directory is not accessible: $outputDir" -ForegroundColor Red
    exit 1
}

# Get file sizes for display
$inputFileSize = (Get-Item $InputPath).Length
$inputFileSizeMB = [math]::Round($inputFileSize / 1MB, 2)

# -----------------------------------------------------------------------
# Conversion Logic
# -----------------------------------------------------------------------

Write-Host "[$timestamp] → Converting: $InputPath ($inputFileSizeMB MB)" -ForegroundColor Cyan
Write-Host "[$timestamp] → Output to: $OutputPath" -ForegroundColor Cyan

$word = $null
$doc = $null
$success = $false
$startTime = Get-Date

try {
    # Create Word COM object
    Write-Host "[$timestamp] → Launching Microsoft Word..." -ForegroundColor Gray
    $word = New-Object -ComObject Word.Application
    if ($null -eq $word) {
        throw "Failed to create Word.Application COM object"
    }

    # Hide Word window
    $word.Visible = $false
    $word.ScreenUpdating = $false

    # Open the DOCX file
    Write-Host "[$timestamp] → Opening document..." -ForegroundColor Gray
    $doc = $word.Documents.Open($InputPath, $false, $false, $false)
    if ($null -eq $doc) {
        throw "Failed to open document: $InputPath"
    }

    # Save as PDF (format ID 17 = PDF)
    Write-Host "[$timestamp] → Converting to PDF..." -ForegroundColor Gray
    $doc.SaveAs($OutputPath, 17)

    # Mark success
    $success = $true
    
    Write-Host "[$timestamp] ✓ Conversion completed successfully" -ForegroundColor Green

} catch {
    Write-Host "[$timestamp] ✗ Conversion ERROR: $_" -ForegroundColor Red
    exit 1

} finally {
    # Cleanup: Close document
    if ($null -ne $doc) {
        try {
            $doc.Close($false)
        }
        catch {
            Write-Host "[$timestamp] ! Warning: Error closing document: $_" -ForegroundColor Yellow
        }
    }

    # Cleanup: Quit Word
    if ($null -ne $word) {
        try {
            $word.Quit()
        }
        catch {
            Write-Host "[$timestamp] ! Warning: Error quitting Word: $_" -ForegroundColor Yellow
        }
    }

    # Release COM objects
    if ($null -ne $doc) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        catch {}
    }

    if ($null -ne $word) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        catch {}
    }

    # Force garbage collection to release handles
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Small delay to ensure file handle is released
    Start-Sleep -Milliseconds 500
}

# -----------------------------------------------------------------------
# Post-Conversion Verification
# -----------------------------------------------------------------------

if ($success) {
    # Verify PDF was created
    if (Test-Path $OutputPath) {
        $outputFileSize = (Get-Item $OutputPath).Length
        $outputFileSizeMB = [math]::Round($outputFileSize / 1MB, 2)
        $duration = ((Get-Date) - $startTime).TotalSeconds

        Write-Host "[$timestamp] ✓ Output: $OutputPath ($outputFileSizeMB MB)" -ForegroundColor Green
        Write-Host "[$timestamp] ✓ Conversion time: $([math]::Round($duration, 1))s" -ForegroundColor Green
        Write-Host "[$timestamp] ✓ Size reduction: $([math]::Round(($inputFileSizeMB - $outputFileSizeMB), 1)) MB" -ForegroundColor Green

        # Open PDF if requested
        if ($OpenPDF) {
            try {
                Start-Process $OutputPath
                Write-Host "[$timestamp] ✓ PDF opened in default viewer" -ForegroundColor Green
            }
            catch {
                Write-Host "[$timestamp] ! Warning: Could not open PDF viewer: $_" -ForegroundColor Yellow
            }
        }

        exit 0
    }
    else {
        Write-Host "[$timestamp] ✗ Error: PDF file was not created" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "[$timestamp] ✗ Conversion failed" -ForegroundColor Red
    exit 1
}
