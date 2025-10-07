Write-Host "Converting syllabi from .docx to .pdf..."

# Locate pandoc bundled with Quarto
$pandocPath = "C:\Users\kibby\AppData\Local\Programs\Quarto\bin\tools\pandoc.exe"

# Fallback to system pandoc if Quarto's not found
if (-Not (Test-Path $pandocPath)) {
    $pandocCmd = Get-Command pandoc -ErrorAction SilentlyContinue
    if ($pandocCmd) {
        $pandocPath = $pandocCmd.Source
    } else {
        Write-Host "⚠️ Could not find pandoc.exe in Quarto or system PATH."
        exit 1
    }
}

Write-Host "Using Pandoc at: $pandocPath"

# Convert all .docx syllabi in the syllabi/ folder
$syllabiDir = Join-Path $PSScriptRoot "syllabi"
Get-ChildItem -Path $syllabiDir -Filter "*.docx" | ForEach-Object {
    $docx = $_.FullName
    $pdf = [System.IO.Path]::ChangeExtension($docx, ".pdf")
    Write-Host " - Converting $($_.Name) → $(Split-Path $pdf -Leaf)"
    & $pandocPath "$docx" -o "$pdf"
}

Write-Host "✅ Done converting syllabi."
