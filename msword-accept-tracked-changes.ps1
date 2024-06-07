Param(
    # Define the -source parameter with a default value of %userprofile%\Documents\msword-accept-tracked-changes
    [string]$source="$env:USERPROFILE\Documents\msword-accept-tracked-changes"
)

# Define source and destination folders
$srcFolder = "$source"
$dstFolder = "$srcFolder\export"

# Test if the source folder exists and exit if it does not
if (-Not (Test-Path -Path "$srcFolder")) {
    Write-Host "The source folder does not exist: $srcFolder"
    Exit 1
}

# Create destination folder if it does not exist
if (-Not (Test-Path -Path "$dstFolder")) {
    Write-Host "Creating export folder: $dstFolder"
    New-Item -ItemType Directory -Path "$dstFolder" | Out-Null
}

# Initialize Word Application
$word = New-Object -ComObject Word.Application
$word.Visible = $false

function Accept-RevisionsAndSaveAsPDF {
    param (
        [string]$docPath,
        [string]$pdfPath
    )

    try {
        # Open the document
        $doc = $word.Documents.Open($docPath)
        
        # Accept all revisions
        $doc.AcceptAllRevisions()

        # Save as PDF
        $doc.SaveAs([ref] $pdfPath, [ref] 17) # 17 is wdFormatPDF

        # Close the document
        $doc.Close([ref] $false)
        
        Write-Host "Converted $docPath to PDF successfully"
    } catch {
        Write-Host "Error converting $docPath to PDF: $_"

        # Ensure the document is closed
        $doc.Close([ref] $false)
    }
}

# Process each .docx file in the source folder
Get-ChildItem -Path "$srcFolder" -Filter *.docx | ForEach-Object {
    $docPath = $_.FullName
    $pdfPath = Join-Path -Path "$dstFolder" -ChildPath "$($_.BaseName).pdf"
    Accept-RevisionsAndSaveAsPDF -docPath "$docPath" -pdfPath "$pdfPath"
}

# Quit Word Application
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)

Write-Host "Processing complete."
 
