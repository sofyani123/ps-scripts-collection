# Import the IText7Module
Import-Module IText7Module

# Define the directory path
$directoryPath = "S:\Business Change and Infrastructure\IT\Reports"
$destinationBasePath = "S:\Business Change and Infrastructure\IT\Reports\IPS Reports"

# Function to process PDF contents
function Process-PdfContent {
    param (
        [string]$pdfPath
    )

    # Extract the date from the filename
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
    $datePattern = "_(\w{3})_(\d{1,2})__(\d{4})_"
    $dateMatch = [regex]::Match($fileName, $datePattern)

    if ($dateMatch.Success) {
        # Extract the date components
        $monthAbbr = $dateMatch.Groups[1].Value
        $day = $dateMatch.Groups[2].Value
        $year = $dateMatch.Groups[3].Value

        # Convert month abbreviation to full month name
        $monthNames = @("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        $monthAbbreviations = @("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        $monthNumber = [array]::IndexOf($monthAbbreviations, $monthAbbr) + 1
        $monthName = $monthNames[$monthNumber - 1]

        # Define the phrases to look for
        $phrases = @("Detect", "Prevent")

        # Variable to track if any phrase is found
        $phraseFound = $false
        $errorPhraseFound = $false

        # Create a PdfReader object
        $pdfReader = New-Object iText.Kernel.Pdf.PdfReader($pdfPath)

        # Create a PdfDocument object
        $pdfDocument = New-Object iText.Kernel.Pdf.PdfDocument($pdfReader)

        # Iterate through each page
        for ($pageNum = 1; $pageNum -le $pdfDocument.GetNumberOfPages(); $pageNum++) {
            # Get the page
            $page = $pdfDocument.GetPage($pageNum)

            # Extract text from the page
            $text = [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page)

            # Check if any of the phrases are in the text
            foreach ($phrase in $phrases) {
                if ($text -match $phrase) {
                    Write-Output "Found phrase '${phrase}' in ${pdfPath}, Page ${pageNum}"
                    $phraseFound = $true
                    if ($phrase -eq "Detect") {
                        $errorPhraseFound = $true
                    }
                }
            }
        }

        # Close the PdfDocument
        $pdfDocument.Close()

        # Define the new file name
        if ($errorPhraseFound) {
            $newFileName = "${fileName} - Error.pdf"
        } else {
            $newFileName = "${fileName}.pdf"
        }
        $newFilePath = Join-Path $directoryPath $newFileName

        # Rename the file
        Rename-Item -Path $pdfPath -NewName $newFileName
        Write-Output "Renamed ${pdfPath} to ${newFileName}"

        # Move the file only if it does not have " - Error" in the filename
        if (-not $newFileName.Contains(" - Error")) {
            # Define the destination directory based on the extracted date
            $destinationYearPath = Join-Path $destinationBasePath $year
            $destinationMonthPath = Join-Path $destinationYearPath "${monthNumber}. ${monthName}"

            # Create the destination directory if it doesn't exist
            if (-not (Test-Path -Path $destinationYearPath)) {
                New-Item -ItemType Directory -Path $destinationYearPath | Out-Null
            }
            if (-not (Test-Path -Path $destinationMonthPath)) {
                New-Item -ItemType Directory -Path $destinationMonthPath | Out-Null
            }

            # Move the file to the destination directory
            $destinationFilePath = Join-Path $destinationMonthPath $newFileName
            Move-Item -Path $newFilePath -Destination $destinationFilePath
            Write-Output "Moved ${newFileName} to ${destinationFilePath}"
        }
    } else {
        Write-Output "No valid date found in the filename ${pdfPath}"
    }
}

# Get all PDF files that match the pattern
$pdfFiles = Get-ChildItem -Path $directoryPath -Filter "Intrusion_Prevention_System__IPS__with_IP_Details__Domain_SGEF_VSX__*.pdf"

# Iterate through each PDF file and process its contents
foreach ($pdfFile in $pdfFiles) {
    Process-PdfContent -pdfPath $pdfFile.FullName
}
