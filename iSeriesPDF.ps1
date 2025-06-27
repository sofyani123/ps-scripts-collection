# Import the IText7Module
Import-Module IText7Module

# Define the directory path
$directoryPath = "S:\Business Change and Infrastructure\IT\Reports"
$destinationBasePath = "S:\Business Change and Infrastructure\IT\Reports\Daily Reports"

# Function to read and print PDF contents
function Read-PdfContent {
    param (
        [string]$pdfPath
    )

    # Create a PdfReader object
    $pdfReader = New-Object iText.Kernel.Pdf.PdfReader($pdfPath)

    # Create a PdfDocument object
    $pdfDocument = New-Object iText.Kernel.Pdf.PdfDocument($pdfReader)

    # Variable to store the extracted date
    $extractedDate = $null

    # Define the phrases to look for
    $phrases = @("Backup Successful", "No Backup Scheduled")

    # Variable to track if any phrase is found
    $phraseFound = $false

    # Iterate through each page
    for ($pageNum = 1; $pageNum -le $pdfDocument.GetNumberOfPages(); $pageNum++) {
        # Get the page
        $page = $pdfDocument.GetPage($pageNum)

        # Extract text from the page
        $text = [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page)

        # Define the regex pattern for the date format
        $datePattern = "Date:\s(\d{2})\s(\w+)\s(\d{4})"

        # Find the first match for the date pattern in the text
        $dateMatch = [regex]::Match($text, $datePattern)

        if ($dateMatch.Success) {
            # Extract the date components
            $day = $dateMatch.Groups[1].Value
            $month = $dateMatch.Groups[2].Value
            $year = $dateMatch.Groups[3].Value

            # Convert month name to month number
            $monthNumber = (Get-Culture).DateTimeFormat.MonthNames.IndexOf($month) + 1

            # Format the date as ddMMyyyy
            $extractedDate = "{0:D2}{1:D2}{2}" -f [int]$day, $monthNumber, $year
        }

        # Check if any of the phrases are in the text
        foreach ($phrase in $phrases) {
            if ($text -match $phrase) {
                Write-Output "Found phrase '${phrase}' in ${pdfPath}, Page ${pageNum}"
                $phraseFound = $true
            }
        }
    }

    # Close the PdfDocument
    $pdfDocument.Close()

    # Rename the file if a date was extracted
    if ($extractedDate) {
        # Define the new file name
        if ($phraseFound) {
            $newFileName = "Daily Service Report_Societe Generale Equipment ${extractedDate}.pdf"
        } else {
            $newFileName = "Daily Service Report_Societe Generale Equipment ${extractedDate} - Error.pdf"
        }
        $newFilePath = Join-Path $directoryPath $newFileName

        # Rename the file
        Rename-Item -Path $pdfPath -NewName $newFileName
        Write-Output "Renamed ${pdfPath} to ${newFileName}"

        # Move the file only if it does not have " - Error" in the filename
        if (-not $newFileName.Contains(" - Error")) {
            # Define the destination directory based on the extracted date
            $destinationYearPath = Join-Path $destinationBasePath $year
            $destinationMonthPath = Join-Path $destinationYearPath "${monthNumber}. ${month}"

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
        Write-Output "No date found in ${pdfPath}"
    }
}

# Get all PDF files that match the pattern
$pdfFiles = Get-ChildItem -Path $directoryPath -Filter "Daily Service Report_Societe Generale Equipment*.pdf"

# Iterate through each PDF file and read its contents
foreach ($pdfFile in $pdfFiles) {
    Read-PdfContent -pdfPath $pdfFile.FullName
}
