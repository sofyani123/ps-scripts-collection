# Import the IText7Module
Import-Module IText7Module

# Define the directory path
$directoryPath = "S:\Business Change and Infrastructure\IT\Reports"
$destinationBasePath = "S:\Business Change and Infrastructure\IT\Reports\Weekly Service Report"

# Function to read and extract the date from PDF contents
function Extract-DateFromPdf {
    param (
        [string]$pdfPath
    )

    # Create a PdfReader object
    $pdfReader = New-Object iText.Kernel.Pdf.PdfReader($pdfPath)

    # Create a PdfDocument object
    $pdfDocument = New-Object iText.Kernel.Pdf.PdfDocument($pdfReader)

    # Variable to store the extracted date
    $extractedDate = $null

    # Iterate through each page
    for ($pageNum = 1; $pageNum -le $pdfDocument.GetNumberOfPages(); $pageNum++) {
        # Get the page
        $page = $pdfDocument.GetPage($pageNum)

        # Extract text from the page
        $text = [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page)

        # Define the regex pattern for the date format
        $datePattern = "Date:\s(\d{1,2})\s(\w+)\s(\d{4})"

        # Find the first match for the date pattern in the text
        $dateMatch = [regex]::Match($text, $datePattern)

        if ($dateMatch.Success) {
            # Extract the date components
            $day = $dateMatch.Groups[1].Value
            $month = $dateMatch.Groups[2].Value
            $year = $dateMatch.Groups[3].Value

            # Convert month name to month number
            $monthNumber = (Get-Culture).DateTimeFormat.MonthNames.IndexOf($month) + 1

            # Format the date as DDMMYYYY
            $extractedDate = "{0:D2}{1:D2}{2}" -f [int]$day, $monthNumber, $year
            break
        }
    }

    # Close the PdfDocument
    $pdfDocument.Close()

    return $extractedDate
}

# Get all PDF files that match the new pattern
$pdfFiles = Get-ChildItem -Path $directoryPath -Filter "Weekly Service Report_Societe Generale Equipment*.pdf"

# Iterate through each PDF file and process its contents
foreach ($pdfFile in $pdfFiles) {
    $pdfPath = $pdfFile.FullName
    Write-Output "Processing file: $pdfPath"

    # Extract the date from the PDF
    $extractedDate = Extract-DateFromPdf -pdfPath $pdfPath

    if ($extractedDate) {
        # Define the new file name
        $newFileName = "Weekly Service Report_Societe Generale Equipment $extractedDate.pdf"
        $newFilePath = Join-Path (Get-Item $pdfPath).DirectoryName $newFileName

        # Rename the file
        Rename-Item -Path $pdfPath -NewName $newFileName
        Write-Output "Renamed $pdfPath to $newFileName"

        # Extract date components for the target directory
        $day = $extractedDate.Substring(0, 2)
        $month = $extractedDate.Substring(2, 2)
        $year = $extractedDate.Substring(4, 4)

        # Convert month number to month name
        $monthNumber = [int]$month
        $monthName = (Get-Culture).DateTimeFormat.MonthNames[$monthNumber - 1]

        # Define the target directory based on the extracted date
        $destinationYearPath = Join-Path $destinationBasePath $year
        $destinationMonthPath = Join-Path $destinationYearPath "${monthNumber}. ${monthName}"

        # Create the target directory if it doesn't exist
        if (-not (Test-Path -Path $destinationYearPath)) {
            New-Item -ItemType Directory -Path $destinationYearPath | Out-Null
        }
        if (-not (Test-Path -Path $destinationMonthPath)) {
            New-Item -ItemType Directory -Path $destinationMonthPath | Out-Null
        }

        # Move the file to the target directory
        $destinationFilePath = Join-Path $destinationMonthPath $newFileName
        Move-Item -Path $newFilePath -Destination $destinationFilePath
        Write-Output "Moved $newFileName to $destinationFilePath"
    } else {
        Write-Output "No valid date found in $pdfPath"
    }
}
