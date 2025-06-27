# Import the IText7Module
Import-Module IText7Module

# Define the directory path
$directoryPath = "S:\Business Change and Infrastructure\IT\Reports"
$destinationBasePath = "S:\Business Change and Infrastructure\IT\Reports\Veeam Backup Reports"

# Function to read and check PDF contents
function Check-PdfContent {
    param (
        [string]$pdfPath
    )

    # Extract the date from the filename
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
    $datePattern = "_(\d{4})_(\d{1,2})_(\d{1,2})$"
    $dateMatch = [regex]::Match($fileName, $datePattern)

    if ($dateMatch.Success) {
        # Extract the date components
        $year = $dateMatch.Groups[1].Value
        $month = $dateMatch.Groups[2].Value
        $day = $dateMatch.Groups[3].Value

        # Convert month number to month name
        $monthNumber = [int]$month
        $monthName = (Get-Culture).DateTimeFormat.MonthNames[$monthNumber - 1]

        # Create a PdfReader object
        $pdfReader = New-Object iText.Kernel.Pdf.PdfReader($pdfPath)

        # Create a PdfDocument object
        $pdfDocument = New-Object iText.Kernel.Pdf.PdfDocument($pdfReader)

        # Variable to track success
        $isSuccess = $false

        # Define the server names to look for
        $serverNames = @("bcasgfpctx02", "BCASGFPFIL02", "BCASGFPFL01", "BCASGFFL01")

        # Iterate through each page
        for ($pageNum = 1; $pageNum -le $pdfDocument.GetNumberOfPages(); $pageNum++) {
            # Get the page
            $page = $pdfDocument.GetPage($pageNum)

            # Extract text from the page
            $text = [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page)

            # Check if any of the server names are followed by six entries with the sixth being "Success"
            foreach ($serverName in $serverNames) {
                $pattern = "(?i)${serverName} (?:[^\n]*? ){5}Success"
                if ($text -match $pattern) {
                    $isSuccess = $true
                    Write-Output "Found success pattern for '${serverName}' in ${pdfPath}, Page ${pageNum}"
                    break
                }
            }

            # If success is found, no need to check further pages
            if ($isSuccess) {
                break
            }
        }

        # Close the PdfDocument
        $pdfDocument.Close()

        # Rename the file if it indicates an error
        if (-not $isSuccess) {
            $newFileName = "${fileName} - Error.pdf"
            $newFilePath = Join-Path $directoryPath $newFileName
            Rename-Item -Path $pdfPath -NewName $newFileName
            Write-Output "Renamed ${pdfPath} to ${newFileName}"
        } else {
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
            $destinationFilePath = Join-Path $destinationMonthPath "${fileName}.pdf"
            Move-Item -Path $pdfPath -Destination $destinationFilePath
            Write-Output "Moved ${pdfPath} to ${destinationFilePath}"
        }

        # Output the result
        if ($isSuccess) {
            Write-Output "The PDF content indicates success."
        } else {
            Write-Output "The PDF content indicates an error."
        }
    } else {
        Write-Output "No valid date found in the filename ${pdfPath}"
    }
}

# Get all PDF files that match the pattern
$pdfFiles = Get-ChildItem -Path $directoryPath -Filter "SGEF Daily Protection Status_*.pdf"

# Iterate through each PDF file and process its contents
foreach ($pdfFile in $pdfFiles) {
    Check-PdfContent -pdfPath $pdfFile.FullName
}
