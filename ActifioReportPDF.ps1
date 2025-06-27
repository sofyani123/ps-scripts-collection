# Import the IText7Module
Import-Module IText7Module

# Define the directory path
$directoryPath = "S:\Business Change and Infrastructure\IT\Reports"
$destinationBasePath = "S:\Business Change and Infrastructure\IT\Reports\Intel Daily Backup Reports (Actifio)"

# Function to sanitize file names
function Sanitize-FileName {
    param (
        [string]$fileName
    )
    # Remove invalid characters from the file name
    $invalidChars = '[\\/:*?"<>|]'
    $sanitizedFileName = $fileName -replace $invalidChars, ''
    return $sanitizedFileName
}

# Function to read and check recovery point time from PDF contents
function Check-RecoveryPointTime {
    param (
        [string]$pdfPath
    )

    # Variable to track if an error is detected
    $isError = $false

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

        # Split the text into lines
        $lines = $text -split "`n"

        # Filter lines that contain the recovery point time
        $filteredLines = $lines | Where-Object { $_ -match "S5 Snapshot \d{4}-\d{2}-\d{2} \d{2}:\d{2} (\d{2})" }

        # Extract and check the recovery point time values
        foreach ($line in $filteredLines) {
            if ($line -match "S5 Snapshot \d{4}-\d{2}-\d{2} \d{2}:\d{2} (\d{2})") {
                $recoveryPointTime = [int]$matches[1]
                Write-Output "Recovery Point Time: $recoveryPointTime"
                if ($recoveryPointTime -gt 20) {
                    $isError = $true
                }
            }
        }
    }

    # Close the PdfDocument
    $pdfDocument.Close()

    # Extract the date from the filename
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
    $datePattern = "(\d{8})(\d{4})"
    $dateMatch = [regex]::Match($fileName, $datePattern)

    if ($dateMatch.Success) {
        # Extract the date components
        $dateString = $dateMatch.Groups[1].Value
        $timeString = $dateMatch.Groups[2].Value
        $year = $dateString.Substring(0, 4)
        $month = $dateString.Substring(4, 2)
        $day = $dateString.Substring(6, 2)

        # Convert month number to month name
        $monthNumber = [int]$month
        $monthName = (Get-Culture).DateTimeFormat.MonthNames[$monthNumber - 1]

        # Rename the file if an error is detected
        if ($isError) {
            $baseName = (Get-Item $pdfPath).BaseName
            $extension = (Get-Item $pdfPath).Extension
            $newFileName = "$baseName - Error$extension"
            $newFileName = Sanitize-FileName -fileName $newFileName
            $newFilePath = Join-Path (Get-Item $pdfPath).DirectoryName $newFileName
            Rename-Item -Path $pdfPath -NewName $newFileName
            Write-Output "Renamed $pdfPath to $newFileName"
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
            $destinationFilePath = Join-Path $destinationMonthPath "$fileName.pdf"
            Move-Item -Path $pdfPath -Destination $destinationFilePath
            Write-Output "Moved $pdfPath to $destinationFilePath"
        }
    } else {
        Write-Output "No valid date found in the filename $pdfPath"
    }
}

# Get all PDF files that match the pattern
$pdfFiles = Get-ChildItem -Path $directoryPath -Filter "SGEF_Recovery_Point_By_Application_Daily-*.pdf" | Where-Object { $_.Name -match "SGEF_Recovery_Point_By_Application_Daily-\d{8}\d{4}\.pdf" }

# Iterate through each PDF file and process its contents
foreach ($pdfFile in $pdfFiles) {
    Check-RecoveryPointTime -pdfPath $pdfFile.FullName
}
