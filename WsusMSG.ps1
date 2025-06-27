# Import the Outlook COM object
$outlook = New-Object -ComObject Outlook.Application

# Define the directory path
$directoryPath = "S:\Business Change and Infrastructure\IT\Reports"
$destinationBasePath = "S:\Business Change and Infrastructure\IT\Reports\WSUS"

# Function to read and extract the sent date from .msg file
function Extract-SentDateFromMsg {
    param (
        [string]$msgFilePath
    )

    # Open the .msg file
    $msg = $outlook.Session.OpenSharedItem($msgFilePath)

    # Extract and return the sent date
    $sentDate = $msg.SentOn

    # Release the COM object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null

    return $sentDate
}

# Function to format the date as DDMMYYYY
function Format-Date {
    param (
        [datetime]$date
    )

    return $date.ToString("ddMMyyyy")
}

# Get all .msg files that match the patterns
$msgFiles = @()
$msgFiles += Get-ChildItem -Path $directoryPath -Filter "WSUS_ New Update(s) Alert From BCASGFPSCCM01*.msg"
$msgFiles += Get-ChildItem -Path $directoryPath -Filter "WSUS_ Update Status Summary From BCASGFPSCCM01*.msg"

# Iterate through each .msg file and process its contents
foreach ($msgFile in $msgFiles) {
    $msgFilePath = $msgFile.FullName
    Write-Output "Processing file: ${msgFilePath}"

    # Extract the sent date from the .msg file
    $sentDate = Extract-SentDateFromMsg -msgFilePath $msgFilePath

    if ($sentDate) {
        # Format the sent date as DDMMYYYY
        $formattedDate = Format-Date -date $sentDate

        # Define the new file name based on the original file name
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($msgFilePath)
        if ($baseName -like "WSUS_ New Update(s) Alert From BCASGFPSCCM01*") {
            $newFileName = "WSUS_ New Update(s) Alert From BCASGFPSCCM01 ${formattedDate}.msg"
        } elseif ($baseName -like "WSUS_ Update Status Summary From BCASGFPSCCM01*") {
            $newFileName = "WSUS_ Update Status Summary From BCASGFPSCCM01 ${formattedDate}.msg"
        }

        $newFilePath = Join-Path (Get-Item $msgFilePath).DirectoryName $newFileName

        # Rename the file
        Rename-Item -Path $msgFilePath -NewName $newFileName
        Write-Output "Renamed ${msgFilePath} to ${newFileName}"

        # Extract date components for the target directory
        $day = $formattedDate.Substring(0, 2)
        $month = $formattedDate.Substring(2, 2)
        $year = $formattedDate.Substring(4, 4)

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
        Write-Output "Moved ${newFileName} to ${destinationFilePath}"
    } else {
        Write-Output "No valid date found in ${msgFilePath}"
    }
}

# Release the Outlook COM object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
Remove-Variable -Name outlook -ErrorAction SilentlyContinue
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
