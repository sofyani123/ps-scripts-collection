# Define the path to the directory
$directoryPath = "S:\Business Change and Infrastructure\IT\Reports"
$targetBasePath = "S:\Business Change and Infrastructure\IT\Reports\ConnectBackup"

# Define the regular expression pattern to search for
$pattern = "(BCASGFPVIEDB01|BCCSGFTVIEDB01) nightly DB backup SUCCESS.*\.msg"

# Check if the directory exists
if (Test-Path -Path $directoryPath) {
    # Get the items in the directory
    $items = Get-ChildItem -Path $directoryPath | Where-Object { $_.Name -match $pattern }

    # Check if any matching files exist
    if ($items) {
        Write-Output "Matching files found in the directory '$directoryPath':"
        $items | ForEach-Object {
            $filePath = $_.FullName
            Write-Output "Processing file: $filePath"

            # Create an Outlook application object
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")

            # Open the .msg file
            $mailItem = $namespace.OpenSharedItem($filePath)

            # Get the date received
            $dateReceived = $mailItem.ReceivedTime

            # Format the date received
            $formattedDate = $dateReceived.ToString("ddMMyyyy")

            # Determine the new file name by removing unwanted suffixes
            $baseName = $_.BaseName -replace " nightly DB backup SUCCESS.*", " nightly DB backup SUCCESS"
            $newFileName = "$baseName $formattedDate.msg"

            # Release the COM object to close the file
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mailItem) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

            # Rename the file
            $newFilePath = Join-Path -Path $directoryPath -ChildPath $newFileName
            Rename-Item -Path $filePath -NewName $newFileName

            # Output the new file name
            Write-Output "Renamed file to: $newFileName"

            # Determine the target directory based on the date received
            $year = $dateReceived.Year
            $month = $dateReceived.Month
            $monthName = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.GetMonthName($month)
            $targetYearPath = Join-Path -Path $targetBasePath -ChildPath $year.ToString()
            $targetMonthPath = Join-Path -Path $targetYearPath -ChildPath "$month. $monthName"

            # Create the target directory if it does not exist
            if (-not (Test-Path -Path $targetMonthPath)) {
                New-Item -ItemType Directory -Path $targetMonthPath | Out-Null
            }

            # Move the file to the target directory
            $targetFilePath = Join-Path -Path $targetMonthPath -ChildPath $newFileName
            Move-Item -Path $newFilePath -Destination $targetFilePath

            # Output the moved file path
            Write-Output "Moved file to: $targetFilePath"
        }
    } else {
        Write-Output "No matching files found in the directory '$directoryPath'."
    }
} else {
    Write-Output "The directory '$directoryPath' does not exist."
}
