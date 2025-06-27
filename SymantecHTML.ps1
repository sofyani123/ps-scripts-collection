# Define the source directory
$sourceDirectory = "S:\Business Change and Infrastructure\IT\Reports"

# Define the target base directory
$targetBaseDirectory = "S:\Business Change and Infrastructure\IT\Reports\Symantec"

# Get all HTML files in the source directory
$htmlFiles = Get-ChildItem -Path $sourceDirectory -Filter "*.html"

# Loop through each HTML file
foreach ($file in $htmlFiles) {
    # Get the file name
    $fileName = $file.Name

    # Check if the file name matches the specified patterns
    if ($fileName -match "Administrator Daily Summary Report _IT__" -or $fileName -match "Daily Security Status Report_") {
        # Use a regular expression to extract the date from the file name
        if ($fileName -match "(\d{2}-[A-Za-z]{3}-\d{4})") {
            $date = $matches[1]

            # Output the file name and date
            Write-Output "File: $fileName"
            Write-Output "Date: $date"

            # Parse the date to determine the target directory
            $parsedDate = [DateTime]::ParseExact($date, "dd-MMM-yyyy", $null)
            $year = $parsedDate.Year
            $monthNumber = $parsedDate.Month
            $monthName = $parsedDate.ToString("MMMM")

            # Define the target directory with the correct month format
            $targetDirectory = Join-Path -Path $targetBaseDirectory -ChildPath "$year\$monthNumber. $monthName"

            # Ensure the target directory exists
            if (-not (Test-Path -Path $targetDirectory)) {
                New-Item -ItemType Directory -Path $targetDirectory
                Write-Output "Created directory: $targetDirectory"
            }

            # Move the file to the target directory
            $destinationPath = Join-Path -Path $targetDirectory -ChildPath $fileName
            Move-Item -Path $file.FullName -Destination $destinationPath

            Write-Output "Moved $fileName to $destinationPath"
            Write-Output "---------------------"
        } else {
            Write-Output "No date found in file name: $fileName"
        }
    } else {
        Write-Output "Skipping file: $fileName (does not match the specified patterns)"
    }
}
