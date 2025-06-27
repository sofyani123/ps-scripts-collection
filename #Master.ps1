# Get the current directory
$currentDirectory = (Get-Item -Path ".\").FullName

# Define the relative paths to the scripts
$relativeScriptPaths = @(
    "WeeklyServiceReportPDF.ps1",
    "WsusMSG.ps1",
    "ActifioReportPDF.ps1",
    "ConnectBackupsMSG.ps1",
    "IPSReportPDF.ps1",
    "iSeriesPDF.ps1",
    "TrendAVReportPDF.ps1",
    "SymantecHTML.ps1",
    "VeeamReportPDF.ps1",
    "WeeklyChangeReportPDF.ps1",
    "CheckForReportWithError.ps1"
)

# Convert relative paths to absolute paths
$scriptPaths = $relativeScriptPaths | ForEach-Object { Join-Path $currentDirectory $_ }

# Function to run a script and update the progress bar
function Run-Script {
    param (
        [string]$scriptPath,
        [int]$index,
        [int]$total
    )

    Write-Progress -Activity "Running Scripts" -Status "Running ${scriptPath}" -PercentComplete (($index / $total) * 100)
    & $scriptPath
}

# Infinite loop to run the scripts every minute
while ($true) {
    # Iterate through each script and run it
    for ($i = 0; $i -lt $scriptPaths.Count; $i++) {
        $scriptPath = $scriptPaths[$i]
        Run-Script -scriptPath $scriptPath -index $i -total $scriptPaths.Count
    }

    Write-Output "All scripts have been executed successfully."

    # Wait for 1 minute (60 seconds) before restarting the loop
    Start-Sleep -Seconds 60
}
