# Define the directory to check
$directoryToCheck = "S:\Business Change and Infrastructure\IT\Reports"

# Function to check for files containing "Error" in their names
function Check-ForErrorFiles {
    param (
        [string]$directory
    )
    $files = Get-ChildItem -Path $directory -File
    foreach ($file in $files) {
        if ($file.Name -like "*Error*") {
            return $true
        }
    }
    return $false
}

# Function to show the popup
function Show-Popup {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show("Report with Errors Detected", "Warning", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
}

# Function to check if the popup is already open
function Is-PopupOpen {
    $processes = Get-Process -Name "powershell" -ErrorAction SilentlyContinue
    foreach ($process in $processes) {
        if ($process.MainWindowTitle -eq "Warning") {
            return $true
        }
    }
    return $false
}

# Main script
if (Check-ForErrorFiles -directory $directoryToCheck -and -not (Is-PopupOpen)) {
    Show-Popup
}
