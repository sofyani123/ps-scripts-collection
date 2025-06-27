# PowerShell script to extract attachments from .msg files, excluding .png files, with verification
# Source and extracted location (same directory)
$sourcePath = "S:\Business Change and Infrastructure\IT\Reports"
$extractedPath = "S:\Business Change and Infrastructure\IT\Reports"
$markerFile = Join-Path -Path $extractedPath -ChildPath "ExtractionComplete.txt"
$verificationReport = Join-Path -Path $extractedPath -ChildPath "VerificationReport.txt"

# Check if extraction has already been completed
if (Test-Path -Path $markerFile) {
    Write-Warning "Extraction already completed. Marker file found at $markerFile. Delete the marker file to re-run."
    exit 0
}

# Verify the source/extracted directory exists
if (-not (Test-Path -Path $extractedPath)) {
    Write-Error "Directory $extractedPath does not exist or is inaccessible."
    exit 1
}

# Initialize Outlook COM object
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Error "Failed to initialize Outlook COM object. Ensure Microsoft Outlook is installed."
    exit 1
}

# Get all .msg files in the source directory
$msgFiles = Get-ChildItem -Path $sourcePath -Filter "*.msg" -File

# Check if any .msg files were found
if ($msgFiles.Count -eq 0) {
    Write-Warning "No .msg files found in $sourcePath"
    $outlook.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    exit 0
}

# Initialize verification log
$verificationLog = @()
$verificationLog += "Attachment Extraction Verification Report - $(Get-Date)"
$verificationLog += "------------------------------------------------"

# Loop through each .msg file exactly once for extraction
foreach ($file in $msgFiles) {
    Write-Host "Processing file: $($file.FullName)"
    try {
        # Open the .msg file
        $msg = $outlook.Session.OpenSharedItem($file.FullName)
        
        # Check if the message has attachments
        if ($msg.Attachments.Count -eq 0) {
            Write-Host "No attachments found in $($file.Name)"
            $msg.Close(1) # olDiscard
            $verificationLog += "File: $($file.Name) - No attachments to extract"
            continue
        }

        # Track attachments for verification
        $expectedAttachments = @()

        # Process each attachment
        foreach ($attachment in $msg.Attachments) {
            $attachmentName = $attachment.FileName

            # Skip .png files (case-insensitive)
            if ($attachmentName -match '\.png$' -or $attachmentName -match '\.PNG$') {
                Write-Host "Skipping .png attachment: $attachmentName"
                continue
            }

            $expectedAttachments += $attachmentName
            $savePath = Join-Path -Path $extractedPath -ChildPath $attachmentName

            # Handle filename conflicts by appending a number
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($attachmentName)
            $extension = [System.IO.Path]::GetExtension($attachmentName)
            $counter = 1
            while (Test-Path -Path $savePath) {
                $newFileName = "${baseName}_${counter}${extension}"
                $savePath = Join-Path -Path $extractedPath -ChildPath $newFileName
                $counter++
            }

            # Save the attachment
            $attachment.SaveAsFile($savePath)
            Write-Host "Extracted attachment: $attachmentName to $savePath"
        }

        # Close the message
        $msg.Close(1) # olDiscard

        # Add to verification log
        if ($expectedAttachments.Count -eq 0) {
            $verificationLog += "File: $($file.Name) - No non-.png attachments to extract"
        } else {
            $verificationLog += "File: $($file.Name) - Extracted $($expectedAttachments.Count) non-.png attachments"
        }
    } catch {
        Write-Error "Error processing file $($file.Name): $_"
        $verificationLog += "File: $($file.Name) - Error during extraction: $_"
        $msg.Close(1) # olDiscard (if $msg is defined)
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        exit 1
    }
}

# Verification step: Check if attachments were extracted
Write-Host "Verifying extracted attachments..."
$verificationLog += "------------------------------------------------"
$verificationLog += "Verification Results"

foreach ($file in $msgFiles) {
    try {
        # Open the .msg file again for verification
        $msg = $outlook.Session.OpenSharedItem($file.FullName)
        $nonPngAttachments = @($msg.Attachments | Where-Object { $_.FileName -notmatch '\.png$' -and $_.FileName -notmatch '\.PNG$' })
        
        if ($nonPngAttachments.Count -eq 0) {
            $verificationLog += "File: $($file.Name) - No non-.png attachments expected (Verified)"
            $msg.Close(1) # olDiscard
            continue
        }

        $missingAttachments = @()
        foreach ($attachment in $nonPngAttachments) {
            $attachmentName = $attachment.FileName
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($attachmentName)
            $extension = [System.IO.Path]::GetExtension($attachmentName)
            
            # Check for original or numbered variants (e.g., file.pdf, file_1.pdf)
            $found = $false
            $counter = 0
            do {
                $checkName = if ($counter -eq 0) { $attachmentName } else { "${baseName}_${counter}${extension}" }
                $checkPath = Join-Path -Path $extractedPath -ChildPath $checkName
                if (Test-Path -Path $checkPath) {
                    $found = $true
                    break
                }
                $counter++
            } while ($counter -lt 100) # Reasonable limit to prevent infinite loops

            if (-not $found) {
                $missingAttachments += $attachmentName
            }
        }

        if ($missingAttachments.Count -eq 0) {
            $verificationLog += "File: $($file.Name) - All $($nonPngAttachments.Count) non-.png attachments found (Verified)"
        } else {
            $verificationLog += "File: $($file.Name) - Missing attachments: $($missingAttachments -join ', ')"
        }

        $msg.Close(1) # olDiscard
    } catch {
        $verificationLog += "File: $($file.Name) - Error during verification: $_"
    }
}

# Save verification report
$verificationLog | Out-File -FilePath $verificationReport

# Create marker file to indicate completion
"Extraction completed on $(Get-Date)" | Out-File -FilePath $markerFile

# Clean up Outlook COM object
$outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
Write-Host "Attachment extraction and verification completed."
Write-Host "Verification report saved to $verificationReport"
Write-Host "Marker file created at $markerFile"