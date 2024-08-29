function Generate-Folders {
    param(
        [string[]]$folders,
        [string]$desktop
    )

    foreach ($folder in $folders) {
        $path = Join-Path -Path $desktop -ChildPath $folder

        if (-not (Test-Path -path $path)) {
            New-Item -path $path -ItemType Directory
            Write-Host "Created folder: $path"
        }
    }
}

function Get-Files {
    param(
        [string]$path,
        [string]$type
    )

    return Get-ChildItem $path -Filter $type
}

function Extract-And-Sort-Archives {
    param(
        [string]$outputPath,
        [System.IO.FileInfo[]]$zipFiles,
        [string]$cdPath,
        [string]$dvdPath,
        [int]$isoSize,
        [string]$winRarPath
    )

    $jobs = @()

    foreach ($zipFile in $zipFiles) {
        $jobs += Start-Job -ScriptBlock {
            param($zipFile, $outputPath, $cdPath, $dvdPath, $isoSize, $winRarPath)

            try {
                # Extract the archive
                $arguments = "x -y `"$($zipFile.FullName)`" `"$outputPath`""
                $winRarPath = "C:\Program Files\WinRAR\WinRAR.exe"
                
                Start-Process -FilePath $winRarPath -ArgumentList $arguments -Wait -NoNewWindow

                # Remove the archive file
                Remove-Item -Path $zipFile.FullName -Force
                Write-Host "Successfully extracted and deleted $($zipFile.Name)"

                # Sort the extracted files
                $extractedFiles = Get-ChildItem -Path $outputPath -Filter "*.iso"
                foreach ($file in $extractedFiles) {
                    if ($file.Length -gt ($isoSize * 1KB)) {
                        $destination = Join-Path -Path $dvdPath -ChildPath $file.Name
                        Write-Host "Moving $($file.Name) to DVD folder"
                    }
                    else {
                        $destination = Join-Path -Path $cdPath -ChildPath $file.Name
                        Write-Host "Moving $($file.Name) to CD folder"
                    }

                    Move-Item -Path $file.FullName -Destination $destination
                    Write-Host "Moved $($file.Name) to $destination"
                }

            }
            catch {
                Write-Host "Failed to process $($zipFile.Name): $($_.Exception.Message)"
            }

        } -ArgumentList $zipFile, $outputPath, $cdPath, $dvdPath, $isoSize, $winRarPath
    }

    # Wait for all jobs to complete
    $jobs | ForEach-Object {
        $_ | Wait-Job | Out-Null
        Remove-Job $_
    }
}

# Variables
$folders = @("MX4SIO\CD", "MX4SIO\DVD", "MX4SIO\Zip")
$desktop = [environment]::GetFolderPath("Desktop")

$cdPath = Join-Path -Path $desktop -ChildPath $folders[0]
$dvdPath = Join-Path -Path $desktop -ChildPath $folders[1]
$zipFolderPath = Join-Path -Path $desktop -ChildPath $folders[2]

$isoSize = 700  # Size in KB
# Update this with the correct path

# 1. Create directories if they do not exist
Write-Host "Starting process..."
Generate-Folders -folders $folders -desktop $desktop

# 2 & 3: Loop every 20 seconds
while ($true) {
    Write-Host "Checking for archives..."

    # Get all archive files in the Zip folder
    $zipFiles = Get-Files -path $zipFolderPath -type "*.7z"
    $hasArchives = $zipFiles.Count -gt 0

    if ($hasArchives) {
        Write-Host "Extracting archives..."
        Extract-And-Sort-Archives -outputPath $zipFolderPath -zipFiles $zipFiles -cdPath $cdPath -dvdPath $dvdPath -isoSize $isoSize -winRarPath $winRarPath
    }
    else {
        Write-Host "No zip files found."
    }

    Write-Host "Waiting for 20 seconds before the next check..."
    Start-Sleep -Seconds 20
}

Write-Host "Process completed."
