# Main menu
function Show-OdtMenu
{
    Clear-Host
    Write-Host "╔════════════════════════════════════════════╗"
    Write-Host "║ Offline Installer Creation Wizard          ║"
    Write-Host "╟────────────────────────────────────────────╢"
    Write-Host "║ 1. Create 32-bit Apps for business         ║"
    Write-Host "║ 2. Create 64-bit Apps for business         ║"
    Write-Host "║ 3. Create 32- and 64-bit Apps for business ║"
    Write-Host "║ 4. Cleanup                                 ║"
    Write-Host "╚════════════════════════════════════════════╝"
    Write-Host
}


# Cleanup function. This deletes all unfinalized and downloaded files from ODT.
# Cleanup does *not* delete supportive files, like ODT itself or 7zip, or any
# successful builds.
function Remove-Cache
{
    Write-Host "Cleaning up..."
    Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build32
    Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build64
    Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\Office
}


# Build 32-bit Apps for business package.
function New-Build
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $bits
    )

    # Init variables
    $arch = "x86"

    # Validate bits
    switch ($bits) {
        "32" {
            $arch = "x86"
        }
        "64" {
            $arch = "x64"
        }
        default {
            Throw "Invalid architecture."
        }
    }

    # Download packages using ODT.
    Write-Host "Downloading packages. Please wait..."
    .\setup.exe /download .\lib\business-$arch.xml

    # Remove and recreate build directory
    Write-Host "Creating build directory..."
    Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build$bits
    New-Item -Name "build$bits" -ItemType "directory"

    # Copy over library files
    Write-Host "Copying helper files..."
    Copy-Item ".\setup.exe" -Destination ".\build$bits\setup.exe"
    Copy-Item ".\lib\business-$arch.xml" -Destination ".\build$bits\business-$arch.xml"
    Copy-Item ".\lib\install$bits.bat" -Destination ".\build$bits\install.bat"

    # Move over downloaded packages
    Write-Host "Moving downloaded files from ODT..."
    Move-Item ".\Office" -Destination ".\build$bits\Office"

    # Compile archive
    Write-Host "Creating build archive..."
    Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build\office$arch.zip    #remove previous build if it exists
    .\7za a -r build\office$arch.zip build$bits

    # Cleanup
    Remove-Cache
}


# Main process
while ($true)
{
    Show-OdtMenu
    $choice = Read-Host "Pick a number...any number"
    Write-Host

    switch ($choice) {
        '1' {
            New-Build -bits "32"
        }
        '2' {
            New-Build -bits "64"
        }
        '3' {
            New-Build -bits "32"
            New-Build -bits "64"
        }
        '4' {
            Remove-Cache
        }
        Default {}
    }
}
