[CmdletBinding()]
param (
    [Parameter()]
    [switch]
    $EnableDebugger
)

if ($EnableDebugger) {
    $DebugPreference = "Continue"
}

Write-Debug "Script launched"

# Main menu
function Show-Menu
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $menu
    )

    switch ($menu) {
        "main" {
            if (-Not $EnableDebugger) {Clear-Host}
            Write-Host "╔════════════════════════════════════════════╗"
            Write-Host "║ Microsoft 365 ODT Wizard                   ║"
            Write-Host "╟────────────────────────────────────────────╢"
            Write-Host "║ 1. Microsoft 365 Apps for business         ║"
            Write-Host "║ 2. Cleanup                                 ║"
            Write-Host "╚════════════════════════════════════════════╝"
            Write-Host
        }
        "bits" {
            Write-Host "Would you like to download a 32 bit or 64 bit package?"
            Write-Host "1: 32-bit"
            Write-Host "2: 64-bit"
            Write-Host
        }
        "install" {
            Write-Host "Is this the target machine?"
            Write-Host "1: Yes"
            Write-Host "2: No (create a portable archive)"
            Write-Host
        }
        Default {}
    }
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


# Build package
function New-Build
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $bits,
        [Parameter()]
        [switch]
        $install
    )

    # Init variables
    $arch = "x86"
    if ($install) {
        Write-Debug "Install switch was passed."
    }

    # Validate bits
    switch ($bits) {
        "32" {
            $arch = "x86"
            Write-Debug "32-bit build"
        }
        "64" {
            $arch = "x64"
            Write-Debug "64-bit build"
        }
        default {
            Throw "Invalid architecture."
        }
    }

    # Download packages using ODT.
    Write-Host "Downloading packages. Please wait..."
    .\setup.exe /download .\lib\business-$arch.xml

    # Install or compile
    if ($install) {
        Write-Debug "Installation step reached."
        .\setup.exe /configure .\lib\business-$arch.xml
    }
    else {
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
        Write-Debug "Build step reached."
        Write-Host "Creating build archive..."
        Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build\office$arch.zip    #remove previous build if it exists
        .\7za a -r build\office$arch.zip build$bits
    }
    Remove-Cache
}


# Main process
:top while ($true)
{
    Show-Menu -menu main
    $choice = Read-Host "> "

    switch ($choice) {
        '1' {
            # 365 Apps for business
            Show-Menu -menu bits
            switch (Read-Host "> ") {
                '1' {
                    # 32 bit 365 Apps for business
                    Show-Menu -menu install
                    switch (Read-Host "> ") {
                        '1' {
                            # Install 32 bit 365 Apps for business
                            New-Build -bits 32 -install
                        }
                        '2' {
                            # Create portable archive of 32 bit 365 Apps for business
                            New-Build -bits 32
                        }
                        Default {continue}
                    }
                }
                '2' {
                    # 64 bit 365 Apps for business
                    Show-Menu -menu install
                    switch (Read-Host "> ") {
                        '1' {
                            # Install 64 bit 365 Apps for business
                            New-Build -bits 64 -install
                        }
                        '2' {
                            # Create portable archive of 64 bit 365 Apps for business
                            New-Build -bits 64
                        }
                        Default {}
                    }
                }
                Default {continue}
            }
        }
        '2' {
            Remove-Cache
        }
        Default {}
    }
}
