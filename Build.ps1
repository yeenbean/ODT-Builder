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
            Write-Host "╔════════════════════════════════════════════════════════════════╗"
            Write-Host "║ Microsoft 365 ODT Wizard                                       ║"
            Write-Host "╟────────────────────────────────────────────────────────────────╢"
            Write-Host "║ 1. Microsoft 365 Apps for business (business standard/premium) ║"
            Write-Host "║ 2. Microsoft 365 Apps for enterprise (E3/E5)                   ║"
            Write-Host "║ 3. Microsoft 365 Home Premium                                  ║"
            Write-Host "╟────────────────────────────────────────────────────────────────╢"
            Write-Host "║ 4. Outlook (business standard/premium)                         ║"
            Write-Host "║ 5. Outlook (E3/E5)                                             ║"
            Write-Host "╟────────────────────────────────────────────────────────────────╢"
            Write-Host "║ 6. Custom (place in lib/custom.xml)                            ║"
            Write-Host "║ 5. Cleanup                                                     ║"
            Write-Host "╚════════════════════════════════════════════════════════════════╝"
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
    Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build
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
        [string]
        $version,
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

    # Ask to clear cache if previous build detected
    

    # Download packages using ODT.
    Write-Host "Downloading packages. Please wait..."
    .\setup.exe /download .\lib\$version-$arch.xml

    # Remove and recreate build directory
    Write-Host "Creating build directory..."
    Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build\$version-$arch
    New-Item -Name "build\$version-$arch" -ItemType "directory"

    # Copy over library files
    Write-Host "Copying helper files..."
    Copy-Item ".\setup.exe" -Destination ".\build\$version-$arch\setup.exe"
    Copy-Item ".\lib\$version-$arch.xml" -Destination ".\build\$version-$arch\$version-$arch.xml"
    #Copy-Item ".\lib\$version-$arch.bat" -Destination ".\build\$version-$arch\install.bat"
    $batContent = ".\setup.exe /configure .\$version-$arch.xml"
    $batFile = ".\build\$version-$arch\install.bat"
    [IO.File]::WriteAllLines($batFile, $batContent)

    # Move over downloaded packages
    Write-Host "Moving downloaded files from ODT..."
    Move-Item ".\Office" -Destination ".\build\$version-$arch\Office"

    # Install or compile
    if ($install) {
        Write-Debug "Installation step reached."
        Start-Process -FilePath ".\setup.exe" -WorkingDirectory ".\build\$version-$arch" -ArgumentList "/configure .\$version-$arch.xml"
    }
    else {
        # Compile archive
        Write-Debug "Build step reached."
        Write-Host "Creating build archive..."
        Remove-Item -Recurse -Force -ErrorAction 'silentlycontinue' .\build\office$arch.zip    #remove previous build if it exists
        .\7za a -r build\$version-$arch.zip build\$version-$arch
    }
}


# Main process
:top while ($true)
{
    # Declare variables.
    $version = "not selected"
    $bits = 0
    $install = $false

    # Main menu.
    Write-Debug "Top menu."
    Show-Menu -menu main
    switch (Read-Host "> ") {
        '1' {
            $version = "business"
        }
        '2' {
            $version = "enterprise"
        }
        '3' {
            $version = "home"
        }
        '4' {
            $version = "outlookbusiness"
        }
        '5' {
            $version = "outlookenterprise"
        }
        '6' {
            Write-Host "Feature not yet implemented."
            Write-Host "Press enter to continue..."
            Read-Host ""
            Continue top
        }
        '7' {
            Remove-Cache
            Continue top
        }
        Default {Continue top}
    }

    # Select architecture
    Write-Debug "Arch menu"
    Show-Menu -menu bits
    switch (Read-Host "> ") {
        '1' {
            $bits = 32
        }
        '2' {
            $bits = 64
        }
        Default {Continue top}
    }

    # Select whether to build or install
    Show-Menu -menu install
    switch (Read-Host) {
        '1' {
            $install = $true
        }
        '2' {
            $install = $false
        }
        Default {Continue top}
    }

    # Run build tool
    if ($install) {New-Build -version $version -bits $bits -install}
    else {New-Build -version $version -bits $bits}
}
