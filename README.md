# ODT Builder

ODT Builder creates offline installers for Microsoft Office 365 Business apps.
This is useful when installing Office in an environment with slow/no Internet,
or when the standard installers are not functioning as expected.

## Usage

### First-time setup

If you are downloading the source or a release from GitHub, you will need
to follow these setup instructions before using the software.

1. Download the [latest release](https://github.com/yeenbean/ODT-Builder/releases)
of ODT Builder and extract the archive.
2. Download the standalone version of 7-zip from
[here](https://www.7-zip.org/a/7z2107-extra.7z). Extract the three files from
the "x64" directory to the same folder as Build.ps1.
3. Download and install
[Microsoft's ODT tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117).
Run the installer and install to the same directory as Build.ps1.

### Creating offline installers

1. Extract ODT Builder into its own directory.
2. Run `Build.ps1`. *You may have to change your PowerShell settings to allow
execution of the script.*
3. Choose one of the options from the main menu to start a build. You can build
both versions simultaneously, or just one version.
4. Let the build complete. You may not see any output during the download
process.
5. Deploy your installation to your target machine. See instructions below.

### Installing from offline installer

1. Copy or upload the version of Office needed. 32 and 64 identify 32-bit and
64-bit versions.
2. Extract the archive into it's own directory.
3. Run `install.bat`. You'll see the familiar installation UI.
4. ?????
5. Profit.

## TODO

- [x] Move support files to a subdirectory to clean up root.
- [ ] Modify build directory names.
- [ ] Add alternative versions of Office, including Apps for enterprise, Office
LTSC versions, and other volume licensed versions.
- [ ] Support custom configurations. This will help future-proof this script if
it goes unmaintained, as well as supporting custom-made configs (i.e. from
https://config.office.com/deploymentsettings)
- [ ] Switches for automatically accepting the EULA, silent installation,
embedding a license key, disabling automatic updates, update channel, etc.
- [x] Rewrite in PowerShell.

## Changelog

All versions are tested before they are published, including installation of
Office in a virtual environment.

### 2022.08.15

- Updated readme TODO list.
- Added a `create-release.bat` script to create portable versions.
- Updated `.gitignore` to avoid weirdness with extracted files from ODT.

### 2022.01.28

- Fixed window border.
- Now published on GitHub!
- Rewrote usage instructions.
- Gitignore generated. Does not include binaries that should not be distributed.

### 2022.01.15

- Rewritten in PowerShell. Script is now more robust and written "correctly".
- Menu system. Now you can build just 32-bit or 64-bit, as well as both!
- Manually run cleanup -- Really only useful if the script crashes mid-build.
- Compiled builds are now identified by architecture shorthand instead of
bit-ness.

### 2022.01.13

- Switched from `tar` to `7zip` for compression. This removes the requirement
for a specific Windows 10 build, and 7zip can make zip files that work natively
in Windows Explorer.
- Builds now output to `.\build\` instead of in the main directory.
- Updated file cleanup process to delete x64 build directory.
- Fixed 64-bit build. Now uses correct XML in bundle.
- Timeout reduced to 1 second.
- Provided insight that there's no progress indicator for download.

### 2022.01.12

First iteration! We'll see how this goes. :)
