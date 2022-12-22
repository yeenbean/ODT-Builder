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
2. Run `Build.ps1` or `run.bat`.
3. Choose a version of Office from the main menu.
4. Choose your architecture.
5. Choose whether to build an archive, or to deploy to the current machine.

The script will then download the required files, then either install Office or
create a zip archive which you can move to the target machine.

### Installing from an archive

1. Transfer the archive to the target machine.
2. Extract the archive into it's own directory.
3. Run `install.bat`. You'll see the familiar installation UI.
4. ?????
5. Profit.

## TODO

- [x] Move support files to a subdirectory to clean up root.
- [x] Modify build directory names.
- [ ] Add additional versions of Office, including Apps for enterprise, Office
LTSC versions, and other volume licensed versions.
- [x] Support custom configurations. This will help future-proof this script if
it goes unmaintained, as well as supporting custom-made configs (i.e. from
https://config.office.com/deploymentsettings)
- [x] Rewrite in PowerShell.

## Changelog

All versions are tested before they are published, including installation of
Office in a virtual environment.

### 2022.12.22

- Re-imagined build wizard.
- Office builds are now created in their own directories instead of a catch-all
build directory.
- Added several new ODT config files: Home retail, Enterprise, and options for
just Outlook.
- Rudimentary implementation for custom ODT config files.
- New "run" and "debug" batch scripts to make launching/testing the PowerShell
script easier.
- Code cleanup, bug squashing, and more. Check out the full changelog below.

### 2022.10.06

- `create-release.bat` never made it into August's commit. It is now uploaded.
- Created `run.bat` script which automatically bypasses execution policy. This
makes the script easier to execute on modern Windows operating systems.

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
