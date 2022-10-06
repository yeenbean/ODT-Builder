@ECHO OFF

REM This script creates an archive that includes everything needed to run the
REM script, essentially creating a portable version.  This is useful when you
REM want to deploy to machines using the most up-to-date versions of Office, or
REM if you want to move this script and bring the dependencies with you.
REM
REM This is the same script I use when building the version we distribute
REM internally at work.
REM
REM The output file will always be .\distribution.zip.  This script will not
REM check if the previous file was deleted, and may fail if it is not moved or
REM deleted first!

.\7za a -r distribution.zip lib\ 7za.dll 7za.exe 7zxa.dll Build.ps1 setup.exe