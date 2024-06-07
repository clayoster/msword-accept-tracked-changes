# msword-accept-tracked-changes
Powershell script to accept tracked changes on Word documents and convert them to PDFs

By default this script will accept tracked changes on any documents within the user's "Documents\msword-accept-tracked-changes" directory and convert them to PDFs. The output files will end up in a folder named "export" beneath the source directory.

The directory containing the Word documents which need processing can be overridden by using the "-source" option

Example: `./msword-accept-tracked-changes.ps1 -source C:\Path\To\Word-Files`
