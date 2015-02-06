# SaveAsNewCentral
Opens Revit Central files and saves them to a new location.

This tool can help with migrations of Revit Central files from one server to another, either before or after the migration.  Central files are stored with a UNC path such that local files will always look to the same location during Sync With Central.  If those central files are moved to a new server, the UNC path will change, and local files will look to the wrong location.  While this is clearly apparent with an error when you open the file manually, doing this with many files can be extremely time consuming.

## Moving file prior to using them in the new location

Create an Excel file (using the included template if you like) that lists the full Source path + filename in one column and the full Destination path + filename in another column.  The tool will open the Source file, detached, with all worksets closed, handles a couple common open error dialogs, such as missing links, then saves the file in the Destination.  It will NOT overwrite a file that already exists in the Destination location.  If the Destination folder does not exist, it will be created.

1.  Start the "Save As New Central" command from the Add-ins ribbon.  You do not need to be in a document to start the command.
2.  Browse to the Excel file.  Make sure the file is not open in Excel.
3.  Select the worksheet to read.  The default is the first sheet.
4.  Select the Source and Destination columns.  The read starts on row 2 to allow for headers on row 1.
5.  The default save option is to "Save to New Location."
6.  Click Start.

## Updating files that have already been moved.

Create an Excel file (using the included template if you like) that lists the full Source path + filename in one column.  The tool will rename the Source file with a suffix you specify, delete the "_backup" folder if it exists, open the renamed Source file, detached, with all worksets closed, handles a couple common open error dialogs, such as missing links, then saves the file as the original Source name.

1.  Start the "Save As New Central" command from the Add-ins ribbon.  You do not need to be in a document to start the command.
2.  Browse to the Excel file.  Make sure the file is not open in Excel.
3.  Select the worksheet to read.  The default is the first sheet.
4.  Select the Source column.  The read starts on row 2 to allow for headers on row 1.
5.  The default save option is to "Save to New Location.", so change it to "Overwrite in Existing Location."
6.  Specify the suffix for the backup filename.  You cannot use "_backup" as the suffix to avoid name conflicts with the backup folder.  If a Source file is "Central1.rvt" and the suffix is "_migration", the backup filename will be "Central1_migration.rvt."
7.  Click Start.

## Future ideas

* Also copy linked files in the same folder.
* Change paths to linked files that also move.

## Open Source

Obviously, this is open source.  Please help out and contribute by submitting Issues, or by forking and submitting a pull request.

### Disclaimer

This is Sample code, and not an Autodesk product.
AUTODESK DOES NOT GUARANTEE THAT YOU WILL BE ABLE TO SUCCESSFULLY DOWNLOAD OR IMPLEMENT ANY SAMPLE CODE.
SAMPLE CODE IS SUBJECT TO CHANGE WITHOUT NOTICE TO YOU. AUTODESK PROVIDES SAMPLE CODE "AS IS" WITHOUT
WARRANTY OF ANY KIND, WHETHER EXPRESS OR IMPLIED, INCLUDING WARRANTIES OF MERCHANTABILITY AND FITNESS
FOR A PARTICULAR PURPOSE. IN NO EVENT SHALL AUTODESK OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER
INCLUDING DIRECT, INDIRECT, INCIDENTAL, CONSEQUENTIAL, LOSS OF DATA, OR LOSS OF BUSINESS PROFITS OR
SPECIAL DAMAGES, THAT MAY OCCUR AS A RESULT OF IMPLEMENTING OR USING ANY SAMPLE CODE, EVEN IF AUTODESK
OR ITS SUPPLIERS HAVE BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. 
