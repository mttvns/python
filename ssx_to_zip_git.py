import os
import shutil

'''
Python Script: Mass Macro Runner
By: Matthew Gary Evans, Copyright 2016 All Rights Reserved
Date: March 27, 2016

This script converts files with extension .ssx to a file with extension
.zip.

.ssx files are proprietary files created by a company that manufactures
CPR manikins that collect data about the CPR session for the purpose of
training. The data is downloadable from the manikin in the form of a .ssx file.
The files can be opened by converting them to zip files, which can then
be extracted.

To use, go to the folder with the ssx file(s) and copy the path to the folder.
Then run this script and paste the folder name into the command line or
IDLE shell.

The script will convert the files from ssx to zip.

This script can of course be modified to convert any file type to a zip file.
'''

print( 'Enter the path to the folder with the .ssx files:' )
folder = input() + '\\'
for folderName, subfolders, filenames in os.walk( folder ):
        for filename in filenames:
            if filename.endswith( '.ssx' ):
                os.rename( folder + filename, folder + filename + ".zip" )
                
