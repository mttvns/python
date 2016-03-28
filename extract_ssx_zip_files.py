import os, zipfile, ntpath

'''
Python Script: Mass Macro Runner
By: Matthew Gary Evans, Copyright 2016, All Rights Reserved
Date: March 27, 2016

ABOUT: This script extracts all zip files in a directory to a new folder in that
same directory named the same as the zip file. SSX refers to the original file
extension that was compressed into a zip file so that its contents could
then be extracted. The original ssx files were saved as [filename].ssx.zip.
'''

print( 'Enter the path to the folder with the zip files to be extracted: ')
top = input()

for folderName, subfolders, filenames in os.walk( top ):
    i = 0
    for filename in filenames:
        #find files that end in .ssx.zip
        ssx = os.path.splitext( filename )[0]
        root = os.path.splitext( ssx )[0]
        if filename.endswith( '.zip' ):
            os.chdir( top )
            src = zipfile.ZipFile( filename )
            #save the file with a unique number to avoid duplicates
            src.extractall( top + '\\' + root + '_' + str(i) )
            print( 'New folder made: ' + top + '\\' + root + '_' + str(i) )
            print( '' )
            src.close()
            i += 1
