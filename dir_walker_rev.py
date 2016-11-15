import os, shutil, ntpath
'''
This script searches all folders and files in a directory for a file
named CPREvents.xml, and copies that file to the root directory.
'''

print( 'Directory name: ')
root = input()
#rootLevel establishes the number of folders in the foldername
rootLevel = root.count( os.sep )
#counter for errors generated when attempting to copy files
copyErrors = 0
for foldername, subfolders, filenames in os.walk( root ):

    #subLevel is the number of folders deep the folder/file is in below mainSub
    subLevel = foldername.count( os.sep )
    
    if subLevel - rootLevel == 1:
        #newNameBase is the first part of the new name for the xml file to be moved
        newNameBase = root + '\\' + ntpath.split( foldername )[1] + '_'
        print( 'newNameBase = ' + newNameBase )

    #split the name of the subfolder up and see if it contains the
    #substring 'bls'; if so, assign the foldername to be newNameBls;
    #newNameBls is the second part of the new file name for the xml file
    #to be moved
    namePart = ntpath.split( foldername )[1]
    if namePart.find( 'bls' ) != -1:
        newNameBls = namePart + '_'
        print( 'newNameBls = ' + newNameBls )
            
    for subfolder in subfolders:
        do = "nothing"
    
    for filename in filenames:
        #find the xml file and move it to the root directory name the user input
        #at the beginning of the script, and rename the xml file to be
        #newNameBase_newNameBls_CPREvents.xml
        if filename == 'CPREvents.xml':
            oldPath = foldername + '\\' + filename
            newPath = newNameBase + newNameBls + filename
            
            if shutil.copy( oldPath, newPath ):
                print( 'CPREvents copied to ' + newPath )
                print( '' )
            else:
                print( '!!!!!!!!!!!!!!!!!!!!')
                print( '!!!!!!!!!!!!!!!!!!!!')
                print( 'Failed to copy CPREvents file at ' + oldPath )
                copyErrors += 1

print( '' )
print( '' )
print( 'Done! Number of errors = ' + str(copyErrors) )
