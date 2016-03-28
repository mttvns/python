import os, win32com.client
'''
Python Script: Mass Macro Runner
By: Matthew Gary Evans, http://www.mttvns.com
Date: March 27, 2016

ABOUT: This script loops through all Excel files in the same folder, runs macros in each file, and saves each file with a new name. The user is prompted to enter the path to the
directory with the files in it, and a prefix to add to the file name to indicate that it has been processed by the script.
'''

#prompt user for the path to the directory with the Excel files
print( 'Enter the path to the directory with the Excel files in it (e.g., C:\\path\\to\\your\\folder):' )
root = input()

#prompt the user to provide a prefix to be added to the files when saved after the script is done running
print( 'Enter the prefix you wish to give each file name once it has been processed (e.g., if you choose a prefix of myPrefix_ a file named myFile.xls will be saved as myPrefix_myFile.xls; if you don\'t enter anything, the file will be saved as \'COPY_OF_myFile.xls\'):')
done_prefix = input()

#assign default value to done_prefix if user didn't enter anything, otherwise use user's input
if done_prefix == '':
        done_prefix = 'COPY_OF_'
else:
        done_prefix = done_prefix

#dispatch Excel; add a workbook with a module with a macro to delete
#the macro once done running in each file; this is necessary to avoid an
#ambiguous name error because otherwise this script will recreate the other
#macros with the same name in the same workbook every time it loops
xl = win32com.client.Dispatch( "Excel.Application" )
wb = xl.Workbooks.Add()
xlmodule = wb.VBProject.VBComponents.Add(1)
code = '''
Sub DeleteModule()
    ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents("Module2")   
End Sub'''
xlmodule.CodeModule.AddFromString( code )

#loop through all Excel files in the directory and run the macros
#the macros could be anything
for foldername, subfolders, filenames in os.walk( root ):
        for filename in filenames:
                if not filename.startswith( done_prefix ) and not os.path.exists( root + "\\" + done_prefix + filename):
                        currentFile = root + "\\" + filename
                        print( 'Completed: ' + filename )
                        xlmodule2 = wb.VBProject.VBComponents.Add(1)
                        code = '''
Sub Sub_1()

Application.ScreenUpdating = False
' your macro here; turn off screen updating so it will work

End Sub
'''
                        xlmodule2.CodeModule.AddFromString( code )
                        xl.Application.Run( "Sub_1" )
                        code = '''
Sub Sub_2()

' your code here
        
End Sub

Sub Sub_3()


End Sub

'''
                        xlmodule2.CodeModule.AddFromString( code )
                        macro = wb.Name + "!Sub_2"
                        xl.Application.Run( macro )
                        #save the file, close the workbook, and then run the macro that deletes the macros above( Subs_1, 2 and 3 in this demo )
                        xl.ActiveWorkbook.SaveAs( root + "\\" + done_prefix + filename )
                        xl.ActiveWorkbook.Close()
                        xl.Application.Run( "DeleteModule" )
                        
print( "Work complete!" )
print( "Press enter to exit.")

#this is necessary to keep the CLI from automatically closing once the script is done
finish = input()
