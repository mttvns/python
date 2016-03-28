# python
python scripts

The scripts in the python repository are short utility scripts.

extract_ssx_zip_files.py: This script extracts all zip files in a directory to a new folder in that same directory named the same as the zip file. SSX refers to the original file extension that was compressed into a zip file so that its contents could
then be extracted. The original ssx files were saved as [filename].ssx.zip. .ssx files are proprietary files created by a company that manufactures CPR manikins that collect data about the CPR session for the purpose of training. The data is downloadable from the manikin in the form of a .ssx file. The files can be opened by converting them to zip files, which can then be extracted.

mass_macro_walker_git.py: This script loops through all Excel files in the same folder, runs macros in each file, and saves each file with a new name. The user is prompted to enter the path to the directory with the files in it, and a prefix to add to the file name to indicate that it has been processed by the script.

ssx_to_zip.py: converts files with extension .ssx to a file with extension .zip; can of course be modified to convert any file type to a zip file. .ssx files are proprietary files created by a company that manufactures CPR manikins that collect data about the CPR session for the purpose of training. The data is downloadable from the manikin in the form of a .ssx file. The files can be opened by converting them to zip files, which can then be extracted.
