# sersol-localholdings
A semi-simple script that will come through MARC record files and creates a CSV file for eventual upload into Serials Solutions to update local holdings. 

Our current workflow to update our local holdings in Serials Solutions requires multiple review files, MarcEdit sessions, and forbidden Excel magic to create a spreadsheet based on the template provided by Serials Solutions. This script is an attempt to lessen the pain of this workflow.

The MARC record files I have used for this script are .out dump files from III's Data Exchange (each file contains ~50K records).

# Shameless forking (attribution)
This script was shamelessly forked from https://gist.github.com/mmccollow/348178. 

# Dependencies
- Python 2.7.3
- pymarc
- Some MARC record files :)

While the original script (see attribution section above) was developed in an *nix environment, the forked script is optimized for Windows 7. Sorry about that.

# What does this script do?
The script will do the following:
- Go through a folder (location specified in script) and search for all files ending in an .out extension
- Loop through each file in the folder and analyze each record to see if it should or should not be excluded from the spreadsheet
- If the record is not excluded, add select metadata (modified if needed) to a CSV file (delimited by an asterisk *) structured according to the SerSol local holdings spreadsheet template

Once the script is finished, the CSV file will need to be opened in Excel to parse out the fields (delimited by an asterisk *).

# TODO
- Moar testing
- Better README documentation
- Generalize internal documentation
- Create php/html GUI for staff to run the script server-side
