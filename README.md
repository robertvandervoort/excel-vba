# This repository holds functions and macros I have needed to create for working with data in Excel.

## stripGUIDorUIDfromURI.vba
Used for stripping GUID or UID from the end of a URI. Why would you want to do this? If you have a table with URIs in it and are wanting to understand how many unique URIs are getting hit, you can't do this because of the GUID, UID or other identifier that is present in the URI. This script has a commented out section where you can directly specify matches to strip as well that might not fit into GUID or UID format (standard formats). Once you create your "clean" column including this function you can then use that column in pivot tables or other functions. 
use: =stripGUIDorUIDfromURI(cell) replace cell with the cell in the row where you want to create a the clean URI from. Then you can apply to the column going all the way down your spreadsheet, or however else you'd normally use these.
