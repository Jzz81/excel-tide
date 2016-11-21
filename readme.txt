This is an excel file with a VBA program embedded, that calculates tidal 
windows.

The program is embedded in the .\TideWin_excel\overzicht_reizen.xlsm file

standard databases (with sample data) are stored in the .\data folder

The program depends on the use of sqlite libraries, also included in the .\data folder.

All of the program gui is written in Dutch language. Code (and most of the database definitions) is written in English.

To use the program, open the TideWin_excel\overzicht_reizen.xlsm file, click on 'Programma instellingen' in the 'Vaarplannen' ribbon.

Now fill in the following paths:
".\data\tidal_data\YearTide_sample-2016.accdb"
".\data\tidal_data\YearTide_sample-2016_HW.accdb"
".\data\databases\TideWin_excel_active_db.accdb"
".\data\databases\sail_plan_archive.accdb"
".\data\SqliteLibs\"
"" (leave empty)

==========
Changelog
==========
Version 4 will host a central program and data base, opposite to the program and data base for each sail plan of the previous versions.
*database will load from the access database each time the program opens and keep the sqlite database in memory as long as the program is open
*All input is to be done using userforms (with the exemption of some data alterations that are permitted from the sail plan list)
*All output (overview list and calculated data presentation) is done in one Excel file (that holds the code base as well)
*All data storage is done in Access database (treshold data, route data, connection data, ship data, sail plan data, etc)

