This is an excel file with a VBA program embedded, that calculates tidal 
windows.

==========
Changelog
==========
*****
Issue 08:
Bugfix in finalize procedure. If ship has RTA, but no ETA, finalize failed, but no warning given. Added 'finalization failed' message, and fixed RTA date input for statistics.
Source:
Bug (noted by client)
Additional:
Updated data to those in use by client
*****
Issue 07:
Add ranges in the VAARPLAN sheet that hold ATA's on several points.
Source:
Request by client
*****
Issue 06:
Visualize which vessels are ingoing and which are outgoing, by filling in the cell left of the Ships' name to "Opvarend schip:" (ingoing) or "Afvarend schip:" (outgoint, and set different colors for both.
Source:
Request by client
Additional:
Added 'client data transfer' routine, to streamline the data transfer process

*****
Issue 05:
Visualize that the vessel is over 340 meters in lenght, by coloring the text in the LOA cell and Ships name cell red.
Source:
Request by client

*****
Issue 04:
Lock the 'VAARPLAN' worksheet, and unlock only the cells which requires input. That way, the user can 'tab' their way through the input cells, improving the usability. No password will be set for the worksheet.
Source:
Request by client
Additional:
Discarded IMO number as input parameter (also in Finalize form).

*****
Issue 03:
Add checkbox for 'vessel underway' (reis onderweg) on the 'VAARPLAN' sheet. Checking this box will change the color of the ship's name and IMO number to visualize.
Source:
Request by client
Additional:
-updated routes and waypoints to those in use by client.

*****
Issue 02:
Adapt the 'finaliseer reis' routine to save the workbook when finished.
Source:
Request by client
Additional:
-updated routes and waypoints to those in use by client.
-changed a header in the 'Waypoints' sheet from 'Waterdiepte' to 'Drempeldiepte'


*****
Issue 01:
Adapt the dropdown menu in the worksheet to show entries sorted alphabetically.
Source:
Request by client
Additional:
updated routes and waypoints to those in use by client
