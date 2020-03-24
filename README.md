# VBA_Comment_Macro

Hi there.

I created this macro because I had to track a set of items over time.

This macro will digest a report and save any new comments pertaining to items already on the master report.
These new comments will be placed in a newly created column labeled with the days date.
The new comments will appear adjacent to the old comments for ease of use.

__It assumes the following:__

  -That the item needing to be tracked has a stable unique identifier named "Reference ID" (without the quotes).

  -That all relevant information pertaining to the item is contained on a single row of the reference excel spreadsheet/report.

  -That the folder the persistent, master workbook lives in contains a folder named "Holding_Bin".

  -That "Holding_Bin" contains a folder named "Archived".

  -That the only thing in Holding_Bin, other than Archived, is the new workbook to be mined for information (which will be appended to the master workbook if the Reference ID already appears in the master workbook)

  -The above points are described in greater detail--with an example--in the code's comments.

  -That the first row of each workbook, rather naturally, contains the column headers (e.g., "Reference ID").

__Here's how it works:__

  -Make sure you've got your master workbook in your directory of choice (formatted as explained above).

  -Drag and drop the new report into Holding_Bin.

  -Run the Macro from the master Workbook

  -Enjoy the time-savings and productivity gain of not needing to parse irrelevant information in the new report!
