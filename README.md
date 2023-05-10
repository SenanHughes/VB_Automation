# VB_Automation

To reduce the time spent on a Report Out task repeated weekly in a previous role, I created this VB script to carry out the work. The file has been adjusted to remove any details that shouldn't be displayed.

The automated task was as follows:
  Create a copy of the previous weeks File and update the name to reflect the current date
  Import data from a .txt file into an empty workbook
  Transfer that information into the appropriate sheets in the CurrentDateFile
  Iterate through the two imported sheets and check if the data there was missing any pertenent details that would need manual updating
  If no manual updates required it iterated through to check if the Release sheet (which would be viewed by those on the distribution list when shared) did not align with the newly imported data
  Any mismatches were then updated and highlighted purple for clarity
  A copy of the Release sheet was made in a separate workbook which was filtered against the purple cells
  The main workbook his all but the Release sheet and removed the purple filter
  
  The sheets were then ready for distribution.
