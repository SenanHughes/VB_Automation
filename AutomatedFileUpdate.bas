Attribute VB_Name = "AutomateGeneric"
Option Explicit
Public NewDateYY As String                      '# Variable to assign todays date in the necessary format for the file update
Public NewDateYYYY As String                    '# Variable to assign todays date in the necessary format for the folder update
Public LastDate As String                       '# Variable to calculate the date for a week prior to day running the script
Public LastDateYY As String                     '# Reformat LastDate variable to fit the file name like NewDateYY
Public LastDateYYYY As String                   '# Reformat LastDate variable to fit the folder name like NewDateYYYY
Public SaveLocale As String                     '# Variable to store the folder name, broken up for clearer reading in code
Public LastSaveLocale As String                 '# Variable that stores the name of the folder from a week prior
Public NewSaveLocale As String                  '# Variable that stores the name for the new folder for this week
Public LastCRScoping As String                  '# Variable to store the name of the previous CRScoping Doc
Public NewCRScoping As String                   '# Variable to store the name of the new CRScoping Doc
Public LastCRChanges As String                  '# Variable to store the name of the old CRScoping Changes Doc
Public NewCRChanges As String                   '# Variable to store the name of the new CRScoping Changes Doc
Public CRWorkbook As Workbook                   '# Variable to create a new blank workbook for importing .txt files into
Public LastRow As Integer                       '# Variable used to track the last row of various worksheets
Public LastRowXXX As Integer                    '# Variable used to track the last row of various worksheets
Public LastRowYYY As Integer                    '# Variable used to track the last row of various worksheets
Public LastRowFormula As Integer                '# Variable when needing to track the last row of a different section of the same worksheet as LastRow
Public Directory As String                      '# Variable to store the drive "MAINDRIVE" is assigned to
Public CRScopingWB As Workbook                  '# To reactive the CRScopingWB
Public XXXError As Boolean                      '# If issues are found in XXX sheet this value is triggered
Public YYYError As Boolean                      '# Same as above for YYY
Public DisplaySheet As Worksheet                '# Creates a worksheet variable for ease of reading and calling
Public Sheet As Integer                         '# A variable used for improving script writing, by assigning a variable of 1 and 2 to this sheet functions can be repeated without having to be written twice
Public SheetName As String                      '# This stores the sheetname for the above mentioned efficiency
Public LastRowName As Integer                   '# This stores the final row of the relevant sheet for the above efficiency
Public RDIIter As Integer                       '# Used to iterate through the ReleaseDesignIntent sheet

Public Function UpdateCRScopingDoc()            '# The main function to Import the Data for a new week of CR and PR's
Dim FSO As New FileSystemObject                 '# Creating a new FileSystemObject allows the modification of folders/directories with VBA script
Dim oWMI As WbemScripting.SWbemServices
Dim oCols As WbemScripting.SWbemObjectSet
Dim oCol As WbemScripting.SWbemObject
Dim sWQL As String
Dim sDriveLetter As String
Dim DriveFound As Boolean
Dim ichr As Integer
Dim Letter As String
Dim i As Integer

'### https://www.devhut.net/determine-the-path-of-a-mapped-drive/ Function referenced from here
If Dir("AFILELOCATION") = "" Then
    DriveFound = False
    ichr = Asc("A")
    For i = 1 To 26
        If ichr > 90 Then ichr = 64 + ichr - 90
            Letter = Chr(ichr)
        sDriveLetter = Letter & ":"
        Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
        sWQL = "SELECT ProviderName FROM Win32_MappedLogicalDisk WHERE Name = '" & sDriveLetter & "'"
        Set oCols = oWMI.ExecQuery(sWQL, , wbemFlagReturnImmediately Or wbemFlagForwardOnly)
        
        For Each oCol In oCols
            Directory = oCol.ProviderName
            If Directory = "ROOTFOLDER" Then
                DriveFound = True
                Exit For
            End If
        Next oCol
        If DriveFound = True Then
            Exit For
        End If
        ichr = ichr + 1
    Next i
    Directory = sDriveLetter
Else
    Directory = "Z:"                             '# Assigning the relevant drive for MAINDRIVE change to match your drive if necessary
End If
NewDateYY = Format(Now, "YY-MM-DD")             '# Assigning and formatting today's date in the way necessary for upcoming file naming convention
NewDateYYYY = Format(Now, "YYYY-MM-DD")         '# Assigning and formatting today's date in the way necessary for upcoming folder naming convention
LastDate = DateAdd("d", -7, Date)               '# Calculating the date for a week previous to today
LastDateYY = Format(LastDate, "YY-MM-DD")       '# Same as NewDateYY above but for last week
LastDateYYYY = Format(LastDate, "YYYY-MM-DD")   '# Same as above but for NewDateYYYY
SaveLocale = Directory & "NEWFILELOCATION"    '# Assigning the top level folder for which this work will carry out in
LastSaveLocale = SaveLocale & LastDateYYYY      '# Assigning the name for the folder for last week, merging SaveLocale with the relevant formatted date
NewSaveLocale = SaveLocale & NewDateYYYY       '# Assigning the name for the new folder for this week, same as above

FSO.CopyFolder LastSaveLocale, NewSaveLocale    '# Using the above FSO we've created we can now copy the previous weeks folder in the new locale with the new name

LastCRScoping = NewSaveLocale & "\DOCUMENTNAME" & LastDateYY & ".xlsx"    '# For readability in the renaming function and ease of update,
                                                                                                            '# this assigns the last weeks CRScoping file to a variable. Merging the necessary components
NewCRScoping = NewSaveLocale & "\DOCUMENTNAME" & NewDateYY & ".xlsx"      '# Same as above but for the new week
Name LastCRScoping As NewCRScoping              '# Using the CRScoping variables we can rename the file to suit this week
Kill NewSaveLocale & "\DOCUMENTNAME" & LastDateYY & ".xlsx"
NewCRChanges = NewSaveLocale & "\DOCUMENTNAMEWITHCHANGES" & NewDateYY & ".xlsx"

MsgBox ("Please copy and paste the TCE data into the newly created .txt Files and delete up as far as | Level |, hit OK when done.")

Application.ScreenUpdating = False  '# Disabling Screen Updating so as to stop the Excel windows opening until the end of the automation to reduce strain on the computer
Application.EnableEvents = False    '# Disabling Events so as to reduce number of pop-ups and strain on the computer
CreateCRWorkbook                    '# Calling the function to create the workbook for importing .txt files
ImportDelimitedText                 '# Calling the function for Importing the .txt files to the created workbook
OpenAndEditCRScoping                '# Calling the function to open the new CRScoping file and edit it
UpdateFormulae                      '# Calling the function to update the relevant formulae
UpdateRDIFunc                       '# Calls function that automates the updating of the main sheet for sharing with customers
CheckForMissingClassOrPart          '# Calls function that will check for any Class or Parts data missing from YYY and XXX sheets
FreezePanesFunc                     '# Calls function to reset Release Design Intent sheet for updating
Application.EnableEvents = True     '# Re-Enabling events now that the function has finished
Application.ScreenUpdating = True   '# Re-Enabling Screen Updating now that the function has finished
End Function

Public Function CreateCRWorkbook()

ChDrive Directory                   '# Making sure the correct drive is being run from
ChDir NewSaveLocale                 '# Changing the directory to work in the newly copied and named folder
Set CRWorkbook = Workbooks.Add      '# Adds in the new blank workbook to be used for data import
CRWorkbook.Worksheets.Add After:=CRWorkbook.Worksheets(1), Count:=1 '# Adds an additional sheet to cover both XXX and YYY imports
End Function

Public Function ImportDelimitedText()
Dim NoOfRows As Integer         '# A variable used to count the number of rows once data has been imported
Dim NameSheet As Worksheet      '# Assigning a name to a worksheet to make referring to it more straightforward
Dim TxtImport As String

For Sheet = 1 To 2          '# Given the functions are the same for YYY and XXX sheet, an efficiency can be created by using the below If loop
    If Sheet = 1 Then       '# Starting at 1, this sets the SheetName variable to match that of the XXX Sheet and the LastRowName variable to be the last row of the XXX sheet
        SheetName = "AAA XXX Conversion Sheet"
        LastRowName = LastRowXXX
        TxtImport = "\FILE.txt"
        Set NameSheet = CRWorkbook.Sheets("Sheet1")
    Else                    '# The second loop sets Sheet to 2 and the SheetName and LastRowName to that of the YYY sheet allowing one writing of the function instead of doubling up
        SheetName = "AAA YYY Conversion Sheet"
        LastRowName = LastRowYYY
        TxtImport = "\FILE2.txt"
        Set NameSheet = CRWorkbook.Sheets("Sheet2")
    End If
    NameSheet.Name = SheetName           '# Renaming the first sheet to the name in the Main Workbook for ease of later callbacks
    ImportTxtFiles NewSaveLocale & TxtImport, 1, NameSheet.Range("A1")      '# Calling the ImportTxtFiles function with the variables to import the related .txt data
    NameSheet.Columns(1).EntireColumn.Delete             '# Deleting a blank column that isn't needed
    NoOfRows = NameSheet.UsedRange.Rows.Count            '# Getting the XXX sheet row count now that data is imported and reformatted
    NameSheet.Range("$A$1", "$K$" & NoOfRows).AutoFilter Field:=10, Criteria1:="LC State", Operator:=xlOr, Criteria2:="="    '# Filtering to only show the blank and "LC State" cells within column J...
    NameSheet.Range("$A$1:$K$" & NoOfRows).Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete                     '# This is done to be able to delete them and give one continous block of data
    If NameSheet.FilterMode = True Then
        NameSheet.ShowAllData            '# To allow the upcoming copy and paste section the data needs to be made visible again
    End If
Next Sheet
End Function

Public Function ImportTxtFiles(ImportedFile, FirstRow, ImportingTo)    '# This function imports the data delimited from .TXT to the new Workbook
With ImportingTo.Parent.QueryTables.Add(Connection:="TEXT;" & ImportedFile, Destination:=ImportingTo) '# Creates a new Table, connection := TEXT notes the source is a text file from the input destination
    .TextFileStartRow = FirstRow        '# This notes where in the .txt file to start importing data from
    .TextFileParseType = xlDelimited    '# States that the data being imported is to be delimited
    .TextFileOtherDelimiter = "|"       '# States the method of delimiting
    .Refresh BackgroundQuery:=False     '# Once all the data has been run through the querytable will communicate with the data source to import the data
                                        '# It waits until the end due to it being setting to false
End With
End Function

Public Function OpenAndEditCRScoping()
Dim WS As Worksheet                             '# Local variable to allow iterating through each WS in the the Workbook

Application.DisplayAlerts = False               '# Disabling Alerts allows the opening of the book and updating the links without input from the user
Workbooks.Open NewCRScoping, UpdateLinks:=1     '# This opens the New CRScoping doc and updates and links automatically
Application.DisplayAlerts = True                '# Re-enabling alerts when the workbook is open
Set CRScopingWB = ActiveWorkbook                '# Assign the CurrentWorkbook to a variable name to store it for future function calls
For Each WS In Worksheets                       '# To iterate through each individual sheet in the workbook and unhide them all
    WS.Visible = xlSheetVisible
Next WS
LastRowXXX = CRScopingWB.Sheets("AAA XXX Conversion Sheet").Range("J3").End(xlDown).Row             '# Determines the LastRow of the XXX sheet from the newly imported data
LastRowYYY = CRScopingWB.Sheets("AAA YYY Conversion Sheet").Range("J3").End(xlDown).Row             '# Same as above but for YYY
For Sheet = 1 To 2          '# Given the functions are the same for YYY and XXX sheet, an efficiency can be created by using the below If loop
    If Sheet = 1 Then       '# Starting at 1, this sets the SheetName variable to match that of the XXX Sheet and the LastRowName variable to be the last row of the XXX sheet
        SheetName = "AAA XXX Conversion Sheet"
        LastRowName = LastRowXXX
    Else                    '# The second loop sets Sheet to 2 and the SheetName and LastRowName to that of the YYY sheet allowing one writing of the function instead of doubling up
        SheetName = "AAA YYY Conversion Sheet"
        LastRowName = LastRowYYY
    End If
    CRScopingWB.Sheets(SheetName).Cells(1, 1).Value = Format(Now, "DD-MM-YYYY")     '# Updates the date to this week in Cell A1 on the XXX sheet
    CRScopingWB.Sheets(SheetName).Range("B3", CRScopingWB.Sheets(SheetName).Range("K3").End(xlDown)).ClearContents  '# Removes all data from columns B to K on the current sheet
    If CRScopingWB.Sheets(SheetName).FilterMode = True Then
        CRScopingWB.Sheets(SheetName).ShowAllData
    End If
    CRWorkbook.Sheets(SheetName).Range("$A$2", CRWorkbook.Sheets(SheetName).Range("J2").End(xlDown)).Copy       '# Copies the newly imported data from the new Workbook for the current sheet
    CRScopingWB.Sheets(SheetName).Range("B3").PasteSpecial Paste:=xlPasteValues     '# Pastes the data using paste special - as values into the appropriate cells in the workbook
    CRScopingWB.Sheets(SheetName).Activate
    CRScopingWB.Sheets(SheetName).Range("A1").Select                       '# Sets the active cell to deselect the group selection
    If CRScopingWB.Sheets(SheetName).FilterMode = True Then                '# Checks to see if there are any filters on the current sheet
        CRScopingWB.Sheets(SheetName).ShowAllData                          '# If there are filters it removes them to show the full sheet data
    End If
Next Sheet
Application.CutCopyMode = False         '# Cancels the copy operation to allow closing without clipboard pop up
CRWorkbook.Close SaveChanges:=False     '# Closes the new workbook without saving as data isn't needed anymore
End Function

Public Function UpdateFormulae()
Dim i As Integer            '# Local integer to allow the increase of a counter in a for loop
Dim StartRow As Integer     '# Local variable to determine start row of function for CR List (Class) sheet
Dim NewRange As Range      '# Local variable to be able to store a Range Address for a filldown operation later
Dim RowCount As Integer     '# Local variable to update a RowCount in the CR List (Class) For loop
Dim PreviousRow As Integer  '# Same as Above
Dim LinkedBookExists As String
Dim LinkedBook As String
Dim fd As Office.FileDialog
Dim NewLinkedBook As String
Dim LinkedFileExtension As String
Dim LinkedFileName As String
Dim IndexVal As Integer

LastRowXXX = CRScopingWB.Sheets("AAA XXX Conversion Sheet").Range("J3").End(xlDown).Row             '# Determines the LastRow of the XXX sheet from the newly imported data
LastRowYYY = CRScopingWB.Sheets("AAA YYY Conversion Sheet").Range("J3").End(xlDown).Row             '# Same as above but for YYY
For Sheet = 1 To 2          '# Given the functions are the same for YYY and XXX sheet, an efficiency can be created by using the below If loop
    If Sheet = 1 Then       '# Starting at 1, this sets the SheetName variable to match that of the XXX Sheet and the LastRowName variable to be the last row of the XXX sheet
        SheetName = "AAA XXX Conversion Sheet"
        LastRowName = LastRowXXX
    Else                    '# The second loop sets Sheet to 2 and the SheetName and LastRowName to that of the YYY sheet allowing one writing of the function instead of doubling up
        SheetName = "AAA YYY Conversion Sheet"
        LastRowName = LastRowYYY
    End If
    LastRowFormula = CRScopingWB.Sheets(SheetName).Range("M3").End(xlDown).Row      '# Determines the LastRow of the Formula section of the current sheet
    If LastRowName > LastRowFormula Then        '# Compares the length of the formula section and the imported data section
                                                '# If the LastRow is bigger it means the formula section needs to be extended
        For i = 13 To 32                    '# Iterating through columns M to AF which is where the formulae need updating
            Set NewRange = Range(Cells(3, i), Cells(LastRowName, i))                '# Determines the Range that covers the whole column down as far as the imported data
            CRScopingWB.Sheets(SheetName).Range(NewRange.Address).FillDown      '# Copies the formula down for the iterated Column until it matches the imported data length
            Next i          '# Iterates to the next column
    Else            '# If the Formula Section is greater than the the imported section then the excess formulae rows need removing
        LastRowName = LastRowName + 1   '# Moving down one row beyond the imported data section
        CRScopingWB.Sheets(SheetName).Range("$M$" & LastRowName, "$AF$" & LastRowFormula + 1).Clear     '# These cells are now cleared of their data
    End If
Next Sheet
'# Here the CR List (Class) sheet is updated, it's simply a case of dragging the formula down until no new data is imported
StartRow = CRScopingWB.Sheets("CR List (Class)").Range("D2").End(xlDown).Row         '# Given the first section of the sheet is manual inputs and subject to change, the start of the formula section needs to be variable
LastRow = CRScopingWB.Sheets("CR List (Class)").Range("D" & StartRow).End(xlDown).Row    '# Using the StartRow and checking down the end of the formula section is determined

LinkedBook = ActiveWorkbook.LinkSources(xlExcelLinks)(1)
LinkedBookExists = Dir(LinkedBook)
LinkedFileExtension = "RELEVANTFILELOCATION"
If LinkedBookExists = "" Then
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    MsgBox "The FILE has had the file name changed, please select the latest version in the next dialogue box.", Title:="Update FILE Link"
    With fd
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx?", 1
        .Title = "Choose an Excel file"
        .AllowMultiSelect = False
        .InitialFileName = LinkedFileExtension
        If .Show = True Then
            NewLinkedBook = .SelectedItems(1)
        End If
    End With
    MsgBox "The Links are updating, this may take a few minutes. Please be patient", Title:="Links About to Update"
    LinkedFileName = Split(NewLinkedBook, "\")(5)
    CRScopingWB.Sheets("CR List (Class)").Cells(StartRow, 1).Formula = "=INDEX('" & LinkedFileExtension & "[" & LinkedFileName & "]FILE'!$A:$C,$D54,1)"
    CRScopingWB.Sheets("CR List (Class)").Cells(StartRow, 3).Formula = "=INDEX('" & LinkedFileExtension & "[" & LinkedFileName & "]FILE'!$A:$C,$D54,3)"
    Set NewRange = Range(Cells(StartRow, 1), Cells(LastRow, 1))
    CRScopingWB.Sheets("CR List (Class)").Range(NewRange.Address).FillDown
    Set NewRange = Range(Cells(StartRow, 3).Address, Cells(LastRow, 3).Address)
    CRScopingWB.Sheets("CR List (Class)").Range(NewRange.Address).FillDown
End If
RowCount = LastRow
While CRScopingWB.Sheets("CR List (Class)").Cells(RowCount, 1).Value <> "00:00:00"       '# Until no new data (Column A Row X = 00:00:00) this loop will repeat
    PreviousRow = RowCount              '# To be able to iterate a non formulaic cell the previous row needs to be recorded before advancing
    RowCount = RowCount + 1             '# Advancing the row count to fill the formula down
    Set NewRange = Range(Cells(PreviousRow, 1), Cells(RowCount, 1))             '# For all the other columns the address for a new range of the previous row to the new row is needed
    CRScopingWB.Sheets("CR List (Class)").Cells(RowCount, 4).Value = CRScopingWB.Sheets("CR List (Class)").Cells(PreviousRow, 4).Value + 1
    CRScopingWB.Sheets("CR List (Class)").Range(NewRange.Address).FillDown                  '# The formulae are then copied down to the new range
    If Not IsError(Application.Match(CRScopingWB.Sheets("CR List (Class)").Cells(RowCount, 1), Range(CRScopingWB.Sheets("CR List (Class)").Cells(1, 1), CRScopingWB.Sheets("CR List (Class)").Cells(StartRow - 1, 1)), 0)) Then
        IndexVal = Application.Match(CRScopingWB.Sheets("CR List (Class)").Cells(RowCount, 1), Range(CRScopingWB.Sheets("CR List (Class)").Cells(1, 1), CRScopingWB.Sheets("CR List (Class)").Cells(StartRow - 1, 1)), 0)
        CRScopingWB.Sheets("CR List (Class)").Cells(IndexVal, 1).EntireRow.ClearContents
    End If
Wend
If CRScopingWB.Sheets("CR List (Class)").Cells(RowCount, 1).Value = "00:00:00" Then         '# Once the loop is broken and there is no new data the incorrect data needs to be deleted
    CRScopingWB.Sheets("CR List (Class)").Range("$A$" & RowCount, "$E$" & RowCount).Clear
    RowCount = RowCount - 1
End If
Set NewRange = Range(Cells(LastRow, 2).Address, Cells(RowCount, 3).Address)
Set NewRange = Union(NewRange, Range(Cells(LastRow, 5).Address, Cells(RowCount, 5).Address))
CRScopingWB.Sheets("CR List (Class)").Range(NewRange.Address).FillDown
End Function

Public Function PartClassMessage(XXXError, YYYError, XXXClassError, YYYClassError, XXXPartError, YYYPartError)
Dim MessageOut  As String       '# An arbitrary variable to facilitate the multi-condition message box

If XXXError = True And YYYError = True Then     '# In the event that there is an error on both the YYY and XXX sheets the following will be called
    If (XXXClassError = True And YYYPartError = True) Or (YYYClassError = True And XXXPartError = True) Then        '# This statement is displayed if there are both YYY AND XXX Part AND Class errors
        MessageOut = MsgBox("The XXX and YYY sheets are missing both Parts and Class details, the ones needing updating have been filtered.", Title:="Update Details")
    ElseIf XXXClassError = True And YYYClassError = True And YYYPartError = False And XXXPartError = False Then     '# This statement is called if there are only XXX AND YYY Class Errors
        MessageOut = MsgBox("The XXX and YYY sheets are missing Class details, the ones needing updating have been filtered.", Title:="Update Details")
    ElseIf XXXClassError = False And YYYClassError = False And YYYPartError = True And XXXPartError = True Then     '# This statement is called if there are only XXX AND YYY Part Errors
        MessageOut = MsgBox("The XXX and YYY sheets are missing Part details, the ones needing updating have been filtered.", Title:="Update Details")
    End If
ElseIf XXXError = True And YYYError = False Then        '# In the event that there are only XXX Errors one of the following will be called
    If XXXClassError = True And XXXPartError = True Then        '# Each of these statements following the same conditions as above only this time ignoring YYY Errors
        MessageOut = MsgBox("The XXX sheet is missing both Parts and Class details, the ones needing updating have been filtered.", Title:="Update Details")
    ElseIf XXXClassError = True And XXXPartError = False Then
        MessageOut = MsgBox("The XXX sheet is missing Class details, the ones needing updating have been filtered.", Title:="Update Details")
    ElseIf XXXClassError = False And XXXPartError = True Then
        MessageOut = MsgBox("The XXX sheet is missing Part details, the ones needing updating have been filtered.", Title:="Update Details")
    End If
ElseIf XXXError = False And YYYError = True Then        '# In the event that there are only YYY Errors one of the following will be called
    If YYYClassError = True And YYYPartError = True Then        '# Each of these statements following the same conditions as above only this time ignoring XXX Errors
        MessageOut = MsgBox("The YYY sheet is missing both Parts and Class details, the ones needing updating have been filtered.", Title:="Update Details")
    ElseIf YYYClassError = True And YYYPartError = False Then
        MessageOut = MsgBox("The YYY sheet is missing Class details, the ones needing updating have been filtered.", Title:="Update Details")
    ElseIf YYYClassError = False And YYYPartError = True Then
        MessageOut = MsgBox("The YYY sheet is missing Part details, the ones needing updating have been filtered.", Title:="Update Details")
    End If
End If
End Function

Public Function UpdateRDIFunc()
Dim RDIEnd As Integer                           '# Variable used to store the final row of the Release Design Intent Sheet
Dim Iter As Integer                             '# Itering variable used to progress through the arrays to compare their values
Dim XXXYYYCellIter As Range                     '# A range itering variable to allow flicking through each row in the XXX/YYY sheets to check whether they've been hidden by an autofilter
Dim NewValues As Range                          '# Creates a variable for storing the non-hidden cells that need to be compared against the old date
Dim OldValues As Range                          '# Variable that stores the old data on the Release Design Intent sheet for comparison to the new
Dim NewRange As Range                           '# Variable for a range that will be used to filldown data when a new line is added to the Release Design Intent Sheet
Dim NewArray(1 To 2, 1 To 5) As Variant         '# An array used to store the values and addresses from the NewValues range because it's easier to use for the non-continuguous nature of the range
Dim OldArray(1 To 2, 1 To 5) As Variant         '# An array used to store the values and addresses from the OldValues range because it's easier to use for the non-continuguous nature of the range
Dim CellIter As Range                           '# Variable to be used to iterate through each individual cell in the New/OldValues range
Dim Count As Integer                            '# An iterating variable used to flick through values in the Arrays
Dim NameError As Boolean                        '# Captures YYY/XXX Error value in a local for use in this function which prevents writing out the same code twice
Dim itering As Integer
Dim RangeEnd As Integer
Dim AddLine As Boolean
Dim SubRDIIter As Integer


Set DisplaySheet = CRScopingWB.Sheets("2.x.x Release Design Intent")        '# Assigns the desired sheet to the variable
DisplaySheet.Activate                                                       '# Activates the Release Design Intent sheet as this is the sheet wanting to be updated
If DisplaySheet.FilterMode = True Then                                      '# Checks to see the Filters from the previous week are applied
    ActiveSheet.ShowAllData                                                 '# If they are applied, it removes them
End If
For Sheet = 1 To 2          '# Given the functions are the same for YYY and XXX sheet, an efficiency can be created by using the below If loop
    If Sheet = 1 Then       '# Starting at 1, this sets the SheetName variable to match that of the XXX Sheet and the LastRowName variable to be the last row of the XXX sheet
        SheetName = "AAA XXX Conversion Sheet"
        LastRowName = LastRowXXX
        NameError = XXXError
    Else                    '# The second loop sets Sheet to 2 and the SheetName and LastRowName to that of the YYY sheet allowing one writing of the function instead of doubling up
        SheetName = "AAA YYY Conversion Sheet"
        LastRowName = LastRowYYY
        NameError = YYYError
    End If
    If NameError = False Then
        CRScopingWB.Sheets(SheetName).Range("$A$3", "$AF$" & LastRowName).AutoFilter Field:=26, Criteria1:="#N/A"       '# Filters the sheet against any "N/A" values in column z
    End If
    RDIEnd = CRScopingWB.Sheets("2.x.x Release Design Intent").Range("D3").End(xlDown).Row                                                          '# Finds the final row of the Release Design Intent sheet and assigns it to a variable
    For Each XXXYYYCellIter In CRScopingWB.Sheets(SheetName).Range("$A$3", "$A$" & LastRowName)                                                     '# Loops through the current sheet to look at each individual row
         If Not XXXYYYCellIter.EntireRow.Hidden And Not IsError(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 27).Value) Then             '# Then checks to see if the row is hidden or not by the autofilter above and whether the value actually exists in the Release Design Intent sheet (column AA value not N/A)
            For RDIIter = 2 To RDIEnd                                                                                                               '# If it is visible then the system loops through the whole Release Design Intent sheet to see if there's a match, RDIIter relating to each row being checked
                If CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 11).Value = CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 14).Value And _
                    CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 14).Value = CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 17).Value And _
                         Not IsError(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 27).Value) Then                                                                  '# The first if condition looks for a match of the CR number and part number along with no error in the column that states whether the input exists or not
                        Set NewValues = CRScopingWB.Sheets(SheetName).Range(Cells(XXXYYYCellIter.Row, 14).Address)                                                            '# Assigns the first cell to the NewValue range, this range will be used for comparison to the old values on the Release Design Intent Sheet
                        Set NewValues = Union(NewValues, CRScopingWB.Sheets(SheetName).Range(Cells(XXXYYYCellIter.Row, 15).Address), _
                                            CRScopingWB.Sheets(SheetName).Range(Cells(XXXYYYCellIter.Row, 17).Address), _
                                                CRScopingWB.Sheets(SheetName).Range(Cells(XXXYYYCellIter.Row, 18).Address), _
                                                    CRScopingWB.Sheets(SheetName).Range(Cells(XXXYYYCellIter.Row, 20).Address))                                               '# Here the additional relevant cells are appended into the NewValues range for comparison
                        Set OldValues = CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells(RDIIter, 11).Address)                                                             '# The same is now done for the OldValues on the Release Design Intent sheet
                        Set OldValues = Union(OldValues, CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells(RDIIter, 12).Address), _
                                            CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells(RDIIter, 14).Address), _
                                                CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells(RDIIter, 15).Address), _
                                                    CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells(RDIIter, 17).Address))
                        Count = 1                                       '# To allow direct comparison of the above non-contiguous ranges an array is used to store the values and address, this count is used to iterate through the array
                        For Each CellIter In NewValues                  '# Iterating through each cell in the NewValues range to pull out the Address and Value to store in the array
                            NewArray(1, Count) = CellIter.Address
                            NewArray(2, Count) = CellIter.Value
                            Count = Count + 1                           '# Advancing the count to store the next cell in the next portion of the array
                        Next CellIter                                            '# Steps to the next cell in the range
                        Count = 1                                       '# Resets the count to go again for OldValues
                        For Each CellIter In OldValues                  '# This repeats the above for the OldValues range
                            OldArray(1, Count) = CellIter.Address
                            OldArray(2, Count) = CellIter.Value
                            Count = Count + 1
                        Next CellIter
                        For Iter = 1 To 5                               '# Using this for loop the array can now be iterated through to check if each of the 5 cells have any differences
                            If NewArray(2, Iter) <> OldArray(2, Iter) Then                                                                      '# If there are any differences between the New and Old values then it steps into this If statement
                                CRScopingWB.Sheets(SheetName).Range(NewArray(1, Iter)).Copy                                    '# The new value value is then copied from the XXX sheet
                                CRScopingWB.Sheets("2.x.x Release Design Intent").Range(OldArray(1, Iter)).PasteSpecial Paste:=xlPasteValues    '# It is then pasted into the Release Design Intent sheet
                                CRScopingWB.Sheets("2.x.x Release Design Intent").Range(OldArray(1, Iter)).Interior.ColorIndex = 29             '# To display the difference for sharing with relevant parties the cell colour is changed to purple
                                If Iter = 5 Then
                                    BaselineColumnFunc Left(NewArray(2, 3), 3), RDIIter, NewArray(2, Iter)             '# Calls this function to save it being written out multiple times, it checks what should be populated in the Baseline column based on the lifecycle state of the CR in question
                                End If
                            End If
                        Next Iter                                           '# Steps to the next position in the Array
                        Select Case CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 17)
                        Case "-Closed-"
                            Set OldValues = CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells(RDIIter, 11).Address, Cells(RDIIter, 19).Address)
                            OldValues.Font.ColorIndex = 15
                        Case Else
                        End Select
                        Exit For
                ElseIf ((Left(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 11).Value, 11) = Left(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 14).Value, 11) And _
                Right(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 11).Value, 1) <> Right(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 14).Value, 1)) Or _
                    (Left(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 11).Value, 10) = Left(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 14).Value, 10) And _
                        Right(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 11).Value, 2) <> Right(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 14).Value, 2))) And _
                            CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 14).Value = CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 17).Value Then                                    '# This is a slightly more specific if check, here looks to see if the part already exists but has been upsuffixed (aa-ab) or (ab-ba), otherwise it checks the same as the above
                         AddNewEntryFunc SheetName, XXXYYYCellIter, RDIEnd
                    Exit For
                End If
            Next RDIIter               '# Iterates to the next row of the Release Design Intent sheet if no
        ElseIf Not XXXYYYCellIter.EntireRow.Hidden And IsError(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 27).Value) Then                '# This final section checks whether the entry can be found at all in the sheet based on column AA's result, if N/A then an entire line will need to be input somewhere in the Release Design Intent Sheet
             For RDIIter = 2 To RDIEnd
                 If CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 15).Value = CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 18).Value Then     '### First look to match the part name itself, then if no match look for YYY! MAYBE GET RID OF AND ONLY DO MATCHING YYY NAME
                     For SubRDIIter = RDIIter To RDIEnd
                         If IsEmpty(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter, 2)) And _
                            CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 18).Value < CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter + 1, 6).Value Then
                             AddNewEntryFunc SheetName, XXXYYYCellIter, RDIEnd
                             Exit For
                         End If
                     Next SubRDIIter
                 Exit For
                 ElseIf Not IsError(CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 16).Value) And _
                    CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 4).Value = CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 16).Value Then '### Looking for YYY in Green to match the YYY in the CR request
                     For SubRDIIter = RDIIter To RDIEnd
                         If IsEmpty(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter, 2)) And _
                            CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 17).Value > CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter, 14).Value And _
                                (CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 17).Value < CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter + 1, 14).Value Or _
                                    Not IsEmpty(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter + 1, 2))) Then
                             RDIIter = SubRDIIter
                             AddNewEntryFunc SheetName, XXXYYYCellIter, RDIEnd
                             Exit For
                         End If
                     Next SubRDIIter
                 Exit For
                 ElseIf CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 13).Value = CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 16).Value Then '### Looking for YYY in Green to match the YYY in the CR request
                     For SubRDIIter = RDIIter To RDIEnd
                         If IsEmpty(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter, 2)) And _
                            CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 17).Value > CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter, 14).Value And _
                                (CRScopingWB.Sheets(SheetName).Cells(XXXYYYCellIter.Row, 17).Value < CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter + 1, 14).Value Or _
                                    Not IsEmpty(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(SubRDIIter + 1, 2))) Then
                             RDIIter = SubRDIIter
                             AddNewEntryFunc SheetName, XXXYYYCellIter, RDIEnd
                             Exit For
                         End If
                     Next SubRDIIter
                 Exit For
                 End If
             Next RDIIter
        End If
    Next XXXYYYCellIter       '# Once the conditions are met the next YYY/XXX row is stepped to
    If CRScopingWB.Sheets(SheetName).FilterMode = True Then         '# Checks to see if the filter is still active on the sheet, if it is it clears the filter to show everything again
        CRScopingWB.Sheets(SheetName).ShowAllData
    End If
Next Sheet            '# Steps onto the YYY sheet
Application.CutCopyMode = False
End Function

Public Function CheckForMissingClassOrPart()
Dim ClassColumn As Range            '# Creates variable for iterating through the class of each CR
Dim PartColumn As Range             '# Creates variable for iterating through the name of each Part
Dim Iter As Integer                 '# A variable that will be used to progress through each row of the worksheet
Dim ErrMess  As String              '# An error message will be output stating if any parts or classes need updating
Dim ErrorList As Range              '# A variable for collecting all cells in the XXX sheet that meet the error conditions
Dim ClassError As Boolean           '# If issues are found in XXX sheet Class column this value is triggered
Dim PartError As Boolean            '# If issues are found in XXX sheet Part column this value is triggered
Dim XXXClassError As Boolean        '# If issues are found in XXX sheet Class column this value is triggered
Dim XXXPartError As Boolean         '# If issues are found in XXX sheet Part column this value is triggered
Dim YYYClassError As Boolean        '# If issues are found in YYY sheet Class column this value is triggered
Dim YYYPartError As Boolean         '# If issues are found in YYY sheet Part column this value is triggered
Dim SheetError As Boolean

YYYError = False
XXXError = False
For Sheet = 1 To 2          '# Given the functions are the same for YYY and XXX sheet, an efficiency can be created by using the below If loop
ClassError = False
SheetError = False
PartError = False
    If Sheet = 1 Then       '# Starting at 1, this sets the SheetName variable to match that of the XXX Sheet and the LastRowName variable to be the last row of the XXX sheet
        SheetName = "AAA XXX Conversion Sheet"
        LastRowName = LastRowXXX
    Else                    '# The second loop sets Sheet to 2 and the SheetName and LastRowName to that of the YYY sheet allowing one writing of the function instead of doubling up
        SheetName = "AAA YYY Conversion Sheet"
        LastRowName = LastRowYYY
    End If
    Set ClassColumn = CRScopingWB.Sheets(SheetName).Range("$S$3:$S$" & LastRow)    '# Attributes the Class Column to the relevant column in the current Sheet
    Set PartColumn = CRScopingWB.Sheets(SheetName).Range("$P$3:$P$" & LastRow)     '# Attributes the Part Column to the relevant column in the current Sheet

    For Iter = 1 To LastRowName                     '# To iterate through each of the rows in the current Sheet
        If IsError(ClassColumn(Iter).Value) And Left(ClassColumn(Iter).Offset(0, -2).Value, 3) <> "405" Then    '# Checks to see if there is an error in the Class column, and then whether it relates to a PR (405).
            ClassError = True                                                           '# If it is an error and not related to a PR then the Class Error is triggered
            If SheetError = False Then                                                  '# If this is the first error found, SheetError will be false triggering this condition
                SheetError = True                                                       '# First the XXXError will now be set True to avoid repeating this step
                Set ErrorList = ClassColumn(Iter)                                       '# Then initialising the ErrorList range with the first value
            Else
                Set ErrorList = Union(ErrorList, ClassColumn(Iter))                     '# For every subsequent error it will be appended into the range here using the Union function
            End If
        ElseIf IsError(PartColumn(Iter)) Then                       '# This repeats the above but for the Parts column, first checking if there's an error in the row
            If PartColumn(Iter).Value = CVErr(xlErrNA) Then         '# Then confirming it is the desired error i.e "#N/A" which links to a missing part name as opposed to a Doc, Code etc which spits out a "#VALUE" error
                PartError = True                                    '# All remaining steps are repeats of the Class list conditions using the Part Column instead
                If SheetError = False Then
                    SheetError = True
                    Set ErrorList = PartColumn(Iter)
                Else
                    Set ErrorList = Union(ErrorList, PartColumn(Iter))
                End If
            End If
        End If
    Next Iter
    If ClassError = True Or PartError = True Then           '# Once the entire sheet has been iterated through and errors have been found
        CRScopingWB.Sheets(SheetName).Range("$A$3:$A$" & LastRowName).EntireRow.Hidden = True '# The entire sheets contents are hidden
        ErrorList.EntireRow.Hidden = False                  '# Then the range which needs to be fixed is unhidden to make fixing the issues more straightforward
    End If
    If Sheet = 1 Then                                       '# To store the Errors matching the XXX sheet the sheet number is checked and then the associated variables are updated
        XXXError = SheetError
        XXXClassError = ClassError
        XXXPartError = PartError
    Else                                                    '# Repeats the above if the sheet number is 2 to match the YYY variables
        YYYError = SheetError
        YYYClassError = ClassError
        YYYPartError = PartError
    End If
Next Sheet
If YYYError = True Or XXXError = True Then                  '# In the event of any errors this function is called which will display a message notifying the user what needs fixing
    ErrMess = PartClassMessage(XXXError, YYYError, XXXClassError, YYYClassError, XXXPartError, YYYPartError)
End If
End Function

Public Function AddNewEntryFunc(SheetName, XXXYYYCellIter, RDIEnd)
Dim itering As Integer
Dim RangeEnd As Integer
Dim NewRange As Range
Dim NewValues As Range                          '# Creates a variable for storing the non-hidden cells that need to be compared against the old date
Dim OldValues As Range                          '# Variable that stores the old data on the Release Design Intent sheet for comparison to the new


RDIIter = RDIIter + 1
CRScopingWB.Sheets("2.x.x Release Design Intent").Cells((RDIIter), 11).EntireRow.Insert         '# In the event of this being the case a new row is required to track the status of the part
CRScopingWB.Sheets("2.x.x Release Design Intent").Cells((RDIIter), 2).Select                    '# The start of this row is selected to then step through and check for blank cells until a non blank cell in column B is found. This means the end of this YYY section has been reached
For itering = 1 To RDIEnd                                                                       '# The stepping through is carried out with this operation, using the selected cell all the way to the end of the Release Design Intent Sheet
    ActiveCell.Offset(1, 0).Select                                                              '# The next row down is selected
    If Not IsEmpty(ActiveCell.Value) Then                                                       '# The cell is checked whether or not it is empty
        RangeEnd = ActiveCell.Row - 1                                                           '# If the cell was empty the end of the range is assigned as this row minus one, as the operation doesn't want to be brought down until the non empty row
        Exit For                                                                                '# The for loop and now be cancelled as the desired point has been met
    ElseIf itering + 1 = RDIIter Then
        RangeEnd = ActiveCell.Row + 1
        Exit For
    End If
Next itering
Set NewRange = CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells((RDIIter - 1), 1), Cells(RangeEnd, 4))                     '# The desired range is created for use in the following "FillDown" operation, columns A-D need to have the formula populated to allow the XXX/YYY sheets to recognise the updates have been made
NewRange.FillDown                                                                                                                       '# This function pulls down all the content from the first row of the range and iterates it through each row to the end of the range
Set NewValues = CRScopingWB.Sheets(SheetName).Range(Cells(XXXYYYCellIter.Row, 14).Address, (Cells(XXXYYYCellIter.Row, 23).Address))     '# An entire contiguous range can be selected for both NewValues and OldValues
Set OldValues = CRScopingWB.Sheets("2.x.x Release Design Intent").Range(Cells(RDIIter, 11).Address, (Cells(RDIIter, 19).Address))       '# this allows a direct copy and paste due to all of the content needing inputting
NewValues.Copy                                                                                                                          '# Similar operation as previous, only a specific cell doesn't need to be specified so it can just copy and paste to the whole range
Select Case CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 17)
    Case "-Closed-"
        OldValues.PasteSpecial Paste:=xlPasteValues
        OldValues.Interior.ColorIndex = 29
        OldValues.Font.ColorIndex = 15
    Case Else
        OldValues.PasteSpecial Paste:=xlPasteValues
        OldValues.Interior.ColorIndex = 29
    End Select
BaselineColumnFunc Left(CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 14).Value, 3), RDIIter, CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 17)           '# Calls this function to save it being written out multiple times, it checks what should be populated in the Baseline column based on the lifecycle state of the CR in question
End Function

Public Function BaselineColumnFunc(LeftCheck, RDIIter, LCStatus)                        '# Function to check what the BaselineColumn value will be, used repeatedly so reduces times it needs to be written out
If LeftCheck <> "405" Then                                                              '# The next check is whether or not the entry relates to a PR or a CR, if it doesn't begin with 405 it's a CR and meets this condition
    Select Case LCStatus                                                                '# The LC State is checked here and then for the appropriate condition it will fill in the required Baseline status and highlight the change as purple again
    Case "Authorized"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Value = "SR12.x Authorised CR"     '# Using the RDIIter variable the relevant row is found for the update to be updated
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
    Case "Actioned"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Value = "SR12.x Actioned CR"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
    Case "-Closed-"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10) = "-"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Font.ColorIndex = 15
    Case Else
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10) = "-"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
    End Select
Else                                                                                                            '# In the event that the CR number does begin with 405 it will be a PR and so has slightly different inputs into the Baseline column
    Select Case LCStatus
    Case "Authorized"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Value = "SR12.x Authorised PR"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
    Case "Actioned"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Value = "SR12.x Actioned PR"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
    Case "-Closed-"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Value = "N/A - PR"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Font.ColorIndex = 15
    Case Else
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Value = "-"
        CRScopingWB.Sheets("2.x.x Release Design Intent").Cells(RDIIter, 10).Interior.ColorIndex = 29
    End Select
End If
End Function

Public Function FreezePanesFunc()
Dim RDIEnd As Integer               '# Variable for storing final row of the RDI sheet
Dim ChangesBook As Workbook         '# Variable to store the creation of a new blank workbook for saving the changes made
Dim Sheet As Worksheet

RDIEnd = CRScopingWB.Sheets("2.x.x Release Design Intent").Range("D3").End(xlDown).Row      '# Assigns the final row of the RDI sheet to a variable
Set DisplaySheet = CRScopingWB.Sheets("2.x.x Release Design Intent")                        '# Assigns the desired sheet to the variable
If DisplaySheet.FilterMode = True Then                                                      '# Checks to see the Filters from the previous week are applied
    DisplaySheet.ShowAllData                                                                '# If they are applied, it removes them
End If
Set ChangesBook = Workbooks.Add                                                             '# Creates a new blank workbook
DisplaySheet.Copy Before:=ChangesBook.Sheets(1)                                             '# Copies across the RDI sheet to the new workbook
Application.DisplayAlerts = False                                                           '# Disables Excel from displaying alerts to allow deleting sheets without a pop-up
ChangesBook.Sheets(2).Delete                                                                '# Deletes the superfluous blank second sheet
Application.DisplayAlerts = True                                                            '# Re-enables the ability for Excel to Display Alerts
ChangesBook.Sheets("2.x.x Release Design Intent").Range("$A$1", "$AF$" & RDIEnd).AutoFilter Field:=17, Criteria1:=RGB(128, 0, 128), Operator:=xlFilterCellColor     '# Applies a filter to the new workbook showing only the purple highlighted cells
ChangesBook.SaveAs Filename:=NewCRChanges                                                                                                                                              '# Returns to the Main Workbook on the RDI sheet
CRScopingWB.Sheets("2.x.x Release Design Intent").Range("$A$1", "$AF$" & RDIEnd).AutoFilter Field:=17, Criteria1:=RGB(128, 0, 128), Operator:=xlFilterCellColor
DisplaySheet.Cells.SpecialCells(xlCellTypeVisible).Interior.ColorIndex = xlNone             '# Removes the purple highlight from the main book that will be distributed
If DisplaySheet.FilterMode = True Then
    DisplaySheet.ShowAllData
End If
For Each Sheet In CRScopingWB.Sheets
    If YYYError = True And XXXError = True Then
        If Sheet.Name <> "2.x.x Release Design Intent" And Sheet.Name <> "AAA YYY Conversion Sheet" And Sheet.Name <> "AAA XXX Conversion Sheet" And Sheet.Name <> "CR List (Class)" Then
            Sheet.Visible = xlSheetHidden
        End If
    ElseIf YYYError = True And XXXError = False Then
        If Sheet.Name <> "2.x.x Release Design Intent" And Sheet.Name <> "AAA YYY Conversion Sheet" And Sheet.Name <> "CR List (Class)" Then
            Sheet.Visible = xlSheetHidden
        End If
    ElseIf YYYError = False And XXXError = True Then
        If Sheet.Name <> "2.x.x Release Design Intent" And Sheet.Name <> "AAA XXX Conversion Sheet" And Sheet.Name <> "CR List (Class)" Then
            Sheet.Visible = xlSheetHidden
        End If
    End If
Next Sheet
ActiveWindow.FreezePanes = False                                                                                                                                    '# Undoes the freeze panes from the previous week
Cells(2, 11).Select                                                                                                                                                 '# Activates Cell K2
ActiveWindow.FreezePanes = True                                                                                                                                     '# Freezes panes from cell K2 to allow ease of editing
CRScopingWB.Sheets("2.x.x Release Design Intent").Range("$A$1", "$AF$" & RDIEnd).AutoFilter Field:=10, Criteria1:="<>*N/A*"                                         '# Filters the RDI sheet against any values in column J that DON'T contain N/A
End Function
