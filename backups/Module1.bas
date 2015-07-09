Attribute VB_Name = "Module1"
Option Explicit

Function getComment(incell) As String
' aceepts a cell as input and returns its comments (if any) back as a string
On Error Resume Next
getComment = incell.Comment.text
End Function

Sub info() '-- notes only

'excel - truncate notes field - averything after '['
'
'this formula will drop anything to the right of the last Ô[Ô , so good for the notes field
'this assumes there are no Ô$Õ in the field (it substitutes the [ with $, then moves everything to the right of $
'
''=IFERROR(LEFT(K2,FIND("$",SUBSTITUTE(K2,CHAR(91),"$",LEN(K2)-LEN(SUBSTITUTE(K2,CHAR(91),""))))),K2)
'
'SEE http://www.excel-university.com/find-the-last-occurrence-of-a-delimiter-to-retrieve-the-lowest-sub-account-from-quickbooks-in-excel/
'
'
'to list file extensions - in column to right of filename
''=IFERROR(RIGHT(A3,LEN(A3)-FIND("$",SUBSTITUTE(A3,".","$",LEN(A3)-LEN(SUBSTITUTE(A3,".",""))))),Ó")

End Sub
Sub ReverseText()
'select cells, or column to reverse --- the next column is overwritten
'Updateby20131128
Dim rng As Range
Dim xValue, xOut, getChar As String
Dim i, xLen As Integer
Dim xTitleId As String
Dim WorkRng As Range
    On Error Resume Next
    xTitleId = "reverse string tool"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
    For Each rng In WorkRng
        xValue = rng.Value
        xLen = VBA.Len(xValue)
        xOut = ""
        For i = 1 To xLen
            getChar = VBA.Right(xValue, 1)
            xValue = VBA.Left(xValue, xLen - i)
            xOut = xOut & getChar
        Next
        rng(1, 2).Value = xOut ' put result in next column
    Next
End Sub

'

Sub insertFileName()
Attribute insertFileName.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' insertFileName Macro
'
' Keyboard Shortcut: Option+Cmd+z
' formula will strip FULL pathname down to filename only, pasting value only
    ActiveCell.Select
    ActiveCell.Formula = "=MID(CELL(""filename""),SEARCH(""["",CELL(""filename""))+1, SEARCH(""]"",CELL(""filename""))-SEARCH(""["",CELL(""filename""))-1)"
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Select
End Sub
Sub insertPathName()
Attribute insertPathName.VB_Description = "insert current full pathname, and pastes as value"
Attribute insertPathName.VB_ProcData.VB_Invoke_Func = "x\n14"
'
' insertPathName Macro
'
' Keyboard Shortcut: Option+Cmd+x
' formula will insert FULL pathname into cell, pasting value only

    ActiveCell.Select
    ActiveCell.Formula = "=CELL(""filename"")"
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Select
End Sub
Sub insertPathNameDate()
'
' insertPathNameDate Macro
'
' Keyboard Shortcut: Option+Cmd+
' formula will insert FULL pathname into cell, pasting value only AND the date file was saved

'Dim dateSaved As Date
    ActiveCell.Select
    ActiveCell.Formula = "=MID(CELL(""filename""),SEARCH(""["",CELL(""filename""))+1, SEARCH(""]"",CELL(""filename""))-SEARCH(""["",CELL(""filename""))-1)"
    
    'dateSaved = Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time"), "short date")
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    ActiveCell.Offset([0], [1]).Select      ' jump over a col to put date
    ActiveCell.Value = Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time"), "short date")
    
End Sub
Sub SortSheets()
' sorts all sheets by tab name
'Step 1: Declare your Variables
    Dim CurrentSheetIndex As Integer
    Dim PrevSheetIndex As Integer
  
'Step 2: Set the starting counts and start looping
    For CurrentSheetIndex = 1 To Sheets.Count
    For PrevSheetIndex = 1 To CurrentSheetIndex - 1
    
'Step 3: Check Current Sheet against Previous Sheet
    If UCase(Sheets(PrevSheetIndex).name) > _
       UCase(Sheets(CurrentSheetIndex).name) Then
    
'Step 4: If Move Current sheet Before Previous
    Sheets(CurrentSheetIndex).Move _
    Before:=Sheets(PrevSheetIndex)
    End If
    
'Step 5 Loop back around to iterate again
    Next PrevSheetIndex
    Next CurrentSheetIndex

End Sub

Sub CreateTOC()
'starting assumption: sheet named '000-Table Of Contents' already exists (if not, create it)
' this sheet will be overwritten with new TOC

'Step 1: Declare Variables
    Dim i As Long
    
'Step 2 - select the current TOC - should already exist
' if sheet does not already exist, create it before running
    Sheets("000-Table Of Contents").Select
    
'Step 3 - select old TOC and clear it away
    Columns("A:A").Select
    Selection.Clear

'Step 4: Start the i Counter
    For i = 1 To Sheets.Count

'Step 5: Select Next available row
    ActiveSheet.Cells(i, 1).Select

'Step 6: Add the Sheet Name and Hyperlink
    ActiveSheet.Hyperlinks.Add _
    Anchor:=ActiveSheet.Cells(i, 1), _
    Address:="", _
    SubAddress:="'" & Sheets(i).name & "'!A1", _
    TextToDisplay:=Sheets(i).name

'Step 7: Loop back to incrment i
    Next i

End Sub

Sub CopyFiltered2NewBook()

'Step 1: Check for AutoFilter - Exit if none exists
    If ActiveSheet.AutoFilterMode = False Then
    Exit Sub
    End If

'Step 2:  Copy the Autofiltered Range to new workbook
    ActiveSheet.AutoFilter.Range.Copy
    Workbooks.Add.Worksheets(1).Paste

'Step 3: Size the columns to fit
    Cells.EntireColumn.AutoFit

End Sub

Sub NewSheet4EachFiltered()

'Step 1: Declare your Variables
  Dim MySheet As Worksheet
    Dim myRange As Range
    Dim UList As Collection
    Dim UListValue As Variant
    Dim i As Long
    
'Step 2:  Set the Sheet that contains the AutoFilter
    Set MySheet = ActiveSheet
    
    
'Step 3: If the sheet is not auto-filtered, then exit
    If MySheet.AutoFilterMode = False Then
        Exit Sub
    End If
    
  
'Step 4: Specify the Column # that holds the data you want filtered
    Set myRange = Range(MySheet.AutoFilter.Range.Columns(1).Address)
    

'Step 5: Create a new Collection Object
    Set UList = New Collection
    

'Step 6:  Fill the Collection Object with Unique Values
    On Error Resume Next
    For i = 2 To myRange.Rows.Count
    UList.Add myRange.Cells(i, 1), CStr(myRange.Cells(i, 1))
    Next i
    On Error GoTo 0
    

'Step 7: Start looping in through the collection Values
    For Each UListValue In UList
   
   
'Step 8: Delete any Sheets that may have bee previously created
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(CStr(UListValue)).Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    
    
'Step 9:  Filter the Autofilter to macth the current Value
        myRange.AutoFilter Field:=1, Criteria1:=UListValue
    
    
'Step 10: Copy the AutoFiltered Range to new Workbook
        MySheet.AutoFilter.Range.Copy
        Worksheets.Add.Paste
        ActiveSheet.name = Left(UListValue, 30)
        Cells.EntireColumn.AutoFit
        

'Step 11: Loop back to get the next collection Value
    Next UListValue


'Step 12: Go back to main Sheet and removed filters
    MySheet.AutoFilter.ShowAllData
    MySheet.Select
  
  End Sub
  
Sub moveSheetSpecial1()
Attribute moveSheetSpecial1.VB_Description = "move active sheet to All-Reports-BOM-v2"
Attribute moveSheetSpecial1.VB_ProcData.VB_Invoke_Func = "k\n14"

  ActiveSheet.Copy Before:=Workbooks( _
       "All-Reports-BOM-v3.xlsm").Sheets(1)


'  Sheets("BOM Bill of Materials").Copy Before:=Workbooks( _
'       "All-Reports-BOM-v2.xlsm").Sheets(1)
End Sub

Function myGetURL(Cell As Range, Optional default_value As Variant)
 'Lists the Hyperlink Address for a Given Cell
 'If cell does not contain a hyperlink, return default_value
      If (Cell.Range("A1").Hyperlinks.Count <> 1) Then
          myGetURL = default_value
      Else
          myGetURL = Cell.Range("A1").Hyperlinks(1).Address
      End If
End Function
Function HLink(rng As Range) As String
' not yet working
'extract URL from hyperlink
'posted by Rick Rothstein
  If rng(1).Hyperlinks.Count Then HLink = rng.Hyperlinks(1).Address
End Function

Sub CleanUp()

' ascII characters at http://www.asciitable.com/
' chr(10) is line-feed
ActiveSheet.UsedRange.Replace What:=Chr(10), Replacement:=" "
'chr(13) is carraige return
ActiveSheet.UsedRange.Replace What:=Chr(13), Replacement:=" "
' replace a string
'ActiveSheet.UsedRange.Replace What:="Copyright", replacement:="CRight"
End Sub

  
  'Option Base 0 assumed

  'POB: fn with byte array is 17 times faster
  ' see http://stackoverflow.com/questions/4243036/levenshtein-distance-in-excel
  
 Function Levenshtein(value1 As Range, value2 As Range) As Long
 ' Sub Levenshtein(value1 As Range, value2 As Range) 'As Long

  Dim i As Long, j As Long, bs1() As Byte, bs2() As Byte
  Dim string1, string2 As String
  Dim string1_length As Long
  Dim string2_length As Long
  Dim distance() As Long
  Dim min1 As Long, min2 As Long, min3 As Long

 string1 = value1.Value
 string2 = value2.Value
  string1_length = Len(string1)
  string2_length = Len(string2)
  ReDim distance(string1_length, string2_length)
  bs1 = string1
  bs2 = string2

  For i = 0 To string1_length
      distance(i, 0) = i
  Next

  For j = 0 To string2_length
      distance(0, j) = j
  Next

  For i = 1 To string1_length
'      For j = 1 To string2_length
'          'slow way: If Mid$(string1, i, 1) = Mid$(string2, j, 1) Then
'          If bs1((i - 1) * 2) = bs2((j - 1) * 2) Then   ' *2 because Unicode every 2nd byte is 0
'              distance(i, j) = distance(i - 1, j - 1)
'          Else
'              'distance(i, j) = Application.WorksheetFunction.Min _
'              (distance(i - 1, j) + 1, _
'               distance(i, j - 1) + 1, _
'               distance(i - 1, j - 1) + 1)
'              ' spell it out, 50 times faster than worksheetfunction.min
'              min1 = distance(i - 1, j) + 1
'              min2 = distance(i, j - 1) + 1
'              min3 = distance(i - 1, j - 1) + 1
'              If min1 <= min2 And min1 <= min3 Then
'                  distance(i, j) = min1
'              ElseIf min2 <= min1 And min2 <= min3 Then
'                  distance(i, j) = min2
'              Else
'                  distance(i, j) = min3
'              End If
'
'          End Ifdis
'      Next
  Next

  Levenshtein = distance(string1_length, string2_length)

  End Function
Sub Extracthyperlinks()
'2014-march'
'select range (column) with hyperlinks --- warning: links will writeover nextcolumn, same row

Dim rng As Range
Dim WorkRng As Range
On Error Resume Next

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each rng In WorkRng
If rng.Hyperlinks.Count > 0 Then
rng(1, 2).Value = rng.Hyperlinks.Item(1).Address  ' write to same row, next column (change to 1,1 to write over top)
End If
Next
End Sub
Sub vLookUp()
Dim c As Range
Dim cnt As Integer
Debug.Print "starting at " & Now

cnt = 0
    For Each c In Selection
        ' get value from 1 cell left, same row
        c.Formula = "=vlookup(A2,Table13,2,false)"
        cnt = cnt + 1
    Next c
Debug.Print "done at " & Now & "records = " & cnt
End Sub

Sub getTheLink()
Attribute getTheLink.VB_Description = "open new window - cell has address"
Attribute getTheLink.VB_ProcData.VB_Invoke_Func = "k\n14"
' open the link in new window (url address must be in the selected cell)
' ... and run this macro with a shorcut key
'.. no cells will change -  only a link will open in new window, current browser

Dim theLInk As String

On Error GoTo Errorhandler
theLInk = ActiveCell.Value
ThisWorkbook.FollowHyperlink Address:=(theLInk), NewWindow:=True

Exit Sub

Errorhandler:
    MsgBox Err.Description, vbCritical, "Link invalid"
End Sub
=
Sub getOLEXPackageLink()
Attribute getOLEXPackageLink.VB_ProcData.VB_Invoke_Func = "p\n14"
' open the link in new window (url address must be in the selected cell)
' ... and run this macro with a shorcut key
'.. no cells will change -  only a link will open in new window, current browser

Dim theLInk As String
theLInk = "http://olex.openlogic.com/packages/" & ActiveCell.Value

ThisWorkbook.FollowHyperlink Address:=(theLInk), NewWindow:=True
End Sub

Sub getOLEXLicenseLink()
Attribute getOLEXLicenseLink.VB_ProcData.VB_Invoke_Func = "l\n14"
' open the link in new window (url address must be in the selected cell)
' ... and run this macro with a shorcut key
'.. no cells will change -  only a link will open in new window, current browser

Dim theLInk As String
theLInk = "http://olex.openlogic.com/licenses/" & ActiveCell.Value

ThisWorkbook.FollowHyperlink Address:=(theLInk), NewWindow:=True
End Sub

Sub pasteValues()
Attribute pasteValues.VB_Description = "paste values"
Attribute pasteValues.VB_ProcData.VB_Invoke_Func = "v\n14"
Selection.Copy
Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
End Sub


 
Sub GetFileType()
    MsgBox MacScript("tell application ""finder"" " & vbCr & "try" & vbCr _
    & "return file type of file ((choose file) as alias)" & vbCr & "end try" & vbCr _
    & "end tell")
End Sub

Sub Select_File_Or_Files_Mac()
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim MySplit As Variant
    Dim n As Long
    Dim Fname As String
    Dim mybook As Workbook

    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")
    'Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"

    ' In the following statement, change true to false in the line "multiple
    ' selections allowed true" if you do not want to be able to select more
    ' than one file. Additionally, if you want to filter for multiple files, change
    ' {""com.microsoft.Excel.xls""} to
    ' {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
    ' if you want to filter on xls and csv files, for example.
    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
               "set theFiles to (choose file of type " & _
             " {""com.microsoft.Excel.xls""} " & _
               "with prompt ""Please select a file or files"" default location alias """ & _
               MyPath & """ multiple selections allowed true) as string" & vbNewLine & _
               "set applescript's text item delimiters to """" " & vbNewLine & _
               "return theFiles"

    MyFiles = MacScript(MyScript)
    On Error GoTo 0

    If MyFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With

        MySplit = Split(MyFiles, ",")
        For n = LBound(MySplit) To UBound(MySplit)

            ' Get the file name only and test to see if it is open.
            Fname = Right(MySplit(n), Len(MySplit(n)) - InStrRev(MySplit(n), Application.PathSeparator, , 1))
            If bIsBookOpen(Fname) = False Then

                Set mybook = Nothing
                On Error Resume Next
                Set mybook = Workbooks.Open(MySplit(n))
                On Error GoTo 0

                If Not mybook Is Nothing Then
                    MsgBox "You open this file : " & MySplit(n) & vbNewLine & _
                           "And after you press OK it will be closed" & vbNewLine & _
                           "without saving, replace this line with your own code."

'Open loop for action to be taken on all selected workbooks.
For x = 1 To UBound(Fname)
Workbooks.Open (Fname(x))
ActiveWorkbook.Sheets.Copy Before:=ThisWorkbook.Sheets(1)
ActiveWorkbook.Close False
Next x



                    mybook.Close SaveChanges:=False
                End If
            Else
                MsgBox "We skipped this file : " & MySplit(n) & " because it Is already open."
            End If
        Next n
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With
    End If
End Sub

Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Contributed by Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function

Sub PasswordBreaker()
    'Breaks worksheet password protection.
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If ActiveSheet.ProtectContents = False Then
        MsgBox "One usable password is " & Chr(i) & Chr(j) & _
            Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
            Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
         Exit Sub
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Sub

Sub InsertBlueTable()
Attribute InsertBlueTable.VB_ProcData.VB_Invoke_Func = "t\n14"

    Dim tblNewTable As ListObject  ' table for Files table
    Dim rng1 As Range
    Dim tempName, tempName2 As String
    Dim pos As Integer

    'Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set rng1 = ActiveSheet.UsedRange
    Set tblNewTable = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
    tblNewTable.TableStyle = "TableStyleLight9"

    'tempName = ActiveWorkbook.Sheets.Application.ActiveSheet.name
    tempName = ActiveSheet.name
    pos = InStr(tempName, "loas")   ' get postion of loas string in name
    tempName2 = Mid(tempName, pos, 9)  ' isolate loas# string -starts at pos and len = 9
    tempName2 = tempName2 & "-000001"  ' append line # for sorting option
    Range("D2").Select
    ActiveCell.FormulaR1C1 = tempName2
    Columns("D:D").ColumnWidth = 18
    Range("A2").Select  'reset postion at top of col
   '  pos = pos
     
End Sub

Sub TwistOLEXPivot()
Attribute TwistOLEXPivot.VB_ProcData.VB_Invoke_Func = "T\n14"
Dim oSh As Worksheet
Dim tblFIles As ListObject  ' table for Files table
Dim rng1 As Range
Dim rList As Range
' define pivots
Dim PCache1 As PivotCache    ' cache the Files table to use more than 1 piv table
Dim pf As PivotField        ' used to loop thru fields in pivot table
Dim pivBOM As PivotTable   ' for BOM open source bu Package/License

On Error Resume Next  'convert table to range if already exists
ActiveSheet.ListObjects("Table1").unlist
On Error GoTo 0

Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
'Set rng1 = ActiveSheet.UsedRange
' above selection for Set rng1 will make a difference on table boundries
Set tblFIles = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
tblFIles.TableStyle = "TableStyleLight9"  'set blue OLEX style
Range("A1").Select  'reset postion at top of sheet
'add the BOM pivot table sheet
Sheets.Add.name = "Pivot_BOMprep"
' create cache from Files tbl
Set PCache1 = ActiveWorkbook.PivotCaches.Create(xlDatabase, tblFIles)
' create the pivot table from that cache
Set pivBOM = PCache1.CreatePivotTable(TableDestination:=("Pivot_BOMprep!R3C1"))
'blank pivot table ready
'next add pivot table fields for BOM
With pivBOM
    'move package and license into pivot
    With .PivotFields("Confirmed Packages")
        .Orientation = xlRowField
        .Position = 1
    End With
    With .PivotFields("Confirmed Licenses")
        .Orientation = xlRowField
        .Position = 2
    End With
    .AddDataField .PivotFields("Filename"), "Files", xlCount  ' add file count col
    .TableStyle2 = "PivotStyleMedium9"   ' set blue OLEX style
    .ShowDrillIndicators = False    ' turn off drill arrows
End With

Range("A3").Select  'reset postion at top of pivot table
End Sub
Sub Test4()
'Check if Table1 exists and either create it or skip this step to prevent duplicate table error
    Dim ListObj As ListObject
    On Error Resume Next
    Set ListObj = ActiveSheet.ListObjects("Table1")
    On Error GoTo 0
    
    If ListObj Is Nothing Then
            'Establish the table for the find company operation
            Range("A1", LastCG).Select
            Range("C6").Activate
            ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1", LastCG), , xlYes). _
                name = "Table1"
    Else
        'If the table does exist clear filter from column C
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3
    End If

'Step into the table
    Range("Table1[[#Headers],[Enroller Name]]").Select
End Sub

Sub TestMacro3()
'
    With ActiveSheet.PivotTables("PivotTable2")
        .HasAutoFormat = False
        .RowGrand = 7
    End With
    Range("A1").Select
    ExecuteExcel4Macro _
        "(""PivotTable2"",""'Confirmed Packages'[All]"",1,TRUE,TRUE)"
    Range("A3").Select

End Sub
Sub SetDefaultTableStyle2Blue()

ActiveWorkbook.DefaultPivotTableStyle = "PivotStyleMedium9"
ActiveWorkbook.DefaultTableStyle = "TableStyleLight9"

End Sub
Sub FixColWidth()
Attribute FixColWidth.VB_ProcData.VB_Invoke_Func = "x\n14"
'
'use shortcut key to invoke; narrow the selected column (for inside pivot tables
    Selection.ColumnWidth = 105
End Sub
