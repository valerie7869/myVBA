Attribute VB_Name = "CreateReport_v4"
Option Explicit

Sub CreateReportMAC_v4()
' createReport4-
' last updates: Feb 2015
' must start with: open 6 sheet OLEX export (tabs: Files, Packages, Licenses, Conflicts, Obigations, Usage)
'
    Application.DisplayStatusBar = True  ' turn on status bar
'    Application.StatusBar = "Now creating report...."
    Application.ScreenUpdating = False  ' turn off screen updates while running
    
    Dim myform As Object    ' use userform to pause and allow user changes
    Dim rng1 As Range
    Dim rng2 As Range
    Dim mySelRange As Range
    Dim entrytext As String
' define pivots
    Dim PCache1 As PivotCache    ' cache the Files table to use more than 1 piv table
    Dim pf As PivotField        ' used to loop thru fields in pivot table
    Dim pivBOM As PivotTable   ' for BOM open source bu Package/License
    Dim pivBOM2 As PivotTable   ' for BOM non-open source
    Dim pivLicense As PivotTable    ' for BOM by License/Package
    Dim pivSMod As PivotTable   ' for Software Model
    
    Dim PCache2 As PivotCache   ' cache the Obligations table
    Dim pivOblig As PivotTable    ' for Obligations
' define tables
    Dim tblFIles As ListObject  ' table for Files table
    Dim tblObligations As ListObject  ' table for Obligations table
    Dim tblPackages As ListObject   'ditto
    Dim tblLicenses As ListObject
'=======================================================================
    'work on Licenses sheet
    Sheets("Licenses").Select  'go to License tab
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblLicenses = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
    tblLicenses.TableStyle = "TableStyleLight9"
    tblLicenses.name = "tblLicenses"
    ' delete unneeded cols
    Range("tblLicenses[Policy]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblLicenses[Source]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column

    Columns("A:A").ColumnWidth = 38.43
    Columns("B:B").ColumnWidth = 59.43  ' taxonomy col
    
    With rng1.Font  ' set all to font size 11
        .Size = 11
    End With
    
    Range("A1").Select  'reset postion at top of sheet
'=======================================================================
    'work on Obligations sheet
    Sheets("Obligations").Select  'go to Obligations tab
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblObligations = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
    tblObligations.TableStyle = "TableStyleLight9"
    tblObligations.name = "tblObligations"
' format the obligations sheet
    Columns("A:A").ColumnWidth = 34.43
    Columns("B:B").ColumnWidth = 23.57
    Columns("C:C").ColumnWidth = 53.29
    Columns("D:D").ColumnWidth = 23.57
    Columns("E:E").ColumnWidth = 23.14
    Columns("F:F").ColumnWidth = 21.29
    Columns("F:F").ColumnWidth = 24.57
    Columns("G:G").ColumnWidth = 33.14
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Columns("G:G").ColumnWidth = 46.43
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A:I").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("tblObligations[Description]").Select
    'replace the html in Description column to make it look better
    Selection.Replace What:="<br>", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    Selection.Replace What:="<p>", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    Selection.Replace What:="</p>", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    Selection.Replace What:="<em>", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    Selection.Replace What:="</em>", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    Selection.Replace What:="", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
 
    Range("A1").Select  'reset postion at top of sheet
'
' create Obligatons pivot for Obligations Summary Sheet
'============================================================================
    Sheets("Obligations").Select  'go to Obligations tab and insert before
    'add the Obligations pivot table sheet
    Sheets.Add.name = "Pivot_ObligationSummary"
    ' create cache from Obligation tbl
    Set PCache2 = ActiveWorkbook.PivotCaches.Create(xlDatabase, tblObligations)
    Set pivOblig = PCache2.CreatePivotTable(TableDestination:=("Pivot_ObligationSummary!R3C1"))
    'blank pivot table ready
    'next add pivot table fields
    With pivOblig
        'move package and license into pivot
        With .PivotFields("License")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("Name")
            .Orientation = xlRowField
            .Position = 2
        End With

        .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
    End With

'    'set column withs for pivot
    Columns("A:A").ColumnWidth = 25
'    Columns("B:B").ColumnWidth = 10
    Range("A1").Select  'reset postion at top of sheet

'============================================================================
    ' delete conflicts tab and usage tab
    Application.DisplayAlerts = False  ' do not ask to confirm deletes
    Sheets("Usage").Select
        ActiveWindow.SelectedSheets.Delete  ' delete it
    Sheets("Conflicts").Select
        ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True    ' turn back on
'=======================================================================
' go format Files sheet
    Sheets("Files").Select  'go to files tab and format it
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblFIles = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
    tblFIles.TableStyle = "TableStyleLight9"
    tblFIles.name = "tblFiles"

'   set Files table style, remove unneeded columns
    ActiveSheet.ListObjects("tblFiles").TableStyle = "TableStyleLight9"
    Range("tblFiles[Language]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblFiles[Was Scanned]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblFiles[Status]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblFiles[OSS Match Count]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblFiles[License Matches]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    'adjusts col widths
    Range("tblFiles[Path]").Select
        Selection.ColumnWidth = 85.11
    Range("tblFiles[Filename]").Select
        Selection.ColumnWidth = 15.11
    Range("tblFiles[Confirmed Packages]").Select
        Selection.ColumnWidth = 29.11
    Range("tblFiles[Confirmed Licenses]").Select
        Selection.ColumnWidth = 29.11
    Range("tblFiles[Copyrights]").Select
        Selection.ColumnWidth = 20.11
    Range("tblFiles[Notes]").Select
        Selection.ColumnWidth = 20.11
        
' sort the packages col in prep for vlookup to get software model from packages table
    Range("tblFiles[Confirmed Packages]").Select
    With ActiveWorkbook.Worksheets("Files").ListObjects("tblFiles").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("tblFiles[Confirmed Packages]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E2").Select  ' position to place vlookup formula for software model
    
   Range("A1").Select      ' return to top
'=======================================================================
    'work on Packages sheet
    Sheets("Packages").Select  'go to Packages tab
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblPackages = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
    tblPackages.TableStyle = "TableStyleLight9"
    tblPackages.name = "tblPackages"
    ' delete unneeded cols
    Range("tblPackages[Certification]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblPackages[Available Support]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblPackages[Policy]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblPackages[Source]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
' set columns for viewing
    Columns("A:A").ColumnWidth = 35.43
    Columns("B:B").ColumnWidth = 50.43
    Columns("C:C").ColumnWidth = 20.43
' sort on package name
    Range("tblPackages[Name]").Select
    With ActiveWorkbook.Worksheets("Packages").ListObjects("tblPackages").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("tblPackages[Name]"), SortOn:=xlSortOnValues, Order:=xlAscending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("Packages").Select
    Range("tblPackages[#All]").Select
    With Selection.Font
        .Size = 11
    End With

Range("A1").Select  'reset postion at top of sheet
    
'////////////////////////////////////////////////////////////////////////////////////////////
' go back to file sheet to insert new pivot
    Sheets("Files").Select  'go to Files tab
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

        'add the count of files
        .AddDataField .PivotFields("Filename"), "Files", xlCount  ' add file count col
        .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
    End With
    'set column withs for pivot
    Columns("A:A").ColumnWidth = 50
    Columns("B:B").ColumnWidth = 15
    Range("A1").Select  'reset postion at top of sheet
'=======================================================================
    'add the BOM2 pivot table - non open source
    ' create the pivot table from already created cache
    ' use same sheet - place beside first pivot - column 4
    Set pivBOM2 = PCache1.CreatePivotTable(TableDestination:=("Pivot_BOMprep!R3C4"))
    'blank pivot table ready
    'next add pivot table fields for BOM
    With pivBOM2
        'move package and license into pivot
        With .PivotFields("Confirmed Packages")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("Confirmed Licenses")
            .Orientation = xlRowField
            .Position = 2
        End With
        
        'add the count of files
        .AddDataField .PivotFields("Filename"), "Files", xlCount  ' add file count col
        .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
'        .ShowPivotTableFieldList = False    ' turn off field list
    End With
    'set column withs for pivot
    Columns("D:D").ColumnWidth = 30
    Columns("E:E").ColumnWidth = 10
    Range("A1").Select  'reset postion at top of sheet
'\\\\\\\\\End of BOM pivot tables\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'/// prepare exit   ///////////////////////////////////////////////////////////////
    Application.StatusBar = False   ' reset from top
    ActiveWindow.TabRatio = 0.745   ' make wider the tab view along bottom
    Application.ScreenUpdating = True  ' turn back on screen
    MsgBox "Done"

End Sub
