Attribute VB_Name = "CreateReportMACdev"
Option Explicit
Sub CreateReportMAC_v4_1_dev()
' createReport4.1 - development version
' last updates: Mar-2015
' add 'Software Model' col to Files, add by-License BOM, and SoftwareModel Pivot
' must start with: open 6 sheet OLEX export (tabs: Files, Packages, Licenses, Conflicts, Obligations, Usage)
'note reference to OpenLogicMacros.xlam to move in the Summary page
'  also note: to modify the default summary page, user can add new Summary tab from 'Edit Menu'
'
ActiveWorkbook.DefaultPivotTableStyle = "PivotStyleMedium9"
ActiveWorkbook.DefaultTableStyle = "TableStyleLight9"
Application.DisplayStatusBar = True  ' turn on status bar
Application.StatusBar = "Now creating report - please wait...."
Application.ScreenUpdating = False  ' turn off screen updates while running

    Dim reportWB As Workbook
    Dim rng1 As Range
    Dim rng2 As Range
    Dim mySelRange As Range
    Dim entrytext As String
    Dim message1 As String
' define pivots
    Dim PCache1 As PivotCache    ' cache the Files table to use more than 1 piv table
    Dim pf As PivotField        ' used to loop thru fields in pivot table
    Dim pi As PivotItem
    Dim pivBOM As PivotTable   ' for BOM open source bu Package/License
    Dim pivBOM2 As PivotTable   ' for BOM non-open source
    Dim pivBOMSheetName As String
    Dim pivLicense As PivotTable    ' for BOM by License/Package
    Dim pivSMod As PivotTable   ' for Software Model
    Dim PCache2 As PivotCache   ' cache the Obligations table
    Dim pivOblig As PivotTable    ' for ObligationsC
' define tables
    Dim tblFIles As ListObject  ' table for Files table
    Dim tblObligations As ListObject  ' table for Obligations table
    Dim tblPackages As ListObject   'ditto
    Dim tblLicenses As ListObject
'=======================================================================
    message1 = "WARNING: OLEX 6-tab export expected for successful execution." & vbCrLf
    message1 = message1 & "Do you want to continue?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then GoTo BailOut

   Application.ScreenUpdating = False  ' turn off screen updates while running
    'work on Licenses sheet
    Sheets("Licenses").Select  'go to License tab
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblLicenses = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
'    tblLicenses.TableStyle = "TableStyleLight9"
    tblLicenses.Name = "tblLicenses"
    ' delete unneeded cols
    Range("tblLicenses[Policy]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column
    Range("tblLicenses[Source]").Select
        Selection.Delete Shift:=xlToLeft    ' delete this column

    Columns("A:A").ColumnWidth = 38.43
    Columns("B:B").ColumnWidth = 59.43  ' taxonomy col
    
    Range("A1").Select  'reset postion at top of sheet
'=======================================================================
    'work on Obligations sheet
    Sheets("Obligations").Select  'go to Obligations tab
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblObligations = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
'    tblObligations.TableStyle = "TableStyleLight9"
    tblObligations.Name = "tblObligations"
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
    
    
' now add a pivot summary table
 'add the Obligations Summary pivot table sheet
    Sheets.Add.Name = "Pivot_ObligationsSummary"
    ' create cache from Files tbl
    Set PCache2 = ActiveWorkbook.PivotCaches.Create(xlDatabase, tblObligations)
    ' create the pivot table from that cache
    Set pivOblig = PCache2.CreatePivotTable(TableDestination:=("Pivot_ObligationsSummary!R3C1"))
    'blank pivot table ready
    'next add pivot table fields for BOM
    'On Error Resume Next
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
    End With
       
    'set column withs for pivot
    Columns("A:A").ColumnWidth = 40
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
'    tblFIles.TableStyle = "TableStyleLight9"
    tblFIles.Name = "tblFiles"

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
        
' insert column , or add new col for software model vlookup
    Range("tblFiles[Copyrights]").Select
    Selection.ListObject.ListColumns.Add Position:=5
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Software Model"
    
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
    
'    Range("A1").Select      ' return to top
'=======================================================================
    'work on Packages sheet
    Sheets("Packages").Select  'go to Packages tab
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblPackages = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
'    tblPackages.TableStyle = "TableStyleLight9"
    tblPackages.Name = "tblPackages"
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
    
' do the vlookup for software model

    Range("A1").Select  'reset postion at top of sheet

' go back to Files
    Sheets("Files").Select  'go to Files tab
    Range("E2").Select  ' position to place vlookup formula for software model
'insert the vlookup to move in software model
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],tblPackages,3,FALSE)"
    Range("E2").Select  ' position to view vlookup formula for software model
'///////////////////////////////////////////////////////////////////////////////////////////
   MsgBox "Macro paused. Allow time for VLOOKUP to complete, then OK to continue.", vbApplicationModal
'////////////////////////////////////////////////////////////////////////////////////////////
' go back to file sheet to insert new pivot
    Sheets("Files").Select  'go to Files tab
    Range("A1").Select  'reset postion at top of sheet
    'add the BOM pivot table sheet
    pivBOMSheetName = "Pivot_BOMprep"
    Sheets.Add.Name = pivBOMSheetName
    ' create cache from Files tbl
    Set PCache1 = ActiveWorkbook.PivotCaches.Create(xlDatabase, tblFIles)
    ' create the pivot table from that cache
    Set pivBOM = PCache1.CreatePivotTable(TableDestination:=("Pivot_BOMprep!R3C1"))
    'blank pivot table ready
    'next add pivot table fields for BOM
    'On Error Resume Next
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
        With .PivotFields("Software Model")
            .Orientation = xlPageField
            .Position = 1
            .EnableMultiplePageItems = True
        End With

        'add the count of files
        .AddDataField .PivotFields("Filename"), "Files", xlCount  ' add file count col
 '       .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
         
        With .PivotFields("Software Model")
           For Each pi In .PivotItems
           ' filter out all items not "Open Source" Software Model, such as Freeware, Shareware, Commercial
                If pi.Name <> "Open Source" Then .PivotItems(pi.Name).Visible = False
'               If pi.name = "Freeware" Then .PivotItems(pi.name).Visible = False
'               If pi.name = "In-house" Then .PivotItems(pi.name).Visible = False
'               If pi.name = "Commercial" Then .PivotItems(pi.name).Visible = False
           Next
        End With
    End With
       
    'set column withs for pivot
    Columns("A:A").ColumnWidth = 40
    Columns("B:B").ColumnWidth = 8
    Columns("C:C").ColumnWidth = 4
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
        With .PivotFields("Software Model")
            .Orientation = xlPageField
            .Position = 1
            .EnableMultiplePageItems = True ' allow selection of items
        End With
        
        'add the count of files
        .AddDataField .PivotFields("Filename"), "Files", xlCount  ' add file count col
'        .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
        With .PivotFields("Software Model")
           For Each pi In .PivotItems  ' must test if these field exists, then uncheck
               If pi.Name = "Open Source" Then .PivotItems(pi.Name).Visible = False
           Next
        End With
            
    End With
    'set column withs for pivot
    Columns("D:D").ColumnWidth = 35
    Columns("E:E").ColumnWidth = 8
    
'=======================================================================
    'add the byLIcense pivot table - all Packages by LIcense
    ' create the pivot table from already created cache
    ' use same sheet - place beside first pivot - column 6
    Set pivLicense = PCache1.CreatePivotTable(TableDestination:=("Pivot_BOMprep!R3C7"))
    'blank pivot table ready
    'next add pivot table fields for BOM
    With pivLicense
        'move package and license into pivot
        With .PivotFields("Confirmed Licenses")
            .Orientation = xlRowField
            .Position = 1
        End With
        With .PivotFields("Confirmed Packages")
            .Orientation = xlRowField
            .Position = 2
        End With
        
        'add the count of files
        .AddDataField .PivotFields("Filename"), "Files", xlCount  ' add file count col
'        .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
            
    End With
    'set column withs for pivot
    Columns("F:F").ColumnWidth = 4
    Columns("G:G").ColumnWidth = 35
    Columns("H:H").ColumnWidth = 8
    
    Range("A1").Select  'reset postion at top of sheet
'\\\\\\\\\End of BOM pivot tables\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' insert new software model pivot sheet and table
    Sheets("Pivot_BOMprep").Select  ' select sheet to insert before
    'add the BOM pivot table sheet
    Sheets.Add.Name = "Pivot_SoftwareModel"
    ' create cache from Files tbl
    ' create the pivot table for Software Model from that same Files cache
    Set pivSMod = PCache1.CreatePivotTable(TableDestination:=("Pivot_SoftwareModel!R3C1"))
 '   blank pivot table ready
 '   next add pivot table fields for BOM
    With pivSMod  'pivot table for Software Model
        'move package and license into pivot
        With .PivotFields("Software Model")
            .Orientation = xlRowField
            .Position = 1
        End With
        .AddDataField .PivotFields("Filename"), "Files", xlCount  ' add file count col
 '       .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
        .RowGrand = False
        .ColumnGrand = False
    End With
    
    pivSMod.RowGrand = False        ' turn off totals cuz they effect chart wrongly

    'set column withs for pivot
    Columns("A:A").ColumnWidth = 35
    Columns("B:B").ColumnWidth = 10
    Range("A1").Select  'reset postion at top of she
'\\\\\\\\\End of Software Model pivot table\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'     go back to 1st sheet
    Sheets("Pivot_SoftwareModel").Select
    Range("A3").Select  'reset postion at top of sheet
    ' select pivot
    Range("A3").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlPie
    ActiveChart.ChartTitle.text = "Files by Software Model"
    ActiveChart.ChartTitle.Characters.Font.Size = 18
    ActiveChart.Legend.Format.TextFrame2.TextRange.Font.Size = 13  ' Legend label txt size

    ActiveSheet.Shapes("Chart 1").ScaleWidth 0.9, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 1").ScaleHeight 1.25, msoFalse, _
        msoScaleFromTopLeft
        
'    ActiveWorkbook.ShowPivotTableFieldList = False
    'move chart over
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 1").IncrementLeft 74.25
    ActiveSheet.Shapes("Chart 1").IncrementTop -45.75
    
    Range("A4").Select
'    ActiveSheet.PivotTables(pivSMod).DisplayErrorString = False
    
    ' set better columns for pivot table
    Columns("A:A").ColumnWidth = 18
    Columns("B:B").ColumnWidth = 8
    Range("A1").Select  'reset postion at top of sheet
    
'=====================================================
' copy over a summary sheet
'
Sheets(pivBOMSheetName).Select  'go to pivot table tab,

'Application.ActiveWorkbook is the workbook used for report
Application.Workbooks("OpenLogicMacros.xlam").Worksheets("Summary").Copy _
    Before:=Application.ActiveWorkbook.Sheets("Pivot_SoftwareModel")

'reportWB.Activate       ' come back to report
Range("B2").Select      'rest here on summary page
'
'/// prepare exit   ///////////////////////////////////////////////////////////////
    ActiveWindow.TabRatio = 0.955   ' make wider the tab view along bottom
    Application.ScreenUpdating = True  ' turn back on screen
    MsgBox "Complete.  If necessary, modify Packages tab: Software Model, then Refresh All."
BailOut:
    Application.StatusBar = ""   ' clear
    Application.StatusBar = False   ' reset from top

End Sub