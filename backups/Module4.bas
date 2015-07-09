Attribute VB_Name = "Module4"
Option Explicit

Function SUMALLSHEETS(Optional Cell As Range) As Variant
' Sums a cell across all sheets
 Dim i As Long
 Dim WkSht As Worksheet, WkBook As Workbook
 Dim ArgIsMissing As Boolean
   Application.Volatile
    i = -1
    If Cell Is Nothing Then
       Set Cell = Application.Caller
       ArgIsMissing = True
    End If
    'Set WkBook = Cell.Parent.Parent
    Set WkBook = Active.Workbook
    For Each WkSht In WkBook.Worksheets
       If Not (ArgIsMissing And (WkSht.name = Cell.ActiveWorkbook.name)) Then
          If Not IsEmpty(WkSht.Range(Cell.Address)) Then
          i = i + 1
          SUMALLSHEETS = SUMALLSHEETS + WkSht.Range(Cell.Address).Value
          End If
       End If
    Next WkSht
    If i = -1 Then SUMALLSHEETS = 0
End Function
