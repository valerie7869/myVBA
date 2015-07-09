Attribute VB_Name = "Module3"
Option Explicit

Sub TempCopyAllSheetsOpen()
'stored here - but must run within its own workbook
    Dim wkb As Workbook
    Dim sWksName As String

    sWksName = "Sheet1"
    For Each wkb In Workbooks
        If wkb.name <> ThisWorkbook.name Then
            wkb.Worksheets(sWksName).Copy _
              Before:=ThisWorkbook.Sheets(1)
        End If
    Next
    Set wkb = Nothing
End Sub


