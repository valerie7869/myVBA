Attribute VB_Name = "Module2"
Option Explicit

Sub MaddChart1()
Attribute MaddChart1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MaddChart1 Macro
'

'
    Range("A3").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlPie
    ActiveChart.SetSourceData Source:=Range("Pivot_SoftwareModel!$A$3:$B$8")
End Sub
Sub fixchart()
Attribute fixchart.VB_ProcData.VB_Invoke_Func = " \n14"
'
' fixchart Macro
'

'
    Range("A4").Select
    ActiveSheet.PivotTables("PivotTable15").DisplayErrorString = False
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.text = "Files"
    Range("A1").Select
End Sub
Sub myLoop()

Dim wb As Workbook
Dim Sheet As Worksheet

For Each wb In Workbooks
        ActiveWorkbook.Sheets("Sheet1").Activate
        Sheet.Copy After:=ThisWorkbook.Sheets(1)

Next wb

End Sub


