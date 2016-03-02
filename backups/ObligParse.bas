Attribute VB_Name = "ObligParse"

Sub ObligationParse()

    Dim newSheet As Excel.Worksheet
    Dim myTable As ListObject
    Dim myRow As ListRow
    Dim sheetName As String
    Dim rowIndex As Long
    Dim nodeCount As Integer
    Dim nodeIndex As Integer
    Dim nameNode As Integer
    Dim rowCount As Long
    
    Dim ObCol As Integer
    Dim nextNodeName As Boolean
    
    Dim toHit As Boolean
    Dim lastrow As Long
    Dim thisNode As String
    Dim ObText As String
    Dim ObTextClean As String
    
    Dim message1 As String
    message1 = "WARNING: TABLE with 'Respnsible Packages' column in and newObs sheet is created for output." & vbCrLf
    message1 = message1 & "Do you want to continue?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
    
     ActiveWorkbook.Sheets(1).Select
     Set myTable = ActiveSheet.ListObjects(1)
     Set myRow = myTable.ListRows(1)    ' init to row 1
    
    'Set newSheet = ActiveSheet.Add      ' new sheet to list history
    Worksheets.Add(After:=Worksheets(1)).name = "newObs"
     Set newSheet = ActiveWorkbook.Sheets("newObs")
    ' insert column headings
    With newSheet
        lastrow = .Range("A" & .Rows.Count).End(xlUp).Row + 1
        .Cells(lastrow, 1).Value = myTable.Range([1], [1])
        .Cells(lastrow, 2).Value = myTable.Range([1], [2])
        .Cells(lastrow, 3).Value = myTable.Range([1], [3])
        .Cells(lastrow, 4).Value = myTable.Range([1], [4])
        .Cells(lastrow, 5).Value = myTable.Range([1], [5])
        .Cells(lastrow, 6).Value = myTable.Range([1], [6])
        .Cells(lastrow, 7).Value = myTable.Range([1], [7])
        .Cells(lastrow, 8).Value = myTable.Range([1], [8])
        .Cells(lastrow, 9).Value = myTable.Range([1], [9])

    End With


     ActiveWorkbook.Sheets(1).Select
    ' Set myTable = ActiveSheet.ListObjects(1)
    ' Set myRow = myTable.ListRows(2)    ' init to row 2 - past headers
     myRow.Range.Select

     ' setup column pointers based on olex-given column headings
     ObCol = myTable.ListColumns("Responsible Packages").Index   'get column number for indexing

        rowCount = myTable.ListRows.Count
        checkCount = 0
        
     For rowIndex = 2 To (rowCount + 1) Step 1 ' work each row loop
        ObText = Cells(rowIndex, ObCol).text       ' get the ersponsible packages field
        nodeCount = dhCountTokens(ObText, ",")
        ' start node loop
        'timeNode = -1
        'nameNode = -1
        
        For nodeIndex = 1 To nodeCount Step 1           ' loop thru each node
            thisNode = dhExtractString(ObText, nodeIndex, ",")
            
       ' test to see if done-done.
            If rowIndex - 1 = rowCount Then
                If nodeIndex = nodeCount Then
                    If thisNode = "" Then
                        Exit Sub
                    End If
                End If
            End If
                                      
  '          GoTo getnextNode
gotNewResolution:
            ' so write the prior resolution
            ' put history in its own row next sheet
                        
            With newSheet
                        lastrow = .Range("A" & .Rows.Count).End(xlUp).Row + 1
                        .Cells(lastrow, 1).Value = myTable.Range(rowIndex, [1])
                        .Cells(lastrow, 2).Value = myTable.Range(rowIndex, [2])
                        .Cells(lastrow, 3).Value = myTable.Range(rowIndex, [3])
                        .Cells(lastrow, 4).Value = myTable.Range(rowIndex, [4])
                        .Cells(lastrow, 5).Value = myTable.Range(rowIndex, [5])
                        .Cells(lastrow, 6).Value = myTable.Range(rowIndex, [6])
                        .Cells(lastrow, 7).Value = myTable.Range(rowIndex, [7])
                        .Cells(lastrow, 8).Value = thisNode
                        .Cells(lastrow, 9).Value = myTable.Range(rowIndex, [9])
            End With

getnextNode:
           Next nodeIndex
    Next rowIndex
'Debug.Print "Tar nodes fixed/dropped: " & checkCount
End Sub