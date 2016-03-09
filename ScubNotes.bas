Sub ScrubNotes_v1()
'reads note column and removes target string: [author-date] string
' this is example of string to target for cutting out of notes:  [Auditor Val, 2014-10-07 21:07:33 UTC]
' we will identify by the brackets and the "UTC]" to test for valid target

Const delim As String = "*UTC]"

    Dim thisSheet As Excel.Worksheet
    Dim myTable As ListObject
    Dim myRow As ListRow

    Dim sheetName As String
    Dim rowIndex As Long
    Dim rowCount As Long
    
    Dim notesCol As Integer
    Dim noteHit As Boolean
    Dim validHit As Boolean
    Dim i, l As Long
    
    Dim thisChar As String
    Dim note As String
    Dim tempNote As String
    Dim newNote As String
    Dim keep As Boolean
    Dim mess1 As String
    
    mess1 = "REQUIRED: Table with 'Notes' column." & vbCrLf
    mess1 = mess1 & "WARNING: 'Notes' column will be modified to exclude: [auditor-date-time] text." & vbCrLf
    mess1 = mess1 & "TIP:  copy 'Notes' column to guard against unexpected results." & vbCrLf & vbCrLf
    
    mess1 = mess1 & "Continue?"
    If MsgBox(mess1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
    
     Set myTable = ActiveSheet.ListObjects(1)
     Set myRow = myTable.ListRows(1)    ' init to row 1
     myRow.Range.Select
     
     keep = True
     validHit = False
     newNote = ""
     tempNote = ""

     ' setup column pointers based on olex-given column headings
     notesCol = myTable.ListColumns("Notes").Index 'get column number for indexing
        rowCount = myTable.ListRows.Count
 '       Debug.Print "Starting at : " & Now & " - rows to process: " & rowCount
 
     For rowIndex = 1 To rowCount Step 1  ' work each row loop
            newNote = ""
            keep = True
            note = Cells(rowIndex + 1, notesCol).text       ' get the notes field

            noteHit = ISLIKE(note, delim)  ' got a 'UTC]' in the note?
            If noteHit = False Then GoTo doneHere
            
            l = Len(note)
            For i = 1 To l    ' go thru each character
                thisChar = Mid(note, i, 1)   'get 1 char at position i
                If thisChar = "[" Then keep = False
                If thisChar = "]" Then keep = True
                
                If keep = False Then tempNote = tempNote & thisChar     'inside [] so save into temp
                
                If thisChar = "]" Then      'test if real
                    tempNote = tempNote & thisChar
                  ' now see if section inside brackets is false hit
                  validHit = ISLIKE(tempNote, delim)     'UTC]' there?  or else must be false hit
                  
                  If validHit = True Then
                    thisChar = ""     'reset & drop the ]
                  End If
                  
                  If validHit = False Then  'keep this in note
                      newNote = newNote & tempNote    'false hit, so all add back to note
                      tempNote = ""   'reset
                      keep = True    'reset
                      GoTo done1
                  End If        'validHit
                  
                tempNote = ""       'reset
               End If       ' thisChar = "]"

                If keep = True Then
                    newNote = newNote & thisChar
                End If
done1:
            Next i
            
            Cells(rowIndex + 1, notesCol).Value = newNote
doneHere:
    tempNote = ""   'reset
    validHit = False        ' reset
    keep = True
    Next rowIndex
End Sub

