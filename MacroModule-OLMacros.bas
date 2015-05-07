Attribute VB_Name = "OpenLogicMacros"
Option Explicit
Sub ExtractExtensions_v1()
'reads filename column and writes to extension column
' MUST HAVE columns: Filename, extionsion
'
    Dim thisSheet As Excel.Worksheet
    Dim myTable As ListObject
    Dim myRow As ListRow

    Dim sheetName As String
    Dim rowIndex As Long
    Dim checkCount As Long
    Dim nodeCount As Integer
    Dim nodeIndex As Integer
    Dim rowCount As Long
    
    Dim FilenameCol As Integer
    Dim extensionCol As Integer
    
    Dim thisNode As String
    Dim FilenameText As String
    
    Dim message1 As String
    message1 = "WARNING: 'extension' column expected for output and." & vbCrLf
    message1 = message1 & "                  'Filename' column expected for parsing." & vbCrLf
    message1 = message1 & "Do you want to continue?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
    
    
     Set myTable = ActiveSheet.ListObjects(1)
     Set myRow = myTable.ListRows(1)    ' init to row 1
     myRow.Range.Select

     ' setup column pointers based on olex-given column headings
     FilenameCol = myTable.ListColumns("Filename").Index 'get column number for indexing
     extensionCol = myTable.ListColumns("extension").Index 'get column number for indexing

    rowCount = myTable.ListRows.Count
    checkCount = 0
             
    For rowIndex = 1 To rowCount Step 1  ' work each row loop
        'myTable.ListRows(rowIndex).Range.Select       ' get/select each row
        FilenameText = Cells(rowIndex + 1, FilenameCol).text       ' get the filename field
        nodeCount = dhCountTokens(FilenameText, ".")
        ' start node loop
        For nodeIndex = 1 To nodeCount Step 1           ' loop thru each node
            thisNode = dhExtractString(FilenameText, nodeIndex, ".")
            If nodeIndex = nodeCount Then   ' got last node?
                     Cells(rowIndex + 1, extensionCol).FormulaR1C1 = thisNode
            End If
        Next nodeIndex
    Next rowIndex
'Debug.Print "Tar nodes fixed/dropped: " & checkCount

End Sub

Sub SetDefaultTable2Blue()

ActiveWorkbook.DefaultPivotTableStyle = "PivotStyleMedium9"
ActiveWorkbook.DefaultTableStyle = "TableStyleLight9"

End Sub
Sub ExtractHyperlink_v1()
'updated Feb2015
'2014-march'
'select range (column) with hyperlinks --- warning: links will writeover nextcolumn, same row
' read the selected column, and overwrite the column headed: link

Dim rng As Range
Dim WorkRng As Range
Dim linkCol As Integer
Dim myTable As ListObject

Dim message1 As String
message1 = "Selected column is parsed for hyperlinks and ." & vbCrLf
message1 = message1 & " WARNING: 'link' column must be present for output." & vbCrLf
message1 = message1 & "Do you still want to continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

'On Error Resume Next
Set myTable = ActiveSheet.ListObjects(1)
linkCol = myTable.ListColumns("link").Index     ' get the link column number to write to
Set WorkRng = Application.Selection
'Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each rng In WorkRng
If rng.Hyperlinks.Count > 0 Then
rng(1, linkCol).Value = rng.Hyperlinks.Item(1).Address  ' write to same row, 'link' column (change to 1,1 to write over top)
End If
Next
End Sub

Sub ExtractVersionDLL_v1()
'reads path string (optionally include filename to pick up dll's and keeps/writes...
'...all nodes containing psuedo version numbers...
' ...and dll's

'
Const jarString As String = "*.jar"                         ' this is checking in node only
Const dllString As String = "*.dll"     ' no wildcard at end
Const zipString As String = "*.zip*"
Const jsString As String = "*.js"       ' no wildcard at end

Const versString0 As String = "*.0*"
Const versString1 As String = "*.1*"
Const versString2 As String = "*.2*"
Const versString3 As String = "*.3*"
Const versString4 As String = "*.4*"
Const versString5 As String = "*.5*"

Const versString00 As String = "0.*"
Const versString10 As String = "1.*"
Const versString20 As String = "2.*"
Const versString30 As String = "3.*"
Const versString40 As String = "4.*"
Const versString50 As String = "5.*"
      
Const versString000 As String = "*0.*"
Const versString100 As String = "*1.*"
Const versString200 As String = "*2.*"
Const versString300 As String = "*3.*"
Const versString400 As String = "*4.*"
Const versString500 As String = "*5.*"

    Dim thisSheet As Excel.Worksheet
    Dim myTable As ListObject
    Dim myRow As ListRow

    Dim sheetName As String
    Dim rowIndex As Long
    Dim checkCount As Long
    Dim nodeCount As Integer
    Dim nodeIndex As Integer
    Dim rowCount As Long
    
    Dim pathCol As Integer
    Dim versionCol As Integer
    Dim jarHit As Boolean
    Dim versionHit As Boolean
    Dim dllHit As Boolean
    Dim zipHit As Boolean
    Dim jsHit As Boolean
    
    Dim thisNode As String
    Dim versionNode As String
    Dim tarNode As String
    Dim pathText As String
     
    Dim message1 As String
    message1 = "WARNING: 'version-dll' column expected for output and." & vbCrLf
    message1 = message1 & "                  'Path' column expected for parsing." & vbCrLf
    message1 = message1 & "Do you want to continue?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
     
     Set myTable = ActiveSheet.ListObjects(1)
     Set myRow = myTable.ListRows(1)    ' init to row 1
     myRow.Range.Select

     ' setup column pointers based on olex-given column headings
     pathCol = myTable.ListColumns("Path").Index 'get column number for indexing
     versionCol = myTable.ListColumns("version-dll").Index 'get column number for indexing

        rowCount = myTable.ListRows.Count
        checkCount = 0
        
 '       Debug.Print "Starting at : " & Now & " - rows to process: " & rowCount
     
     For rowIndex = 1 To rowCount Step 1  ' work each row loop
            jarHit = False
            versionHit = False
            dllHit = False
            zipHit = False
            jsHit = False
            
            'myTable.ListRows(rowIndex).Range.Select       ' get/select each row
            pathText = Cells(rowIndex + 1, pathCol).text       ' get the path field

            jarHit = ISLIKE(pathText, jarString)  ' got a .jar in the pathname?
            If jarHit = True Then GoTo doneHere
            
            dllHit = ISLIKE(pathText, dllString)  ' got a .dll file?
            If dllHit = True Then GoTo doneHere
            
'            zipHit = ISLIKE(pathText, zipString)  ' got a .zip in the pathname?
'            If zipHit = True Then GoTo doneHere
'
            jsHit = ISLIKE(pathText, jsString)  ' got a .js file?
            If jsHit = True Then GoTo doneHere
           
            versionHit = ISLIKE(pathText, versString0)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString00)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString1)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
         
            versionHit = ISLIKE(pathText, versString2)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString3)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString4)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString5)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString10)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
         
            versionHit = ISLIKE(pathText, versString20)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString30)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString40)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString50)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString100)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
         
            versionHit = ISLIKE(pathText, versString200)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString300)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString400)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString500)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
          
GoTo getnextRec     ' not a row with a version or jar
            
doneHere:
          versionNode = ""        ' initialize
          versionHit = False      'reset
           jarHit = False              'reset
           dllHit = False
           zipHit = False
           jsHit = False
           
        ' fall thru only to find the node(s) we want to save in version column
        nodeCount = dhCountTokens(pathText, "/")
        For nodeIndex = 1 To nodeCount Step 1           ' loop thru each node
            
          thisNode = dhExtractString(pathText, nodeIndex, "/")
              
        If nodeIndex = nodeCount Then  ' if at last node then take dll, .jars, and .zip
              jarHit = ISLIKE(thisNode, jarString)  ' got the jar node?
                If jarHit = True Then
                    versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                    GoTo gotVersionField
                End If

              dllHit = ISLIKE(thisNode, dllString)  ' got the dll node?
                If dllHit = True Then
                    versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                    GoTo gotVersionField
                End If

'              zipHit = ISLIKE(thisNode, zipString)  ' got the zip node?
'                If zipHit = True Then
'                    versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
'                    GoTo gotVersionField
'                End If
                
              jsHit = ISLIKE(thisNode, jsString)  ' got the js node?
                If jsHit = True Then
                    versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                    GoTo gotVersionField
                End If
            ' only falling thru - we're on last node and not caught above, so done with this row
            GoTo gotVersionField        ' get out
        End If
                      
              versionHit = ISLIKE(thisNode, versString0)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString1)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString2)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString3)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString4)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString5)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
               
               versionHit = ISLIKE(thisNode, versString00)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString10)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString20)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString30)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString40)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString50)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              
              
              versionHit = ISLIKE(thisNode, versString000)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString100)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString200)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString300)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString400)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString500)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
                      
getnextNode:
           Next nodeIndex

gotVersionField:
' put newPath in its own column
     Cells(rowIndex + 1, versionCol).FormulaR1C1 = versionNode
getnextRec:
    Next rowIndex
'Debug.Print "Tar nodes fixed/dropped: " & checkCount
End Sub

Sub ExtractVersion_v1()
'reads path string and keeps/writes...
'...all nodes containing psuedo version numbers...
' MUST HAVE columns: Path version
'
Const jarString As String = "*.jar*"                         ' this is checking in node only
Const dllString As String = "*.dll"     ' no wildcard at end
Const zipString As String = "*.zip*"
Const tgzString As String = "*tgz.*"
Const jsString As String = "*.js"       ' no wildcard at end
Const classString As String = "*.class"

Const versString0 As String = "*.0*"
Const versString1 As String = "*.1*"
Const versString2 As String = "*.2*"
Const versString3 As String = "*.3*"
Const versString4 As String = "*.4*"
Const versString5 As String = "*.5*"

Const versString00 As String = "0.*"
Const versString10 As String = "1.*"
Const versString20 As String = "2.*"
Const versString30 As String = "3.*"
Const versString40 As String = "4.*"
Const versString50 As String = "5.*"
      
Const versString000 As String = "*0.*"
Const versString100 As String = "*1.*"
Const versString200 As String = "*2.*"
Const versString300 As String = "*3.*"
Const versString400 As String = "*4.*"
Const versString500 As String = "*5.*"

    Dim thisSheet As Excel.Worksheet
    Dim myTable As ListObject
    Dim myRow As ListRow

    Dim sheetName As String
    Dim rowIndex As Long
    Dim checkCount As Long
    Dim nodeCount As Integer
    Dim nodeIndex As Integer
    Dim rowCount As Long
    
    Dim pathCol As Integer
    Dim versionCol As Integer
    Dim jarHit As Boolean
    Dim versionHit As Boolean
    Dim dllHit As Boolean
    Dim zipHit As Boolean
    Dim jsHit As Boolean
    Dim tgzHit As Boolean
    Dim classFile As Boolean
    
    Dim thisNode As String
    Dim versionNode As String
    Dim tarNode As String
    Dim pathText As String

    Dim message1 As String
    message1 = "WARNING: 'version' column expected for output and." & vbCrLf
    message1 = message1 & "                  'Path' column expected for parsing." & vbCrLf
    message1 = message1 & "Do you want to continue?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

     Set myTable = ActiveSheet.ListObjects(1)
     Set myRow = myTable.ListRows(1)    ' init to row 1
     myRow.Range.Select

     ' setup column pointers based on olex-given column headings
     pathCol = myTable.ListColumns("Path").Index 'get column number for indexing
     versionCol = myTable.ListColumns("version").Index 'get column number for indexing

        rowCount = myTable.ListRows.Count
        checkCount = 0
        
 '       Debug.Print "Starting at : " & Now & " - rows to process: " & rowCount
     
     For rowIndex = 1 To rowCount Step 1  ' work each row loop
            jarHit = False
            versionHit = False
            dllHit = False
            zipHit = False
            jsHit = False
            tgzHit = False
            classFile = False
            
            'myTable.ListRows(rowIndex).Range.Select       ' get/select each row
            pathText = Cells(rowIndex + 1, pathCol).text       ' get the path field

            jarHit = ISLIKE(pathText, jarString)  ' got a .jar in the pathname?
            If jarHit = True Then GoTo doneHere

            dllHit = ISLIKE(pathText, dllString)  ' got a .dll file?
            If dllHit = True Then GoTo doneHere

            zipHit = ISLIKE(pathText, zipString)  ' got a .zip in the pathname?
            If zipHit = True Then GoTo doneHere
'
'skip .js check here because that only matters if versions numbers are also found
'           jsHit = ISLIKE(pathText, jsString)  ' got a .js file?
'            If jsHit = True Then GoTo doneHere
           
            versionHit = ISLIKE(pathText, versString0)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString00)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString10)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
         
            versionHit = ISLIKE(pathText, versString20)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString000)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString100)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
         
            versionHit = ISLIKE(pathText, versString200)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
                        
            versionHit = ISLIKE(pathText, versString1)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
         
            versionHit = ISLIKE(pathText, versString2)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString3)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString4)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString5)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString30)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString40)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString50)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
                        
            versionHit = ISLIKE(pathText, versString300)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString400)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
            
            versionHit = ISLIKE(pathText, versString500)  ' got a version like this?
            If versionHit = True Then GoTo doneHere
          
GoTo getnextRec     ' not a row with a version or jar
            
doneHere:
' arriving here = got something in path, so now go find which node
' we are ignoring .js files except those with version number
          versionNode = ""        ' initialize
           
        ' fall thru only to find the node(s) we want to save in version column
        nodeCount = dhCountTokens(pathText, "/")
' start node loop
        For nodeIndex = 1 To nodeCount Step 1           ' loop thru each node
           versionHit = False      'reset
           jarHit = False              'reset
           dllHit = False
           zipHit = False
           jsHit = False
           tgzHit = False
           classFile = False
                        
          thisNode = dhExtractString(pathText, nodeIndex, "/")
'
            jarHit = ISLIKE(thisNode, jarString)  ' got the jar node? take no matter what
            If jarHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
            End If
'         do same for dll's here if you want to see ALL DLL files

          zipHit = ISLIKE(thisNode, zipString)  ' got the zip node? take no matter what
            If zipHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
            End If
            
          tgzHit = ISLIKE(thisNode, tgzString)  ' got the tgz node? take no matter what
            If tgzHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
            End If
            
 ' fall thru only when not .jar,.zip, etc
            If nodeIndex = nodeCount Then   ' got last node?
              jsHit = ISLIKE(thisNode, jsString)   ' got a .js file?
              If jsHit = False Then GoTo getnextNode                      ' if last node, only .js files are kept
            End If
' look for version #'s next
              versionHit = ISLIKE(thisNode, versString0)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString1)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
               versionHit = ISLIKE(thisNode, versString00)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString000)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString100)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString200)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString10)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString20)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString2)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString3)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString4)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString5)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
               
              versionHit = ISLIKE(thisNode, versString30)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString40)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString50)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
                           
              versionHit = ISLIKE(thisNode, versString300)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString400)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
              versionHit = ISLIKE(thisNode, versString500)  ' got a version here?
              If versionHit = True Then
                versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
                GoTo getnextNode
              End If
      
getnextNode:
           Next nodeIndex

gotVersionField:

    If nodeIndex = nodeCount Then GoTo getnextRec 'got last node, = filename
        
' put newPath in its own column
     Cells(rowIndex + 1, versionCol).FormulaR1C1 = versionNode
getnextRec:
    Next rowIndex
'Debug.Print "Tar nodes fixed/dropped: " & checkCount
End Sub

Function ISLIKE(text As String, pattern As String) As Boolean
    'returns true if arg 1 is like arg2
    ISLIKE = text Like pattern
End Function

Sub CreateReportMAC_v3_1()
' createReport-v3
' last updates: Feb 2015
' must start with: open 6 sheet OLEX export (tabs: Files, Packages, Licenses, Conflicts, Obigations, Usage)
'
    Application.DisplayStatusBar = True  ' turn on status bar
'    Application.StatusBar = "Now creating report...."
    Application.ScreenUpdating = False  ' turn off screen updates while running
' set default table styles
    ActiveWorkbook.DefaultPivotTableStyle = "PivotStyleMedium9"
    ActiveWorkbook.DefaultTableStyle = "TableStyleLight9"
    
    Dim wb As Workbook
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
'=====================================
    'work on Licenses sheet
    Sheets("Licenses").Select  'go to License tab
    ' create table for files tab
    Set rng1 = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tblLicenses = ActiveSheet.ListObjects.Add(xlSrcRange, rng1, , xlYes)
    'tblLicenses.TableStyle = "TableStyleLight9"
    tblLicenses.Name = "tblLicenses"
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
    'tblObligations.TableStyle = "TableStyleLight9"
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
'
' create Obligatons pivot for Obligations Summary Sheet
'==============================================
    Sheets("Obligations").Select  'go to Obligations tab and insert before
    'add the Obligations pivot table sheet
    Sheets.Add.Name = "Pivot_ObligationSummary"
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

    '    .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
    End With

'    'set column withs for pivot
    Columns("A:A").ColumnWidth = 25
'    Columns("B:B").ColumnWidth = 10
    Range("A1").Select  'reset postion at top of sheet

'================================================
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
'    ActiveSheet.ListObjects("tblFiles").TableStyle = "TableStyleLight9"
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
    
    Sheets("Packages").Select
    Range("tblPackages[#All]").Select
    With Selection.Font
        .Size = 11
    End With

Range("A1").Select  'reset postion at top of sheet
    
'/////////////////////////////////////////////////////////////////////////
' go back to file sheet to insert new pivot
    Sheets("Files").Select  'go to Files tab
    Range("A1").Select  'reset postion at top of sheet
    'add the BOM pivot table sheet
    Sheets.Add.Name = "Pivot_BOMprep"
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
  '      .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
    End With
    'set column withs for pivot
    Columns("A:A").ColumnWidth = 50
    Columns("B:B").ColumnWidth = 15
    Range("A1").Select  'reset postion at top of sheet
'============================================
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
 '       .TableStyle2 = "PivotStyleMedium9"
        .ShowDrillIndicators = False    ' turn off drill arrows
'        .ShowPivotTableFieldList = False    ' turn off field list
    End With
    'set column withs for pivot
    Columns("D:D").ColumnWidth = 30
    Columns("E:E").ColumnWidth = 10
    Range("A1").Select  'reset postion at top of sheet
'\\\\\\\\\End of BOM pivot tables\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'/// prepare exit   //////////////////////////////////////////////////
    Application.StatusBar = False   ' reset from top
    ActiveWindow.TabRatio = 0.745   ' make wider the tab view along bottom
    Application.ScreenUpdating = True  ' turn back on screen
    MsgBox "Done"

End Sub

Sub MyMacro1()
    MsgBox "This macro 1."
End Sub

Sub MyMacro2()
    MsgBox "This macro 2."
End Sub

Sub MyMacro3()
    MsgBox "This macro 3."
End Sub

Sub MyMacro4()
    MsgBox "This macro 4."
End Sub

Sub MyMacro5()
    MsgBox "This macro 5."
End Sub

Sub MyMacro6()
    MsgBox "This macro 6."
End Sub

Sub MyMacro7()
    MsgBox "This macro 7."
End Sub

Sub MyMacro8()
    MsgBox "This macro 8."
End Sub

Sub MyMacro9()
    MsgBox "This macro 9."
End Sub

Sub MyMacro10()
    MsgBox "This macro 10."
End Sub

Sub MyMacro11()
    MsgBox "This macro 11."
End Sub

Sub MyMacro12()
    MsgBox "This macro 12."
End Sub

Sub MyMacro13()
    MsgBox "This macro 13."
End Sub

Sub MyMacro14()
    MsgBox "This macro 14."
End Sub
Option Explicit
Sub CreateReportMAC_v4_1()
' createReport4.1
' last updates: Mar-2015
' add 'Software Model' col to Files, add by-License BOM, and SoftwareModel Pivot
' must start with: open 6 sheet OLEX export (tabs: Files, Packages, Licenses, Conflicts, Obigations, Usage)
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
Sub ExtractFilename_v1()
'assumed in file is fossy export with Fullpath
'reads pathname column and writes to Filename column
' MUST HAVE columns: Filename, extionsion
'
    Dim thisSheet As Excel.Worksheet
    Dim myTable As ListObject
    Dim myRow As ListRow

    Dim sheetName As String
    Dim rowIndex As Long
    Dim checkCount As Long
    Dim nodeCount As Integer
    Dim nodeIndex As Integer
    Dim rowCount As Long
    
    Dim FilenameCol As Integer
    Dim FullPathCol As Integer
    
    Dim thisNode As String
    Dim FullPathText As String
    Dim newPath As String
    
    Dim message1 As String
    message1 = "WARNING: 'Filename' column expected for output and." & vbCrLf
    message1 = message1 & "                  'FullPath' column expected for parsing." & vbCrLf
    message1 = message1 & "Do you want to continue?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
    
    
     Set myTable = ActiveSheet.ListObjects(1)
     Set myRow = myTable.ListRows(1)    ' init to row 1
     myRow.Range.Select

     ' setup column pointers based on olex-given column headings
     FullPathCol = myTable.ListColumns("FullPath").Index 'get column number for indexing
     FilenameCol = myTable.ListColumns("Filename").Index 'get column number for indexing

    rowCount = myTable.ListRows.Count
    checkCount = 0
             
    For rowIndex = 1 To rowCount Step 1  ' work each row loop
        newPath = ""        'reset for each new row
        'myTable.ListRows(rowIndex).Range.Select       ' get/select each row
        FullPathText = Cells(rowIndex + 1, FullPathCol).text       ' get the FullPath field
        nodeCount = dhCountTokens(FullPathText, "/")    ' delimeter for filename start
        ' start node loop
        For nodeIndex = 1 To nodeCount Step 1           ' loop thru each node
            thisNode = dhExtractString(FullPathText, nodeIndex, "/")
            If nodeIndex = nodeCount Then   ' got last node?- then must be filename
                     Cells(rowIndex + 1, FilenameCol).FormulaR1C1 = thisNode
            Else
                    newPath = newPath & "/" & thisNode      'newpath = FillPath minus last node (filename is dropped)
            End If
        Next nodeIndex
        Cells(rowIndex + 1, FullPathCol).FormulaR1C1 = newPath      ' replace FullPath with Path
        
    Next rowIndex
End Sub
