Attribute VB_Name = "Module9"
Option Explicit
Sub PullVersionNode()
'
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
            
            'myTable.ListRows(rowIndex).Range.Select       ' get/select each row
            pathText = Cells(rowIndex + 1, pathCol).text       ' get the path field

'            jarHit = ISLIKE(pathText, jarString)  ' got a .jar in the pathname?
'            If jarHit = True Then GoTo doneHere
            
            dllHit = ISLIKE(pathText, dllString)  ' got a .dll file?
            If dllHit = True Then GoTo doneHere
            
'            zipHit = ISLIKE(pathText, zipString)  ' got a .zip in the pathname?
'            If zipHit = True Then GoTo doneHere
'
'            jsHit = ISLIKE(pathText, jsString)  ' got a .js file?
'            If jsHit = True Then GoTo doneHere
           
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
                
'              jsHit = ISLIKE(thisNode, jsString)  ' got the zip node?
'                If jsHit = True Then
'                    versionNode = thisNode & "   :   " & versionNode       ' move this node to the version node string
'                    GoTo gotVersionField
'                End If
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
Function ISLIKE(text As String, pattern As String) As Boolean
    'returns true if arg 1 is like arg2
    ISLIKE = text Like pattern
End Function

