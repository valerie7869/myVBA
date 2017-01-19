Attribute VB_Name = "Module1"
Option Explicit

Dim position As Long
Dim textPosition As Long
Dim theReportBook As Workbook   ' the report selected as input to convert
Dim HTMLname As String
Dim myFiles As String

' execShell() function courtesy of Robert Knight via StackOverflow
' http://stackoverflow.com/questions/6136798/vba-shell-function-in-office-2011-for-mac
Private Declare Function popen Lib "libc.dylib" (ByVal command As String, ByVal mode As String) As Long
Private Declare Function pclose Lib "libc.dylib" (ByVal file As Long) As Long
Private Declare Function fread Lib "libc.dylib" (ByVal outStr As String, ByVal size As Long, ByVal items As Long, ByVal stream As Long) As Long
Private Declare Function feof Lib "libc.dylib" (ByVal file As Long) As Long

Function bIsBookOpen(ByRef szBookName As String) As Boolean
    ' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function
Sub Select_File_PC_Mac2()
    'Select files in Mac Excel with the format that you want
    'Working in Mac Excel 2011 and 2016
    'Ron de Bruin, 20 March 2016
    Dim MyPath As String
    Dim MyScript As String
'    Dim MyFiles As String
    Dim MySplit As Variant
    Dim N As Long
    'Dim HTMLname As String
    Dim OneFile As Boolean
    Dim FileFormat As String

    'In this example you can only select xlsx files
    'See my webpage how to use other and more formats.
'    FileFormat = "{""org.openxmlformats.spreadsheetml.sheet"", ""public.html""}"
    FileFormat = "{""public.html""}"

    ' Set to True if you only want to be able to select one file
    ' And to False to be able to select one or more files
    OneFile = True

    On Error Resume Next
    MyPath = MacScript("return (path to downloads folder) as String")
    'Or use A full path with as separator the :
    'MyPath = "HarddriveName:Users:<UserName>:Desktop:YourFolder:"

    'Building the applescript string, do not change this
    If Val(Application.Version) < 15 Then
        'This is Mac Excel 2011
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return theFile"
    Else
        'This is Mac Excel 2016
        MsgBox ("Please run 32-bit Excel 2011")         ' for curl gets
        Exit Sub
    End If

    myFiles = MacScript(MyScript)
    On Error GoTo 0

    'If you select one or more files MyFiles is not empty
    'We can do things with the file paths now like I show you below
    If myFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With

        MySplit = Split(myFiles, Chr(10))
        For N = LBound(MySplit) To UBound(MySplit)

            'Get file name only and test if it is open
            HTMLname = Right(MySplit(N), Len(MySplit(N)) - InStrRev(MySplit(N), _
                Application.PathSeparator, , 1))

            If bIsBookOpen(HTMLname) = False Then

                Set theReportBook = Nothing
                On Error Resume Next
                'Set theReportBook = Workbooks.Open(MySplit(N))
                On Error GoTo 0

                If Not theReportBook Is Nothing Then
                    MsgBox "You selected this file : " & MySplit(N) & vbNewLine & _
                    " OK to continue?"
                    ' " OK to continue" & vbNewLine & _
                    ' "without saving, replace this line with your own code."
                    ' theReportBook.Close savechanges:=False
                End If
            Else
                MsgBox "We skip this file : " & MySplit(N) & " because it Is already open"
                Set theReportBook = Application.Workbooks(HTMLname)
            End If

            Next N
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With
    End If
    
    ' workbook to use as wb.report is theReportBook.name
End Sub

Sub Select_File_PC_Mac()
    'Select files in Mac Excel with the format that you want
    'Working in Mac Excel 2011 and 2016
    'Ron de Bruin, 20 March 2016
    Dim MyPath As String
    Dim MyScript As String
    Dim myFiles As String
    Dim MySplit As Variant
    Dim N As Long
    Dim Fname As String
    Dim OneFile As Boolean
    Dim FileFormat As String

    'In this example you can only select xlsx files
    'See my webpage how to use other and more formats.
    FileFormat = "{""org.openxmlformats.spreadsheetml.sheet"", ""public.html""}"
'    FileFormat = "{""public.html""}"

    ' Set to True if you only want to be able to select one file
    ' And to False to be able to select one or more files
    OneFile = True

    On Error Resume Next
    MyPath = MacScript("return (path to downloads folder) as String")
    'Or use A full path with as separator the :
    'MyPath = "HarddriveName:Users:<UserName>:Desktop:YourFolder:"

    'Building the applescript string, do not change this
    If Val(Application.Version) < 15 Then
        'This is Mac Excel 2011
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return theFile"
    Else
        'This is Mac Excel 2016
        MsgBox ("Please run 32-bit Excel 2011")         ' for curl gets
        Exit Sub
    End If

    myFiles = MacScript(MyScript)
    On Error GoTo 0

    'If you select one or more files MyFiles is not empty
    'We can do things with the file paths now like I show you below
    If myFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With

        MySplit = Split(myFiles, Chr(10))
        For N = LBound(MySplit) To UBound(MySplit)

            'Get file name only and test if it is open
            Fname = Right(MySplit(N), Len(MySplit(N)) - InStrRev(MySplit(N), _
                Application.PathSeparator, , 1))

            If bIsBookOpen(Fname) = False Then

                Set theReportBook = Nothing
                On Error Resume Next
                Set theReportBook = Workbooks.Open(MySplit(N), ReadOnly:=True)
                On Error GoTo 0

                If Not theReportBook Is Nothing Then
                    MsgBox "You selected this file : " & MySplit(N) & vbNewLine & _
                    " OK to continue?"
                    ' " OK to continue" & vbNewLine & _
                    ' "without saving, replace this line with your own code."
                    ' theReportBook.Close savechanges:=False
                End If
            Else
                MsgBox "Cannot re-open: " & MySplit(N) & " - Close it and try again. "
                'Set theReportBook = Application.Workbooks(Fname)
            End If

            Next N
        With Application
            .ScreenUpdating = True
            .EnableEvents = True
        End With
    End If
    
    ' workbook to use as wb.report is theReportBook.name
End Sub


Sub createCVE_JSONtab()
Dim wsJSON As Worksheet
Dim lookupRange, tempRange, rng1, cell As Range
Dim tblTemp As ListObject
Dim lookString, packageID, packageName As String
Dim ws As Worksheet
Dim lastRow As Long
Dim message1 As String
Dim theRange As Range

message1 = message1 & "Create/Replace the json_cve tab and CVE Details Report. YES ot Continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

Sheets("GET_CVEs").Range("H8").Value = Now      ' set start time
'clear any old results before starting
Application.DisplayAlerts = False  ' do not ask to confirm deletes
On Error Resume Next        ' if already there, delete to rebuild
    ActiveWorkbook.Sheets("json_cve").Delete
Set wsJSON = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))

With wsJSON
'    .Select
    .Name = "json_cve"
    .Tab.Color = 65535  ' yellow tab
End With
    
On Error GoTo 0
'
Sheets("GET_CVEs").Select
Range("E11").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
'tempRange.Copy
'Sheets("json_cve").Select
Sheets("json_cve").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Application.CutCopyMode = False
Sheets("GET_CVEs").Range("A1").Select   'landing
Sheets("json_cve").Select
Sheets("json_cve").Range("A1").Select   'landing
' convert to table and sort on package name
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Set theRange = Selection
    ActiveSheet.ListObjects.Add(xlSrcRange, theRange, , xlYes).Name = "tblCVEDetail"
    ActiveSheet.ListObjects(1).TableStyle = "TableStyleLight9"
    
' sort on packageName
ActiveWorkbook.Worksheets("json_cve").ListObjects("tblCVEDetail").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("json_cve").ListObjects("tblCVEDetail").Sort. _
        SortFields.Add Key:=Range("tblCVEDetail[[#All],[PackageName]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("json_cve").ListObjects("tblCVEDetail").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
' conditional format the severity column
    Range("tblCVEDetail[Severity]").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="HIGH", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .Font.Color = RGB(156, 20, 29)
        .Interior.Color = RGB(254, 199, 206)
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="LOW", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .Font.Color = RGB(25, 96, 22)   'dark green
        .Interior.Color = RGB(199, 238, 206)    ' lite green
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="MEDIUM", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .Font.Color = RGB(155, 106, 59)         ' dark yellow
        .Interior.Color = RGB(254, 234, 160)        ' lite yellow
    End With
    Selection.FormatConditions(1).StopIfTrue = False
Selection.Font.Bold = True

' conditional format the AccessComplexity column
    Range("tblCVEDetail[AccessComplexity]").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="LOW", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .Font.Bold = True
    '    .Font.Color = RGB(223, 10, 23)  'Red
    '    .Interior.Color = RGB(254, 199, 206)
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="HIGH", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
     '   .Font.Color = RGB(25, 96, 22)   'dark green
     '   .Interior.Color = RGB(199, 238, 206)    ' lite green
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="MEDIUM", TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
    '    .Font.Color = RGB(243, 142, 73)      'Orange
    '    .Interior.Color = RGB(254, 234, 160)        ' lite yellow
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'Selection.Font.Bold = True

Range("A1").Select

Call SaveCVE_Report ' copy and save tab as CVEReport

Application.ScreenUpdating = True
End Sub

Sub getCVEs()
Dim wsJSON, wsGETCVE As Worksheet
Dim myRequest, myRequest2 As Object
'Dim theResponse As New DOMDocument
'Dim theResponse2 As New DOMDocument
Dim theTag, theTag2, theResponse, theContent, theResponseString As String
Dim cell, tempRange As Range
'Dim myNode, myNode2 As IXMLDOMNode
Dim myNode, myNode2 As Range
'Dim packageList, packageList2, securityList, versionList As IXMLDOMNodeList
'Dim securityList As IXMLDOMNodeList'Dim cveList As IXMLDOMNodeList
'dim cveList, securityList, packageList, versionList a
Dim severity, versionList As String
'Dim xmlAttribute As IXMLDOMAttribute
Dim pathStr, pathStr2 As String
Dim pathLen, tempLen, packCount, cveCount, itemCount, itemCount2 As Integer
Dim url, login_id, pass, loas, packageID, packageName As String
Dim riskID, tempString, versionStr, countText, countText2 As String
Dim packIDList As Range
Dim i, x, y, col, row As Integer
Dim riskNum, message1 As String
Dim tempPosition As Long

message1 = message1 & "Call NEW CVE list for loas. YES ot Continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

Debug.Print "Starting AT : "; Now
Application.ScreenUpdating = False

Set wsGETCVE = ActiveSheet
wsGETCVE.Range("F1").Value = Now
'clear all old results:
wsGETCVE.Range("A12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
wsGETCVE.Range("B12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
For col = 0 To 13
    wsGETCVE.Range("E12").Offset(0, col).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.Style = "Normal"
    col = col
Next col
wsGETCVE.Range("loas").Select    'land here
DoEvents

login_id = wsGETCVE.Range("loginID").Value
pass = wsGETCVE.Range("password").Value
loas = wsGETCVE.Range("loas")
''''''''''
''jan5
    Dim sCmd As String
    Dim lExitCode As Long
    Dim xml2 As String
 ' go get the loas, used to find the packageID and packageNames
    ' example url : https:  and  //audit dot openlogic dot com/loas/42411/
    url = "https://audit.openlogic.com/loas/" & loas
    'sCmd = "curl file:///Users/valerie/licenses-sample.html"
    sCmd = "curl -X GET -u """ & login_id & ":" & pass & """ -H ""Accept: application/vnd.openlogic.olexgovernance+xml""   -H ""Content-Type: application/vnd.openlogic.olexgovernance+xml"" " & url
' pull the trigger > get the loas
theResponseString = execShell(sCmd, lExitCode)

theContent = theResponseString
' get basic loas info from the response
theTag = "version_name"
tempString = getTag1(theTag, theResponseString)
wsGETCVE.Range("scan_version_name").Value = tempString

theTag = "created_at"
tempString = getTag1(theTag, theResponseString)
wsGETCVE.Range("created_at").Value = tempString

theTag = "updated_at"
tempString = getTag1(theTag, theResponseString)
wsGETCVE.Range("updated_at").Value = tempString

' now get package count
position = 0        'init
countText = "package id="     ' count of package's
packCount = (Len(theContent) - Len(Replace(theContent, countText, ""))) / Len(countText)
position = InStr(200, theContent, "<packages>") ' initialize position in reply

Debug.Print "Start search position: " & position
' jan8 - todo - make custom routine to create the securitylist from:  cveCount = how many
'''''''''''   sample: <security_risk id="https://audit.openlogic.com/security_risks/44518"/>
theTag = "package id="
theTag2 = "<title>"
Debug.Print "listing packages found: "
row = 0
For row = 0 To packCount - 1
    pathStr = getTagSetPosition2(theTag, theResponseString, position)
    pathLen = Len(pathStr)
    packageID = Right(pathStr, (pathLen - 38))
    ' 38 = fixed length of path chars to drop= ""https://audit.openlogic.com/packages/"
    pathStr2 = getTagSetPosition2(theTag2, theResponseString, position)
    pathLen = Len(pathStr2)
    packageName = Left(pathStr2, (pathLen - 6))
    Debug.Print packageID & " -  " & packageName
    wsGETCVE.Range("A12").Offset(row, 0).Value = packageID
    wsGETCVE.Range("A12").Offset(row, 1).Value = packageName
    ' = row + 1
Next row
    
' now go get the security risk list
    'riskID = "itext"
    'packageID = cell.Value2
    'url = cell.Offset(0, 6).Value2  ' this is home page URL
    ' https://audit.openlogic.com/loas/42411/security_risks
    url = "https://audit.openlogic.com/loas/" & loas & "/security_risks"
    'sCmd = "curl file:///Users/valerie/licenses-sample.html"
    sCmd = "curl -X GET -u """ & login_id & ":" & pass & """ -H ""Accept: application/vnd.openlogic.olexgovernance+xml""   -H ""Content-Type: application/vnd.openlogic.olexgovernance+xml"" " & url
' pull the trigger >
    theResponseString = execShell(sCmd, lExitCode)
theContent = theResponseString
' get cve count

countText = "risk id="     ' count of cve's
cveCount = (Len(theContent) - Len(Replace(theContent, countText, ""))) / Len(countText)
position = InStr(170, theContent, "security_risk id") ' initialize position in reply
' list to get_cve tab the cve's

' jan8 - todo - make custom routine to create the securitylist from:  cveCount = how many
'''''''''''   sample: <security_risk id="https://audit.openlogic.com/security_risks/44518"/>
theTag = "security_risk id="
Debug.Print "listing security risk URLs (CVEs) found: "
row = 0
For row = 0 To cveCount - 1
    pathStr = getTagSetPosition2(theTag, theResponseString, position)
    ' drop quotes on bothleft and right ends of path
    pathLen = Len(pathStr)
    pathStr = Left(pathStr, pathLen - 1)
    pathLen = Len(pathStr)  ' new length
    pathStr = Right(pathStr, pathLen - 1)
    pathLen = Len(pathStr)  ' new length
    Debug.Print pathStr
    wsGETCVE.Range("F12").Offset(row, 0).Value = pathStr
Next row
 
''''''''''' create the named range for packageID's
''''''''''wsgetcve.Range("A12").Select
''''''''''Range(Selection, Selection.End(xlDown)).Select
''''''''''Set tempRange = Selection
''''''''''tempRange.Name = "packIDList"   ' name for later use

wsGETCVE.Select
wsGETCVE.Range("securityRiskCount").Value = "#risks:    " & cveCount
wsGETCVE.Range("packageCount").Value = "#total packages:    " & packCount
wsGETCVE.Range("loas").Select    'land here
Debug.Print "1st half COMPLETE AT : "; Now
Application.ScreenUpdating = True       ' allow user to see progress

DoEvents
    message1 = " Found " & cveCount & " CVEs. Fetch details now?" & vbCrLf
    message1 = message1 & "  note: expect up to 1 minute per 30 CVE's" & vbCrLf
    message1 = message1 & "Do you want to continue now?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
    
'Application.ScreenUpdating = False         ' leave true so user can see progress
versionList = ""    ' init
' call API for each CVE
wsGETCVE.Select
For row = 0 To cveCount - 1
' example url to get CVE = https://audit.openlogic.com/security_risks/44518
    url = wsGETCVE.Range("E12").Offset(row, 1).Text
    
    wsGETCVE.Range("E12").Offset(row, 1).Select
    DoEvents
    sCmd = "curl -X GET -u """ & login_id & ":" & pass & """ -H ""Accept: application/vnd.openlogic.olexgovernance+xml""   -H ""Content-Type: application/vnd.openlogic.olexgovernance+xml"" " & url
    ' pull the trigger > get a CVE
    theResponseString = execShell(sCmd, lExitCode)
    theContent = theResponseString

    theTag = "cve_number>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 3).Value = tempString

    theTag = "score>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 4).Value = tempString
' get severity
    If tempString < 4 Then
        severity = "LOW"
        wsGETCVE.Range("F12").Offset(row, 5).Font.Color = RGB(15, 127, 18)      ' dk green
        wsGETCVE.Range("F12").Offset(row, 5).Interior.Color = RGB(205, 254, 205)    ' lite green
    Else
        If tempString < 7 Then
            severity = "MEDIUM"
        wsGETCVE.Range("F12").Offset(row, 5).Font.Color = RGB(155, 106, 59)     ' dark yellow
        wsGETCVE.Range("F12").Offset(row, 5).Interior.Color = RGB(254, 234, 160)    ' lite yellow
        
        Else
            severity = "HIGH"
            wsGETCVE.Range("F12").Offset(row, 5).Font.Color = RGB(156, 20, 29)      ' dk red
            wsGETCVE.Range("F12").Offset(row, 5).Interior.Color = RGB(254, 199, 206)  'pink
        End If
    End If
    wsGETCVE.Range("F12").Offset(row, 5).Value = severity
    wsGETCVE.Range("F12").Offset(row, 5).Font.Bold = True
    
    theTag = "access_complexity>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 6).Value = tempString

    theTag = "updated_at>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 7).Value = tempString

    theTag = "nvd_link>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 8).Value = tempString

    theTag = "created_at>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 9).Value = tempString

    theTag = "summary>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 10).Value = tempString

    theTag = "access_vector>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 11).Value = tempString

    theTag = "integrity_impact>"
    tempString = getTag1(theTag, theResponseString)
    wsGETCVE.Range("F12").Offset(row, 12).Value = tempString


countText = "<package id="     ' count this
itemCount = (Len(theContent) - Len(Replace(theContent, countText, ""))) / Len(countText)
countText2 = "version id="     ' count this to get versions count
itemCount2 = (Len(theContent) - Len(Replace(theContent, countText2, ""))) / Len(countText2)
If itemCount2 = 0 Then versionList = "No Versions LIsted"
position = 0
Debug.Print "Position= " & position
position = InStr(200, theContent, "<packages>") ' initialize position in XML response
'get all packages for this CVE
theTag = "package id="
' for each package listed for this CVE, find the one matching this loas
    For x = 1 To itemCount
    Debug.Print "Position= " & position
            pathStr = getTagSetPosition2(theTag, theResponseString, position)   ' this will advance 'postion' (as a global variable)
            ' jan9 resume here --- get the package(s) for this CVE (should I get all, or the 1st?)
            pathLen = Len(pathStr)
            packageID = Right(pathStr, (pathLen - 38))      ' trim right 38 characters (first part of URL) leaving only packageID
            pathLen = Len(packageID)        ' new length to trim
            packageID = Left(packageID, (pathLen - 1))    ' trim off 1 char
    
            ' 37 = fixed length of path chars to drop= "https://audit.openlogic.com/packages/"
            packageID = packageID
            Range("A12").Select
            Range(Selection, Selection.End(xlDown)).Select
            Set packIDList = Selection
            ' note: this may not match the named range in the spreadsheet, but thats ok
            Set cell = Range("packIDList").Find(packageID, , xlValues, xlWhole)
            If Not cell Is Nothing Then 'found
                packageName = cell.Offset(0, 1).Value2
                wsGETCVE.Range("F12").Offset(row, -1).Value = packageID
                wsGETCVE.Range("F12").Offset(row, 0).Value = packageName
                GoTo doneHere
            Else    'not found
                packageName = ""
            End If
        
    Next x  ' next package listed for CVE
'''''
doneHere:   ' got a matching package in the CVE to this loas so get the versions
tempPosition = InStr(position, theContent, "<package_versions>") ' initialize position in XML response
'get all packages for this CVE
theTag2 = "version id="
' for each package listed for this CVE, find the one matching this loas
    For y = 1 To itemCount2
            Debug.Print "versions Position= " & tempPosition
            pathStr = getTagSetPosition2(theTag2, theResponseString, tempPosition)   ' this will advance 'postion' (as a global variable)
            ' jan9 resume here --- get the package(s) for this CVE (should I get all, or the 1st?)
            pathLen = Len(pathStr)
            versionStr = Right(pathStr, (pathLen - 38))      ' trim right 38 characters (first part of URL) leaving only packageID & version
            pathLen = Len(versionStr)        ' new length to trim
            versionStr = Left(versionStr, (pathLen - 1))    ' trim off 1 char
            versionList = versionList & versionStr
            If y < itemCount2 Then versionList = versionList & ", "  ' cuz there are more
    Next y  ' next package listed for CVE
wsGETCVE.Range("F12").Offset(row, 2).Value = versionList        ' 2 cols over
versionList = ""      ' init

Debug.Print (row & " " & packageID)
Next row
'.......................................................................................................
wsGETCVE.Select
wsGETCVE.Range("securityRiskCount").Value = "#risks:    " & cveCount
' wrap versions and top justify
'Columns("G:G").WrapText = True
'Rows("12:6000").VerticalAlignment = xlTop

wsGETCVE.Range("loas").Select    'land here
Debug.Print "COMPLETE AT : "; Now
Application.ScreenUpdating = True
ActiveSheet.Range("A12").Select 'landing
'ActiveSheet.Range("loas").Select    'land here
MsgBox "Complete"
End Sub

Sub NewJSONTabs2()
Dim ws As Worksheet
Dim wbReport, wbMacro As Workbook
Dim createFiles As Boolean
'Dim wbMacro, wbReport As Workbook       ' both workbooks will be open
Dim r, c, pCol, lookup_col As Integer
Dim x, lastRow, xRow, xCol, xCount As Long
Dim packageListCol, cell, tempRange, theTable As Range
Dim url, tempString, message1 As String
Dim debug1 As Boolean
Dim reportIn As String

createFiles = False
message1 = " CONVERT .xlsx Report to JSONtabs." & vbCrLf
message1 = message1 & " Please next select a completed OLEX Report (Excel)" & vbCrLf
message1 = message1 & "Ready to continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

Application.DisplayAlerts = False  ' do not ask to confirm deletes
Application.ScreenUpdating = False
'DoEvents
Select_File_PC_Mac
Set wbReport = theReportBook

'Set wbReport = ActiveWorkbook
Set wbMacro = ThisWorkbook
' activeworkbook will be the report, while Thisworkbook will be the macro and results
On Error Resume Next        ' if already there, delete to rebuild
    wbMacro.Sheets("json_summary").Delete
    wbMacro.Sheets("json_bom").Delete
    wbMacro.Sheets("json_softwaremodel").Delete
    wbMacro.Sheets("json_bylicense").Delete
    wbMacro.Sheets("json_obligations").Delete
    wbMacro.Sheets("Packages").Delete
    wbMacro.Sheets("Licenses").Delete
On Error GoTo 0
Application.DisplayAlerts = True  ' turn back on
wbReport.Sheets("Packages").Copy After:=wbMacro.Sheets(wbMacro.Sheets.Count)
wbReport.Sheets("Licenses").Copy After:=wbMacro.Sheets(wbMacro.Sheets.Count)

' -jan17 - change columns and sort licenses sheet
    Sheets("Licenses").Select
    Columns("B:B").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("tblLicenses[[#Headers],[Name]]").Select
    ActiveWorkbook.Worksheets("Licenses").ListObjects("tblLicenses").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Licenses").ListObjects("tblLicenses").Sort. _
        SortFields.Add Key:=Range("tblLicenses[[#All],[Name]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Licenses").ListObjects("tblLicenses").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("tblLicenses[[#Headers],[Name]]").Select

'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'++++++++++add json_summary tab ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Set ws = wbMacro.Sheets.Add(After:=wbMacro.Sheets(wbMacro.Sheets.Count))
    ws.Name = "json_summary"
'    ThisWorkbook.Sheets("json_summary").Select
With wbMacro.Sheets("json_summary")
    .Tab.Color = 65535  ' yellow tab
    .Range("A1") = "Client"
    .Range("A2") = "Project name"
    .Range("A3") = "Filenames"
    .Range("A4") = "Source of project files"
    .Range("A5") = "SHA1 Checksum"
    .Range("A6") = "Number of files scanned"
    .Range("A7") = "Number of Open Source Package Identified"
    .Range("A8") = "Number of Open Source Licenses Identified"
    .Range("A9") = "ReportDate"             'jan16
    .Range("A10") = "ReportIntro"             'jan16
    .Range("A11") = "Disclaim1"
    .Range("A12") = "Disclaim2"
    .Range("A13") = "HelpBOM"
    .Range("A14") = "HelpCVE"
    
'    Range("B1") = Sheets("Summary").Range("Client Name").Text
    .Range("B1") = "Client Name"
    ' below is using named ranges from spreadsheet
    .Range("B2") = wbReport.Sheets("Summary").Range("projectName").Value2
    .Range("B3") = wbReport.Sheets("Summary").Range("projectFilenames").Value2
    .Range("B4") = wbReport.Sheets("Summary").Range("projectSourceFiles").Value2
    .Range("B5") = wbReport.Sheets("Summary").Range("projectChecksum").Value2
'    .Range("B5") = "SHA1 Checksum"
    .Range("B6") = wbReport.Sheets("Summary").Range("Summary.FileCount").Value2
    .Range("B7") = wbReport.Sheets("Summary").Range("Summary.PackageCount").Value2
    .Range("B8") = wbReport.Sheets("Summary").Range("Summary.LicenseCount").Value2
    .Range("B9") = wbReport.Sheets("Summary").Range("Summary.ReportDate").Value2       'jan16
    .Range("B10") = wbReport.Sheets("Summary").Range("Summary.ReportIntro").Value2       'jan16
    .Range("B11") = wbReport.Sheets("Summary").Range("Summary.ReportDisclaim1").Value
    .Range("B12") = wbReport.Sheets("Summary").Range("Disclaim2").Value
    .Range("B13") = wbReport.Sheets("Summary").Range("HelpBOM").Value
    .Range("B14") = wbReport.Sheets("Summary").Range("HelpCVE").Value
    
'    .Range("A1").Select  'landing
'++ end of summary json +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End With
'++++++++ add json_bom +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Set ws = wbMacro.Sheets.Add(After:=wbMacro.Sheets(wbMacro.Sheets.Count))
ws.Name = "json_bom"
    pCol = wbReport.Sheets("BOM (Bill of Material)").Range("tblBOM[choiceType]").Column   ' get col number for later deletion in json
    wbReport.Sheets("BOM (Bill of Material)").Range("tblBOM[#All]").Copy
    With wbMacro.Sheets("json_bom")
        .Tab.Color = 65535  ' yellow tab
        .Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Application.CutCopyMode = False
        .Columns(pCol).EntireColumn.Delete   ' delete from json the choiceType column
    Application.CutCopyMode = False
        .Range("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.CutCopyMode = False
        .Range("A1").FormulaR1C1 = "PackageID"
        .Range("B1").FormulaR1C1 = "LicenseID"
'+++++++jan2
        .Range("A1:J1").Replace What:=" ", Replacement:=""
 '       .Rows(1).Replace What:=" ", Replacement:=""
'            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
            False, ReplaceFormat:=False
    
        .Rows(1).Replace What:="(s)", Replacement:="s"
'        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
'++++++++jan2

 End With

'    Range("A1").Select
'now do vlookups using Packages Tab to add packageID and licenseID to above table
' lookup packageid and licenseid
    
    lookup_col = wbMacro.Sheets("Packages").Range("tblpackages[ID]").Column
  '  .Range("A2").FormulaR1C1 = "=VLOOKUP(RC[2]," & Range("temprange") & lookup_col & ",FALSE)"
    
    Range("A2").FormulaR1C1 = "=VLOOKUP(RC[2],tblPackages," & lookup_col & ",FALSE)"
    ' copy formula all the way down
    lastRow = wbMacro.Sheets("json_bom").Cells(Rows.Count, 3).End(xlUp).row
    Range("A2:A" & lastRow).FillDown
    DoEvents
    Range("A2:A" & lastRow).Value = Range("A2:A" & lastRow).Value   ' paste values over vlookup formula
    
'    Range("B2").Select
    lookup_col = wbMacro.Sheets("Licenses").Range("tblLicenses[ID]").Column
    Range("B2").FormulaR1C1 = "=VLOOKUP(RC[4],tblLicenses," & lookup_col & ",FALSE)"
    ' copy formula all the way down
    Range("B2:B" & lastRow).FillDown
    DoEvents
    Range("B2:B" & lastRow).Value = Range("B2:B" & lastRow).Value   ' paste values over vlookup formula
    Range("A1").Select
'===================================
'    wbmacro.Sheets("json_bom").range("A1").activate
    
'''' convert theTable to JSON for BOM
'    Debug.Print "Convert BOM to json"

'++++++++++++++++++++++ add json_obligations++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Set ws = wbMacro.Sheets.Add(After:=wbMacro.Sheets(wbMacro.Sheets.Count))
    ws.Name = "json_obligations"
    
    wbReport.Sheets("Obligations Details").Range("tblObligations[#All]").Copy
'    wbreport.Sheets("Obligations Details").Select
'    Range("tblObligations[#All]").Select
'    Selection.Copy
    With wbMacro.Sheets("json_obligations")
        .Tab.Color = 65535  ' yellow tab
        .Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Application.CutCopyMode = False
    wbMacro.Sheets("json_obligations").Range("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove  ' for LicenseId
    Application.CutCopyMode = False
    wbMacro.Sheets("json_obligations").Range("A1").FormulaR1C1 = "LicenseID"
    'Range("A1").Select
    End With
    With wbMacro.Sheets("json_obligations").Rows(1)     ' fix headers expected for json
        .Replace What:=" ", Replacement:=""
 '           LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
            False, ReplaceFormat:=False
        .Replace What:="(s)", Replacement:="s"
  '          LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
            False, ReplaceFormat:=False
    End With
'now do vlookups to add licenseID to above table
'    Sheets("json_obligations").Select
 '*****************************dec27

'    Range("A2").Select
    lastRow = wbMacro.Sheets("json_obligations").Cells(Rows.Count, 3).End(xlUp).row
    lookup_col = wbReport.Sheets("Licenses").Range("tblLicenses[ID]").Column
    wbMacro.Sheets("json_obligations").Range("A2").FormulaR1C1 = "=VLOOKUP(RC[2],tblLicenses," & lookup_col & ",FALSE)"
    ' copy formula all the way down
    Range("A2:A" & lastRow).FillDown
    DoEvents
    Range("A2:A" & lastRow).Value = Range("A2:A" & lastRow).Value   ' paste values over vlookup formula
    Range("A1").Select      ' landing for json tab
    
    'Sheets("Obligations Details").Select
'''' convert theTable to JSON for Obligations
    Debug.Print "Convert Obligations to json"
    

'+++++++++ done adding json_obligations  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'++++++ add json_software model and by license  +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Set ws = wbMacro.Sheets.Add(After:=wbMacro.Sheets(wbMacro.Sheets.Count))
    ws.Name = "json_softwaremodel"

'    tempRange = wbReport.Sheets("Pivot_UserRefresh").Range("D4")  ' point to software model table
   Set tempRange = Range(wbReport.Sheets("Pivot_UserRefresh").Range("D4"), wbReport.Sheets("Pivot_UserRefresh").Range("D4").End(xlToRight))
   Set tempRange = Range(tempRange, tempRange.End(xlDown))
    tempRange.Copy
    With wbMacro.Sheets("json_softwaremodel")
        .Tab.Color = 65535  ' yellow tab
        .Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    End With
    wbMacro.Sheets("json_softwaremodel").Range("B1") = "Files"
    wbMacro.Sheets("json_softwaremodel").Range("A1") = "SoftwareModel"
'''' convert theTable to JSON for softwaremodel
    Application.CutCopyMode = False
    Range("A1").Select
    Debug.Print "Convert softwareModel to json"
    
    
' now get license type json tab
    Set ws = wbMacro.Sheets.Add(After:=wbMacro.Sheets(wbMacro.Sheets.Count))
    ws.Name = "json_byLicense"
   Set tempRange = Range(wbReport.Sheets("Pivot_UserRefresh").Range("I4"), wbReport.Sheets("Pivot_UserRefresh").Range("I4").End(xlToRight))
   Set tempRange = Range(tempRange, tempRange.End(xlDown))
    tempRange.Copy
    With wbMacro.Sheets("json_bylicense")
        .Tab.Color = 65535  ' yellow tab
        .Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    End With
    wbMacro.Sheets("json_bylicense").Range("B1") = "Files"
    wbMacro.Sheets("json_bylicense").Range("A1") = "LicenseType"
'''' convert theTable to JSON for byLicense
    Application.CutCopyMode = False
    Range("A1").Select
    Debug.Print "Convert byLicense to json"
    
'    Selection = Nothing
    wbReport.Close SaveChanges:=False
    
Sheets("creds").Select
Application.ScreenUpdating = True
MsgBox ("Completed json tabs")
End Sub

Sub createJSONLicenses()
Dim message1 As String
Dim weirdStr As String
Dim copyRightStr As String
Dim createFiles As Boolean

createFiles = False

Dim wbMacro As Workbook
Set wbMacro = ThisWorkbook

Dim ws As Worksheet

' do licenses to json here
' check if prepared:
On Error GoTo abort1
Sheets("Licenses").Select   ' abort if tab not found
On Error GoTo 0     ' reset

message1 = "REQUIRED: licenses.html export " & vbCrLf
message1 = message1 & "  Select the licenes.html next." & vbCrLf
message1 = message1 & "Do you want to continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
Application.ScreenUpdating = False
Application.DisplayAlerts = False  ' do not ask to confirm deletes
On Error Resume Next        ' if already there, delete to rebuild
    wbMacro.Sheets("json_licenseinfo").Delete
On Error GoTo 0
'Application.DisplayAlerts = True  ' turn back on

Dim licenseListCol As Range
Dim licenseID, licenseIDx, licenseName1, licenseNameHTML, licenseTax, licenseText As String
Dim theTag, cleanText2, cleanText1, tempText As String

Dim myRequest As Object
Dim cell, row, col As Range
'Dim url As String
Dim debug1 As Boolean
Dim fso2 As Object     ' for output text files
Dim fileOut As Object
Dim fileOutLic As Object
Dim filenameIn As String
Dim textEnd, tempLen, diff As Long
Dim i, itemCount, pCount, lCount As Long

Select_File_PC_Mac2
'filenameIn = Application.GetOpenFilename()
filenameIn = myFiles
Application.ScreenUpdating = False

'filenameIn = "Macintosh HD:Users:valerie:Downloads:licenses.html"
Open filenameIn For Input As #1
Dim text2, textline, countText, theContent As String
'Dim textPosition As Long
Do Until EOF(1)
    Input #1, textline
    text2 = text2 & textline
Loop

Close #1
theContent = text2
' get license count

countText = "<section>"     ' count this
itemCount = (Len(theContent) - Len(Replace(theContent, countText, ""))) / Len(countText)

position = InStr(200, theContent, "<section>") ' initialize position in HTML file
If createFiles = True Then
' parse text1
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    Set fileOutLic = fso2.createtextfile(wbMacro.Path & "\json_licenseinfo.js", True, True)
    fileOutLic.write "var json_licenseinfo = [ {" & vbCrLf ' one-time header
End If
'set up loop here to get only licenses from olex export license tab - check for match on name? - then
''..then pass the id and the taxonomy from the spreadsheet

' for each cell in licenseList
'On Error GoTo abort1
Sheets("Licenses").Select
'On Error GoTo 0     ' reset
' sort the liceneses
    wbMacro.Worksheets("Licenses").ListObjects("tblLicenses").Sort. _
        SortFields.Clear
    wbMacro.Worksheets("Licenses").ListObjects("tblLicenses").Sort. _
        SortFields.Add Key:=Range("tblLicenses[[#All],[Name]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wbMacro.Worksheets("Licenses").ListObjects("tblLicenses").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select

Set ws = wbMacro.Sheets.Add(After:=wbMacro.Sheets(wbMacro.Sheets.Count))
ws.Name = "json_licenseinfo"
wbMacro.Sheets("json_licenseinfo").Select
With wbMacro.Sheets("json_licenseinfo")
    .Tab.Color = 65535  ' yellow tab
    ' headings
    .Range("A1") = "license_id"
    .Range("B1") = "license_name"
    .Range("C1") = "taxonomy"
    .Range("D1") = "text"
    .Range("D:D").ColumnWidth = 65.1
    .Range("D:D").WrapText = False
    .Range("A1").Select
End With
DoEvents
Set licenseListCol = wbMacro.Sheets("Licenses").Range("tblLicenses[Name]") ' loop thru each license name
'xCount = licenseListCol.Count
'xRow = 0
'xCol = 0
'Do Until xRow > (xCount + 1)
lCount = 0
row = 1 ' output grid

'********************************************************************************************************************
' double loop follows : 1st: loop thru each license listed in cells, and for each, then loop thru each license found in HTML export to locate text
'**********************************************************************************************************************
    For Each cell In licenseListCol
                    'cell.Activate   ' so user can see progress
                    lCount = lCount + 1
                    licenseName1 = cell.Text
                    licenseID = cell.Offset(0, 1).Text  ' get ID from next col on this row
                    licenseTax = cell.Offset(0, 2).Text ' ...and Taxonomy is 2 cols over
                    
                    Debug.Print "seeking: " & licenseID
                    
                    theTag = "<h1><li>"     ' tag used to find license name in HTML
                    For i = 1 To itemCount
                    i = i   'debug
                        licenseNameHTML = getLicNameTag2(theTag, theContent, position)
                        ' if not the same, loop now:
                        If licenseNameHTML <> licenseName1 Then GoTo loopAgain  'not it,so get next since 'position' and 'textPosition' are updated
                    '   else, got a match to license text, so write the json >>
                        textEnd = InStr(textPosition, theContent, "</section>")
                        tempLen = textEnd - textPosition
                        licenseText = Mid(theContent, textPosition, tempLen)
                        textPosition = textPosition + tempLen + 1
                        On Error Resume Next
                        cleanText1 = Application.WorksheetFunction.Clean(licenseText)
                        cleanText1 = licenseText
                    
                        ' jan13 --replace weird quotes with standard quote:
'                        cleanText1 = Replace(cleanText1, Chr("&H9D"), Chr("&H22"))
'                        cleanText1 = Replace(cleanText1, Chr("&H9C"), Chr("&H22"))
'                        cleanText1 = Replace(cleanText1, Chr("&H84"), Chr("&H22"))
'                        cleanText1 = Replace(cleanText1, Chr("&H80"), "")
'                        cleanText1 = Replace(cleanText1, Chr("&H93"), Chr("&H22"))
'                        cleanText1 = Replace(cleanText1, Chr("&H94"), Chr("&H22"))
                         '''''cleanText1 = Replace(cleanText1, Chr("&H22"), "x")
                         ' replace bad characters from a circle-c copyright char, now 3 bad characters
                       weirdStr = Chr("&HEF") & Chr("&HBF") & Chr("&HBD")   ' these chars are translations? of circle-c (or so it seems)
                       copyRightStr = Chr("&H28") & "c" & Chr("&H29")   ' this is: (c)
                        'Debug.Print cleanText1
   '                    copyRightStr = Chr("&H28") & Chr("&HBF") & Chr("&H29"))
   '                     cleanText1 = Replace(cleanText1, weirdStr, copyRightStr)
                        'Debug.Print cleanText1
                    'replace the unicode character "Œ" if there
                    '    Range("tblFiles[Confirmed Packages]").Replace
                        On Error GoTo 0
                        cleanText1 = Replace(cleanText1, "Œ", vbNullString)
                        cleanText1 = Replace(cleanText1, "í", vbNullString)
                        cleanText1 = Replace(cleanText1, "h2>", "h5>")  ' make smaller the title
                        cleanText1 = Replace(cleanText1, "<h1><li>", "<h4>")
                        cleanText1 = Replace(cleanText1, "</li></h1>", "</h4>")
                                                
                        cleanText1 = Replace(cleanText1, """", "\""")
                        'cleanText1 = Replace(cleanText1, Hex(156), Hex(34))
                        
             If createFiles = True Then     ' make new json files, else only make json_tabs
                        fileOutLic.write " ""license_id"" : """ & licenseID & """," & vbCrLf
                        fileOutLic.write " ""license_name"" : """ & licenseName1 & """," & vbCrLf
                        ' write some taxonomy here
                        fileOutLic.write " ""taxonomy"" : """ & licenseTax & """," & vbCrLf
                        cleanText1 = Replace(cleanText1, """", "\""")
                        cleanText1 = Replace(cleanText1, "//", "/")
                        

                        fileOutLic.write " ""text"" : "" " & cleanText1 & """," & vbCrLf
                    '    On Error GoTo 0
                        If lCount = itemCount Then
                            fileOutLic.write "} ];" & vbCrLf
                        Else
                            fileOutLic.write "}, {" & vbCrLf
                        End If
            End If
            ' here write the items into cells on new tab
            With wbMacro.Sheets("json_licenseinfo")
                .Range("A1").Offset(row, 0).Value = licenseID
                .Range("A1").Offset(row, 1).Value = licenseName1
                .Range("A1").Offset(row, 2).Value = licenseTax
                .Range("A1").Offset(row, 3).Value = cleanText1
                .Range("A1").Offset(row, 4).Value = ""              ' cleanText2 for longer texts
             '   .Range("A1").Offset(row, 3).Width = 35.01
            End With
            ' if license text too long (> 32757 is limit for cell) then split into 2
            If Len(cleanText1) > 32700 Then     ' exceeds cell limit
                diff = Len(cleanText1) - 32700
                tempText = cleanText1
                cleanText1 = Left(tempText, 32700)
                cleanText2 = Right(tempText, diff)
                With wbMacro.Sheets("json_licenseinfo")
                    .Range("A1").Offset(row, 3).Value = cleanText1
                    .Range("A1").Offset(row, 4).Value = cleanText2             ' cleanText2 for longer texts
                End With
            End If      ' extra long license text
            
            DoEvents
            row = row + 1
                        GoTo getNextLicense 'done loop thru html
loopAgain:
                    Next i  ' next item in license html string
getNextLicense:
    Next cell       ' next license name
'***************** end of double loop
wbMacro.Sheets("json_licenseinfo").Range("D:D").WrapText = False
wbMacro.Sheets("json_licenseinfo").Range("E:E").WrapText = False


Range("A1").Select
Sheets("creds").Select
message1 = "Found " & row - 1 & " licenses; Expected " & lCount & " from " & itemCount & " total exported" & "."
MsgBox message1
Application.ScreenUpdating = True
Application.DisplayAlerts = True  ' turn back on
Exit Sub
abort1:
Application.ScreenUpdating = True
Application.DisplayAlerts = True  ' turn back on
message1 = "Abort. No LIcense tab found." & vbCrLf
message1 = message1 & "Required: Run 'Create CORE json_tabs' first."
MsgBox message1
End Sub
Sub createJSONPackages()
Dim createFilesNow As Boolean
createFilesNow = False
' read Package list and get (via API)each package's details and write to JSON file
' same for Licenses, except get read on HTML export

Dim message1, errMsg As String
Dim login_id, pass, loas As String
Dim packageID, packageName, packageDesc, cleanDesc, cleanDesc2 As String
Dim myRequest As Object
Dim cell, packageList As Range
Dim row As Long
Dim theResponseString As String
Dim url As String
Dim debug1 As Boolean
Dim fso, fso2 As Object     ' for output text files
Dim fileOut As Object
Dim filenameIn As String
Dim i, pCount, lCount, lineCountPackages As Long
Dim wbMacro As Workbook
Dim ws As Worksheet

Dim pathStr As String
Dim pathLen, tempLen, textEnd As Long
Dim theTag As String
Dim tempStr, homeURL As String
'Dim licenseID, licenseName, licenseText, countText As String
Dim itemCount As Long

Application.ScreenUpdating = False
Set wbMacro = ThisWorkbook

On Error GoTo errorOut2
    wbMacro.Worksheets("Packages").Select       ' required tab , else abort
On Error GoTo 0
debug1 = False  'default
message1 = "REQUIRED: 'creds' tab with login name, password, " & vbCrLf
message1 = message1 & " in 1st column, 2 rows." & vbCrLf
'message1 = message1 & "ALSO: before running, delete unnecessary rows from " & vbCrLf
'message1 = message1 & "   the Packages and Licenses tabs." & vbCrLf
message1 = message1 & "Continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

message1 = "Debug ON? "
If MsgBox(message1, vbYesNo + vbQuestion, "Turn ON debug (view responses)?") = vbYes Then debug1 = True

Application.ScreenUpdating = True   ' see progress

On Err GoTo errorOut
    login_id = ThisWorkbook.Worksheets("creds").Range("A1").Value
    pass = ThisWorkbook.Worksheets("creds").Range("A2").Value
    loas = ThisWorkbook.Worksheets("creds").Range("A3").Value
On Error GoTo 0

Application.DisplayAlerts = False  ' do not ask to confirm deletes
On Error Resume Next        ' if output tab already there, delete to rebuild
    wbMacro.Sheets("json_packageinfo").Delete
On Error GoTo 0
Application.DisplayAlerts = True  ' turn back on

'====write the json grid out ==============================================================
Set ws = wbMacro.Sheets.Add(After:=wbMacro.Sheets(wbMacro.Sheets.Count))
ws.Name = "json_packageinfo"
With wbMacro.Sheets("json_packageinfo")
    .Tab.Color = 65535  ' yellow tab
    ' headings
    .Range("A1") = "package_id"
    .Range("B1") = "package_name"
    .Range("C1") = "languages"
    .Range("D1") = "homepage_url"
    .Range("E1") = "description"
'    .Range("A1").Select
End With
    wbMacro.Sheets("json_packageinfo").Select
DoEvents

'On Error GoTo errorOut2
wbMacro.Worksheets("Packages").Select
'On Error Resume Next
wbMacro.Worksheets("Packages").ListObjects("tblPackages").Sort. _
    SortFields.Clear
wbMacro.Worksheets("Packages").ListObjects("tblPackages").Sort. _
    SortFields.Add Key:=Range("tblPackages[[#All],[ID]]"), SortOn:= _
    xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With wbMacro.Worksheets("Packages").ListObjects("tblPackages").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

DoEvents

pCount = Range("tblPackages[ID]").Count
If pCount < 1 Then  ' get out if not at least 1 package listed
    MsgBox ("Aborting program: 0 packages found.")
    Exit Sub
End If

lineCountPackages = 0   ' count packages thru loop
i = 0 ' init
On Error GoTo 0

If createFilesNow = True Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileOut = fso.createtextfile(ThisWorkbook.Path & "\json_packageinfo.js", True, True)
    fileOut.write "var json_packageinfo = [ {" & vbCrLf ' one-time header
End If

row = 1
'Application.VBE.Windows.Item("Immediate").Open      ' open immediate window
'Application.VBE.Windows.Item("Immediate").SetFocus      ' open immediate window

Range("tblPackages[ID]").Select       ' select all packages here
Set packageList = Selection
For Each cell In packageList
        cell.Activate   ' to show user the progress
        'DoEvents
        GoTo startMAChere
        '==================================================================================
        ''''''''' call get4windows here============================================================
        '!!!!!!!!!!!end of get4windows get call ===============================================================
startMAChere:
        Dim sCmd As String
        Dim lExitCode As Long
        Dim xml2 As String
        'packageID = "itext"
        packageID = cell.Value2
        url = cell.Offset(0, 6).Value2  ' this is home page URL
        'sCmd = "curl file:///Users/valerie/licenses-sample.html"
        sCmd = "curl -X GET -u """ & login_id & ":" & pass & """ -H ""Accept: application/vnd.openlogic.olexgovernance+xml""   -H ""Content-Type: application/vnd.openlogic.olexgovernance+xml"" " & url
        
        'sCmd = "curl -X GET -u ""bomgar+valerie@openlogicsoftware.com:openlogic1"" -H ""Accept: application/vnd.openlogic.olexgovernance+xml""   -H ""Content-Type: application/vnd.openlogic.olexgovernance+xml"" https://audit.openlogic.com/packages/" & packageID
        theResponseString = execShell(sCmd, lExitCode)
        
        Debug.Print packageID
        Debug.Print Now
        '}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}} jan2
         On Error Resume Next
             errMsg = ""
             errMsg = getTag2("message>", theResponseString) ' if fails, error message here
             If errMsg <> "" Then
                 MsgBox ("Package not found: " & packageID)
                ' get out - try next package
                GoTo escapeOUT
             End If
                       
             packageName = "" ' reset
             packageName = getTag2("name>", theResponseString) ' this will fail if package not found
             
        ' this will fail if package return no home URL, or description
             homeURL = "N/A"     ' set default if fails
             homeURL = getTag2("homepage_url>", theResponseString)
             
             packageDesc = "N/A" ' set default if fails
             'description should also include usage notes:
             packageDesc = getTag2("description>", theResponseString)
        '     packageDesc = Replace(packageDesc, vbLf, "\n")
        '    packageDesc = Replace(packageDesc, vbCr, "\n")
        '    packageDesc = Replace(packageDesc, vbCrLf, "\n")
             packageDesc = Replace(packageDesc, "   <![CDATA[", "")    ' unneeded XML data
             packageDesc = Replace(packageDesc, "]]>", "")    ' unneeded XML data
             'packageDesc = Replace(packageDesc, "href=", "href=\")
             cleanDesc = Application.WorksheetFunction.Clean(packageDesc)
             'Debug.Print packageDesc
             'Debug.Print cleanDesc
             
        If createFilesNow = True Then
            fileOut.write " ""package_id"" : """ & packageID & """," & vbCrLf
            fileOut.write " ""package_name"" : """ & packageName & """," & vbCrLf
            fileOut.write " ""languages"" : [ ]," & vbCrLf
            fileOut.write " ""homepage_url"" : """ & homeURL & """," & vbCrLf
        '    fileOut.write " ""description"" : """ & cleanDesc & """" & vbCrLf
            packageDesc = Replace(packageDesc, """", "\""")
            fileOut.write " ""description"" : """ & packageDesc & """" & vbCrLf
         On Error GoTo 0 ' cancel the skip errors
            If i <> (pCount - 1) Then fileOut.write " },{" & vbCrLf   ' setup the next
            
escapeOUT:          ' branch here if package not found
            If i = (pCount - 1) Then fileOut.write "} ];" & vbCrLf  'all done
        End If
        '******* write the json grid *********************************
        With wbMacro.Sheets("json_packageinfo")
            .Range("A1").Offset(row, 0).Value = packageID
 '           .Range("A1").Offset(row, 0).Select
            .Range("A1").Offset(row, 1).Value = packageName
            .Range("A1").Offset(row, 2).Value = ""      ' languages
            .Range("A1").Offset(row, 3).Value = homeURL
            .Range("A1").Offset(row, 4).Value = packageDesc
        End With
        row = row + 1       ' output json grid row#
        '*************************************************************
            i = i + 1
            lineCountPackages = lineCountPackages + 1
            Set myRequest = Nothing
        '    Set theResponse = Nothing
        
        If debug1 = True Then MsgBox theResponseString
            'message1 = "Quit now for test? " & vbCrLf
'            Sheets("json_packageinfo").Select       ' select all licenses there
'            wbMacro.Sheets("json_packageinfo").Range("E:E").WrapText = False
'            'Range("A1").Select
            'DoEvents
'            Application.ScreenUpdating = True
'            DoEvents
    If i = 10 Then
            message1 = "Stop now? "
            If MsgBox(message1, vbYesNo + vbQuestion, "Bail Out Early?") = vbYes Then GoTo doneHere
'            Else
'                   Sheets("Packages").Select       ' go back to packagelist
'                   cell.Select
'                   cell.Activate
            End If
            DoEvents

Next cell
'======================================================================================================
'fileOut.Close
doneHere:

wbMacro.Sheets("json_packageinfo").Range("E:E").WrapText = False
Range("A1").Select

'MsgBox ("Complete - see new json_packageinfo.js files in same directory")
MsgBox "Complete"
'=============================end dec11
Exit Sub 'cuz no error
errorOut:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True  ' turn back on
    MsgBox ("Failed to find login creds")
    Exit Sub
    
errorOut2:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True  ' turn back on
    message1 = "Packages' tab not found. Aborting." & vbCrLf
    message1 = message1 & "Required: Run 'Create CORE json_tabs' first."
    MsgBox message1
End Sub
Function execShell(command As String, Optional ByRef exitCode As Long) As String
Dim file As Long
file = popen(command, "r")
If file = 0 Then
Exit Function
End If
While feof(file) = 0
Dim chunk As String
Dim read As Long
chunk = Space(50)
read = fread(chunk, 1, Len(chunk) - 1, file)
If read > 0 Then
chunk = Left$(chunk, read)
execShell = execShell & chunk
End If
Wend
exitCode = pclose(file)
End Function

Function getLicNameTag2(tag, theResponse As String, positionX As Long) As String
Dim start1, end1, len1 As Long
'Dim tag2, endTag As String

'tag2 = "<" & tag
start1 = InStr(positionX, theResponse, tag)
start1 = start1 + Len(tag)
'start1 = start1 + 16
end1 = InStr((start1), theResponse, "</li")
'end1 = end1 - 8
len1 = end1 - start1
getLicNameTag2 = Mid(theResponse, start1, len1)
position = end1
textPosition = start1 - 8
End Function


Function getTagSetPosition(tag, theResponse As String, positionX As Long) As String
Dim start1, end1, len1 As Long
'Dim tag2, endTag As String

'tag2 = "<" & tag
start1 = InStr(positionX, theResponse, tag)
start1 = start1 + Len(tag)
start1 = start1 + 7
end1 = InStr((start1), theResponse, ">")
end1 = end1 - 1
len1 = end1 - start1
getTagSetPosition = Mid(theResponse, start1, len1)
position = end1
End Function

Function getTagSetPosition2(tag, theResponse As String, positionX As Long) As String
Dim start1, end1, len1 As Long
'Dim tag2, endTag As String

'tag2 = "<" & tag
start1 = InStr(positionX, theResponse, tag)
start1 = start1 + Len(tag)
start1 = start1
end1 = InStr((start1), theResponse, ">")
end1 = end1 - 1
len1 = end1 - start1
getTagSetPosition2 = Mid(theResponse, start1, len1)
position = end1
End Function

Private Function getTag2(tag, theResponse As String) As String
Dim start1, end1, len1 As Long
'Dim tag2, endTag As String

'tag2 = "<" & tag
start1 = InStr(20, theResponse, tag)
start1 = start1 + Len(tag)
end1 = InStr((start1), theResponse, "</" & tag)
len1 = end1 - start1

getTag2 = Mid(theResponse, start1, len1)
End Function

Private Function getTag1(tag, theResponse As String) As String
Dim start1, end1, len1 As Integer

start1 = InStr(200, theResponse, tag)
start1 = start1 + Len(tag)
end1 = InStr((start1), theResponse, "</")
len1 = end1 - start1
getTag1 = Mid(theResponse, start1, len1)
End Function

Sub makeCVE2Table()
Attribute makeCVE2Table.VB_ProcData.VB_Invoke_Func = " \n14"
Dim theRange As Range
'
' make this permanent at end of CVE json tab logic
'
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Set theRange = Selection
    ActiveSheet.ListObjects.Add(xlSrcRange, theRange, , xlYes).Name = "tblCVEDetail"
    ActiveSheet.ListObjects(1).TableStyle = "TableStyleLight9"
  
ActiveWorkbook.Worksheets("json_cve").ListObjects("tblCVEDetail").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("json_cve").ListObjects("tblCVEDetail").Sort. _
        SortFields.Add Key:=Range("tblCVEDetail[[#All],[PackageName]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("json_cve").ListObjects("tblCVEDetail").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("tblCVEDetail[[#Headers],[PackageID]]").Select
End Sub
Sub SaveCVE_Report()
Attribute SaveCVE_Report.VB_ProcData.VB_Invoke_Func = " \n14"
Dim path2save As String
'
' SaveCVE_Report Macro
' fix to save in same directory, with loas# in name
path2save = "Macintosh HD:Users:valerie:Documents:CVEReport.xlsx"
    Sheets("json_cve").Copy
 '   ChDir "Macintosh HD:Users:valerie:Documents:"
    ActiveWorkbook.SaveAs Filename:=path2save, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    ActiveWorkbook.Save
    Sheets("GET_CVEs").Select
    MsgBox "Complete. Find new CVE report in same directory."
End Sub
