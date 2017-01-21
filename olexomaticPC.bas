Attribute VB_Name = "OLEXomaticPC"
Option Explicit

Dim position As Long
Dim textPosition As Long

Sub createCVE_JSONtab()
Dim lookupRange, theRange, tempRange, rng1, cell As Range
Dim tblTemp As ListObject
Dim lookString, packageID, packageName As String
Dim ws As Worksheet
Dim lastRow As Long
Dim message1 As String

message1 = message1 & "Create/Replace the json_cve tab and CVE Details Report. YES ot Continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

ActiveSheet.Range("H9").Value = Now

Application.ScreenUpdating = False
Application.DisplayAlerts = False  ' do not ask to confirm deletes
On Error Resume Next        ' if already there, delete to rebuild
    Sheets("json_cve").Delete
On Error GoTo 0
Application.DisplayAlerts = True  ' turn back on

Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
ws.Name = "json_cve"
Sheets("json_cve").Select
With Sheets("json_cve")
    .Tab.Color = 65535  ' yellow tab
    ' headings
End With
DoEvents

Sheets("GET_CVEs").Select
Range("F12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Clear ' clear column to add package name
'Range("A11").Select

'
Range("E12").Select
Range(Selection, Selection.End(xlDown)).Select
For Each cell In Selection
    packageID = cell.Value2
    Set rng1 = Range("A:A").Find(packageID, , xlValues, xlWhole)
        If Not rng1 Is Nothing Then 'found
            packageName = rng1.Offset(0, 1).Value2
        Else    'not found
            MsgBox packageID & " not found"
            packageName = ""
        End If
    cell.Offset(0, 1).Value = packageName
Next cell

';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Range("E11").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
'tempRange.Copy
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
'    Range("tblCVEDetail[AccessComplexity]").Select
'    Selection.FormatConditions.Add Type:=xlTextString, String:="LOW", TextOperator:=xlContains
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    With Selection.FormatConditions(1)
'        .Font.Color = RGB(156, 20, 29)
'        .Interior.Color = RGB(254, 199, 206)
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
'    Selection.FormatConditions.Add Type:=xlTextString, String:="HIGH", TextOperator:=xlContains
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    With Selection.FormatConditions(1)
'        .Font.Color = RGB(25, 96, 22)   'dark green
'        .Interior.Color = RGB(199, 238, 206)    ' lite green
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
'    Selection.FormatConditions.Add Type:=xlTextString, String:="MEDIUM", TextOperator:=xlContains
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    With Selection.FormatConditions(1)
'        .Font.Color = RGB(155, 106, 59)         ' dark yellow
'        .Interior.Color = RGB(254, 234, 160)        ' lite yellow
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
'Selection.Font.Bold = True
Range("A1").Select

Call SaveCVE_Report ' copy and save tab as CVEReport

Application.ScreenUpdating = True
End Sub
Sub SaveCVE_Report()
Dim tempPath, path2save As String
'
' SaveCVE_Report Macro
' fix to save in same directory, with loas# in name
tempPath = ActiveWorkbook.Path & "/CVEReport.xlsx"
path2save = tempPath
'path2save = "Macintosh HD:Users:valerie:Documents:CVEReport.xlsx"
    Sheets("json_cve").Copy
 '   ChDir "Macintosh HD:Users:valerie:Documents:"
    ActiveWorkbook.SaveAs Filename:=path2save, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    ActiveWorkbook.Save
    Sheets("GET_CVEs").Select
    MsgBox "Complete. Find new CVE report in same directory."
End Sub

Sub getCVEs()
Dim myRequest, myRequest2 As Object
Dim theResponse As New DOMDocument
Dim theResponse2 As New DOMDocument
Dim cell, tempRange As Range
Dim myNode, myNode2 As IXMLDOMNode
Dim packageList, packageList2, securityList, versionList As IXMLDOMNodeList
'Dim securityList As IXMLDOMNodeList
Dim cveList As IXMLDOMNodeList
Dim xmlAttribute As IXMLDOMAttribute
Dim pathStr As String
Dim pathLen, tempLen As Integer
Dim url, id, pass, loas, packageID, cveStr, tempStr, severity, versionStr As String
Dim i, x As Integer
Dim riskNum, message1 As String
Application.ScreenUpdating = False
'clear any old results before starting
On Error Resume Next
    ActiveSheet.Range("packIDList").ClearContents       ' delete if there
On Error GoTo 0
ActiveSheet.Range("F1").Value = Now

ActiveSheet.Range("A12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("B12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("E12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("F12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("G12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("H12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("I12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("J12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("K12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("L12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("M12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("N12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("O12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("P12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("Q12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("R12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

ActiveSheet.Range("loas").Select    'land here
DoEvents

id = ActiveSheet.Range("loginID").Value
pass = ActiveSheet.Range("password").Value
loas = ActiveSheet.Range("loas")
url = "https://audit.openlogic.com/loas/" & loas
    Set myRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    myRequest.Open "GET", url
'    myRequest.SetCredentials "valerie@openlogicsoftware.com", "openlogic1", 0
    myRequest.SetCredentials id, pass, 0
    myRequest.SetRequestHeader "Content-Type", "application/vnd.openlogic.olexgovernance+xml"
    myRequest.SetRequestHeader "Accept", "application/vnd.openlogic.olexgovernance+xml"

    myRequest.Send      'pull trigger
        
    theResponse.LoadXML myRequest.responseText
'    MsgBox myRequest.responseText

    Set packageList = theResponse.getElementsByTagName("package")
    ActiveSheet.Range("packageCount").Value = "#packages:    " & packageList.Length
    i = 0
  For Each myNode In packageList
        Set xmlAttribute = packageList.Item(i).Attributes.getNamedItem("id")
        pathStr = xmlAttribute.Text
        pathLen = Len(pathStr)
        packageID = Right(pathStr, (pathLen - 37))
        ' 37 = fixed length of path chars to drop= "https://audit.openlogic.com/packages/"
        ActiveSheet.Range("A12").Offset(i, 0).Select
'        Selection.Value = xmlAttribute.Text
        Selection.Value = packageID
        i = i + 1
    Next myNode
    i = 0
  For Each myNode In packageList
 '       Set xmlAttribute = packageList.Item(i).Attributes.getNamedItem("id")
        ActiveSheet.Range("A12").Offset(i, 1).Select
        Selection.Value = packageList.Item(i).Text
        i = i + 1
    Next myNode
' create the named range for packageID's
ActiveSheet.Range("A12").Select
Range(Selection, Selection.End(xlDown)).Select
Set tempRange = Selection
tempRange.Name = "packIDList"   ' name for later use

Set myNode2 = theResponse.getElementsByTagName("version_name").Item(0)

On Error GoTo errorMsg1 ' failure will occur here if bad load or account or somesuch
ActiveSheet.Range("scan_version_name").Value = myNode2.Text

Set myNode2 = theResponse.getElementsByTagName("created_at").Item(0)
ActiveSheet.Range("created_at").Value = myNode2.nodeTypedValue

On Error GoTo 0         ' resume normal error processing

Set myNode2 = theResponse.getElementsByTagName("updated_at").Item(0)
ActiveSheet.Range("updated_at").Value = myNode2.nodeTypedValue
'Set theResponse = Nothing
'theResponse = New DOMDocument
'______________add the cve list_______________________________________________________

url = "https://audit.openlogic.com/loas/" & loas & "/security_risks"
    Set myRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    myRequest.Open "GET", url
 '   myRequest.SetCredentials "valerie@openlogicsoftware.com", "password1", 0
    myRequest.SetCredentials id, pass, 0
    myRequest.SetRequestHeader "Content-Type", "application/vnd.openlogic.olexgovernance+xml"
    myRequest.SetRequestHeader "Accept", "application/vnd.openlogic.olexgovernance+xml"

    myRequest.Send      'pull trigger
    DoEvents
    theResponse.LoadXML myRequest.responseText
    DoEvents
'    MsgBox myRequest.responseText
    Set securityList = theResponse.getElementsByTagName("security_risk")
    ActiveSheet.Range("securityRiskCount").Value = "#risks:    " & securityList.Length
  
  i = 0
  For Each myNode In securityList
        Set xmlAttribute = securityList.Item(i).Attributes.getNamedItem("id")
        ActiveSheet.Range("F12").Offset(i, 0).Select
        Selection.Value = xmlAttribute.Text
        i = i + 1
    Next myNode
    
Debug.Print "1st half COMPLETE AT : "; Now
Application.ScreenUpdating = True       ' allow user to see progress

ActiveSheet.Range("A12").Select
DoEvents
    message1 = message1 & " Found " & securityList.Length & " CVEs. Fetch details now?" & vbCrLf
    message1 = message1 & "  note: expect up to 1 minute per 30 CVE's" & vbCrLf
    message1 = message1 & "Do you want to continue now?"
    If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

'-----------nov9---------------------------------------------------------
' for each item in security list, get url
  i = 0
  For Each myNode In securityList
        Set xmlAttribute = securityList.Item(i).Attributes.getNamedItem("id")
        tempStr = xmlAttribute.Text
        tempLen = Len(tempStr)
        cveStr = Right(tempStr, (tempLen - 43))
        ' 43 = fixed length of path chars to drop= "https://audit.openlogic.com/security_risks/"
        riskNum = cveStr
'        riskNum = "44"
        url = "https://audit.openlogic.com/security_risks/" & riskNum
            Set myRequest2 = CreateObject("WinHttp.WinHttpRequest.5.1")
            myRequest2.Open "GET", url
            'myRequest2.SetCredentials "bomgar+valerie@openlogicsoftware.com", "openlogic1", 0
            myRequest2.SetCredentials id, pass, 0
            myRequest2.SetRequestHeader "Content-Type", "application/vnd.openlogic.olexgovernance+xml"
            myRequest2.SetRequestHeader "Accept", "application/vnd.openlogic.olexgovernance+xml"

            myRequest2.Send      'pull trigger
            DoEvents
            
            theResponse2.LoadXML myRequest2.responseText
            DoEvents
            Set myNode2 = theResponse2.getElementsByTagName("cve_number").Item(0)
            cveStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 3).Select
            Selection.Value = cveStr
        
            Set myNode2 = theResponse2.getElementsByTagName("score").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 4).Select
            Selection.Value = tempStr
       ' set severity from score (stored in tempStr)
            If tempStr < 4 Then
                severity = "LOW"
            Else
                If tempStr > 6 Then
                    severity = "HIGH"
                Else
                    severity = "MEDIUM"
                End If
            End If
            ActiveSheet.Range("F12").Offset(i, 5).Select
            Selection.Value = severity
            
            Set myNode2 = theResponse2.getElementsByTagName("access_complexity").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 6).Select
            Selection.Value = tempStr
       
            Set myNode2 = theResponse2.getElementsByTagName("updated_at").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 7).Select
            Selection.Value = tempStr
       
            Set myNode2 = theResponse2.getElementsByTagName("nvd_link").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 8).Select
            Selection.Value = tempStr
            
            Set myNode2 = theResponse2.getElementsByTagName("created_at").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 9).Select
            Selection.Value = tempStr
       
            Set myNode2 = theResponse2.getElementsByTagName("summary").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 10).Select
            
            Selection.Value = tempStr
            Set myNode2 = theResponse2.getElementsByTagName("access_vector").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 11).Select
            Selection.Value = tempStr
       
            Set myNode2 = theResponse2.getElementsByTagName("integrity_impact").Item(0)
            tempStr = myNode2.Text
            ActiveSheet.Range("F12").Offset(i, 12).Select
            Selection.Value = tempStr
       
 ' ------nov11---1
        Set packageList2 = theResponse2.getElementsByTagName("package") '
        If packageList2.Length = 1 Then
            ' we got only 1 package assoiciated with this CVE, so name this packageID in line item
            Set xmlAttribute = packageList2.Item(0).Attributes.getNamedItem("id")
            pathStr = xmlAttribute.Text
            pathLen = Len(pathStr)
            packageID = Right(pathStr, (pathLen - 37))
            ' 37 = fixed length of path chars to drop= "https://audit.openlogic.com/packages/"
            ActiveSheet.Range("F12").Offset(i, -1).Select
    '        Selection.Value = xmlAttribute.Text
            Selection.Value = packageID
        End If
        x = 0       ' x is counter for packages listed for this CVE
'--- nov11-2 ----------------------
        If packageList2.Length > 1 Then
            ' then we must loop thru packages for this CVE to find 1 that matches our loas
            For Each myNode2 In packageList2
                Set xmlAttribute = packageList2.Item(x).Attributes.getNamedItem("id")
                tempStr = xmlAttribute.Text ' this is a full URL, and we want only the id on end
                tempLen = Len(tempStr)
                'ActiveSheet.Range("F12").Offset(i, -1).Select  ' target to paste
                ' 37 = fixed length of path chars to drop= "https://audit.openlogic.com/packages/"
                packageID = Right(tempStr, (tempLen - 37))      ' pull id off right end of URL
                ' now check if this packageID in contained anywhere in packageList for this loas
                ' because each CVE might be associated with many packages, included some not found in this loa
                ' goal is to list only the package that matches in package list for this loas.
                For Each cell In Range("packIDList").Cells
                    If packageID = cell.Value Then
                        ' got hit - use this packageID
                        ActiveSheet.Range("F12").Offset(i, -1).Select       ' select target to paste value
                        Selection.Value = packageID     ' line-item: correct packageID for this CVE
                        packageID = "done here"    ' erase this ID, so forces If-loop exit, thus ignore additional packages
                    End If
                Next cell
                x = x + 1       'x = items in packageList2, packages for this CVE
            Next myNode2
        End If
'----nov11-- end --------------------------------
'========nov13
        tempLen = 0
        pathStr = ""
        versionStr = ""
        Set versionList = theResponse2.getElementsByTagName("package_version")
        tempLen = versionList.Length
        If tempLen = 0 Then ' got no versions
            versionStr = "No versions listed"
        Else    ' got some versions to loop thru
            tempLen = tempLen
            x = 0   ' version counter
            For Each myNode2 In versionList
                Set xmlAttribute = versionList.Item(x).Attributes.getNamedItem("id")
                pathStr = xmlAttribute.Text
                pathLen = Len(pathStr)
                tempStr = Right(pathStr, (pathLen - 37))
                versionStr = versionStr & tempStr & "  "
                ' 37 = fixed length of path chars to drop= "https://audit.openlogic.com/packages/"
                x = x + 1
            Next myNode2
        End If
'write the versions
        ActiveSheet.Range("F12").Offset(i, 2).Select    'write versions to 2 cols over
'        Selection.Value = xmlAttribute.Text
        Selection.Value = versionStr
        tempStr = ""
'============end nov13
        i = i + 1       ' i = items in securityList
    Next myNode
ActiveSheet.Range("securityRiskCount").Value = "#risks:    " & securityList.Length
ActiveSheet.Range("loas").Select    'land here
Debug.Print "COMPLETE AT : "; Now
Application.ScreenUpdating = True
ActiveSheet.Range("A12").Select 'landing
ActiveSheet.Range("loas").Select    'land here
MsgBox "Complete"
Exit Sub
errorMsg1:
    MsgBox "Abort. Check loas#, login, or connection."
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
Application.ScreenUpdating = False
Application.DisplayAlerts = False  ' do not ask to confirm deletes

createFiles = False
message1 = " CONVERT .xlsx Report to JSONtabs." & vbCrLf
message1 = message1 & " Please next select a completed OLEX Report (Excel)" & vbCrLf
message1 = message1 & "Ready to continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub
reportIn = Application.GetOpenFilename _
(Title:="Please choose the Report to convert", _
FileFilter:="Excel Files *.xlsx (*.xlsx),")
If reportIn = "" Then
    MsgBox "No file selected.", vbExclamation, "Sorry!"
    Exit Sub
Else
Workbooks.Open Filename:=reportIn
End If

DoEvents        ' to make sure it opens

Set wbReport = ActiveWorkbook
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
    .Range("A9") = "ReportDate"
    .Range("A10") = "ReportIntro"
    .Range("A11") = "Disclaim1"
    .Range("A12") = "Disclaim2"
    .Range("A13") = "HelpBOM"
    .Range("A14") = "HelpCVE"

'    Range("B1") = Sheets("Summary").Range("Client Name").Text
    .Range("B1") = "Client Name"
    .Range("B2") = wbReport.Sheets("Summary").Range("projectName").Value2
    .Range("B3") = wbReport.Sheets("Summary").Range("projectFilenames").Value2
    .Range("B4") = wbReport.Sheets("Summary").Range("projectSourceFiles").Value2
    .Range("B5") = wbReport.Sheets("Summary").Range("projectChecksum").Value2
'    .Range("B5") = "SHA1 Checksum"
    .Range("B6") = wbReport.Sheets("Summary").Range("Summary.FileCount").Value2
    .Range("B7") = wbReport.Sheets("Summary").Range("Summary.PackageCount").Value2
    .Range("B8") = wbReport.Sheets("Summary").Range("Summary.LicenseCount").Value2
    
    .Range("B9") = wbReport.Sheets("Summary").Range("Summary.ReportDate").Value
    .Range("B9").NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    
    .Range("B10") = wbReport.Sheets("Summary").Range("Summary.ReportIntro").Value
    .Range("B11") = wbReport.Sheets("Summary").Range("Summary.ReportDisclaim1").Value
    .Range("B12") = wbReport.Sheets("Summary").Range("Disclaim2").Value
    tempString = wbReport.Sheets("Summary").Range("HelpBOM").Value
    tempString = "<pre>" & tempString
    .Range("B13") = tempString
'    .Range("B13") = wbReport.Sheets("Summary").Range("HelpBOM").Value
    tempString = wbReport.Sheets("Summary").Range("HelpCVE").Value
    tempString = "<pre>" & tempString
    .Range("B14") = tempString
'    .Range("B14") = wbReport.Sheets("Summary").Range("HelpCVE").Value

    Range("A1:B14").Select
    Selection.Copy
    Range("A15").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Rows("1:14").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("A1").Select


'Sheets("json_summary").Range("A:A").ColumnWidth = 30
'Sheets("json_summary").Range("B:B").WrapText = False
'Sheets("json_summary").Range("B:B").ColumnWidth = 25

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
        .Rows(1).Replace What:=" ", Replacement:="", _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
            False, ReplaceFormat:=False
    
        .Rows(1).Replace What:="(s)", Replacement:="s", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
 End With
'    Range("A1").Select
'now do vlookups using Packages Tab to add packageID and licenseID to above table
' lookup packageid and licenseid
    Sheets("json_bom").Select
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
        .Replace What:=" ", Replacement:="", _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
            False, ReplaceFormat:=False
        .Replace What:="(s)", Replacement:="s", _
            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
            False, ReplaceFormat:=False
    End With
'now do vlookups to add licenseID to above table
'    Sheets("json_obligations").Select
 '*****************************dec27

'    Range("A2").Select
    lastRow = wbMacro.Sheets("json_obligations").Cells(Rows.Count, 3).End(xlUp).row
    lookup_col = wbMacro.Sheets("Licenses").Range("tblLicenses[ID]").Column
    wbMacro.Sheets("json_obligations").Range("A2").FormulaR1C1 = "=VLOOKUP(RC[2],tblLicenses," & lookup_col & ",FALSE)"
    ' copy formula all the way down
    Range("A2:A" & lastRow).FillDown
    DoEvents
    Range("A2:A" & lastRow).Value = Range("A2:A" & lastRow).Value   ' paste values over vlookup formula
    
    Range("C2").Select
    tempRange = Range(Selection, Selection.End(xlDown)).Select

 '   tempRange = Range("C2:C" & lastRow).FillDown
 '   tempRange = Selection
 'jan20
    Columns("C:C").Select
    ActiveWorkbook.Worksheets("json_obligations").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("json_obligations").Sort.SortFields.Add Key:=Range( _
        "C1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("json_obligations").Sort
        .SetRange Range("A2:I2" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
    
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
    
    wbReport.Close Savechanges:=False
    Range("A1").Select
    Debug.Print "Convert byLicense to json"
    
    Application.ScreenUpdating = True
'    Selection = Nothing
    
MsgBox ("Completed json tabs")
End Sub

Sub createJSONLicenses()
Dim message1 As String
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

' do licenses to json here
message1 = "REQUIRED: licenses.html export " & vbCrLf
message1 = message1 & "  Select the licenes.html next." & vbCrLf
message1 = message1 & "Do you want to continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

Application.DisplayAlerts = False  ' do not ask to confirm deletes
On Error Resume Next        ' if already there, delete to rebuild
    wbMacro.Sheets("json_licenseinfo").Delete
On Error GoTo 0
Application.DisplayAlerts = True  ' turn back on

Dim licenseListCol As Range
Dim licenseID, licenseIDx, licenseName1, licenseNameHTML, licenseTax, licenseText As String
Dim theTag, tempText, cleanText1, cleanText2 As String

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

filenameIn = Application.GetOpenFilename()
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
Sheets("Licenses").Select
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
    .Range("E1") = "text2"
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
        
      '  Debug.Print "seeking: " & licenseID
        
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
 
     DoEvents
    ' cleanText1 = StrConv(cleanText1, vbUnicode)
     DoEvents
 '           cleanText1 = licenseText
        'replace the unicode character "Â" if there
        '    Range("tblFiles[Confirmed Packages]").Replace
            On Error GoTo 0
            cleanText1 = Replace(cleanText1, "Â", vbNullString)
            cleanText1 = Replace(cleanText1, "Ã", vbNullString)
            cleanText1 = Replace(cleanText1, "h2>", "h4>")  ' make smaller the title
            cleanText1 = Replace(cleanText1, "<h1><li>", "<h3>")
            cleanText1 = Replace(cleanText1, "</li></h1>", "</h3>")
            
            cleanText1 = Replace(cleanText1, "â€", """")
            cleanText1 = Replace(cleanText1, "â€œ", """")
            cleanText1 = Replace(cleanText1, "ï¿½", "©")
            cleanText1 = Replace(cleanText1, "œ", "")
'            cleanText1 = Replace(cleanText1, "&H0D", "")
'            cleanText1 = Replace(cleanText1, "&HAC", "")
            
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
    .Range("A1").Offset(row, 4).Value = ""      ' assumes no text2 cuz not > 32767 chars
End With
'max cell sixe = 32767 characters
If Len(cleanText1) > 32700 Then ' break into 2 parts/cells
    diff = Len(cleanText1) - 32700
    tempText = cleanText1   ' save full text
    cleanText1 = Left(tempText, 32700)    ' set part 1
    cleanText2 = Right(tempText, diff)      ' set part 2
    Debug.Print licenseName1 & " is " & Len(cleanText1)
    With wbMacro.Sheets("json_licenseinfo")
        .Range("A1").Offset(row, 3).Value = cleanText1
        .Range("A1").Offset(row, 4).Value = cleanText2
    End With
End If
DoEvents
row = row + 1
            GoTo getNextLicense 'done loop thru html
loopAgain:
        Next i  ' next item in license html string
getNextLicense:
    Next cell       ' next license name
'***************** end of double loop
' stop text wrap of long text:
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
message1 = "Abort. No License tab found." & vbCrLf
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
Dim cell As Range
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

Set wbMacro = ThisWorkbook
On Error GoTo errorOut2
    wbMacro.Worksheets("Packages").Select       ' required tab , else abort
On Error GoTo 0

debug1 = False
message1 = "REQUIRED: 'creds' tab with login name, password, " & vbCrLf
message1 = message1 & "  loas in 1st column, 2 rows." & vbCrLf
'message1 = message1 & "ALSO: before running, delete unnecessary rows from " & vbCrLf
'message1 = message1 & "   the Packages and Licenses tabs." & vbCrLf
message1 = message1 & "Continue?"
If MsgBox(message1, vbYesNo + vbQuestion, "Requirements:") = vbNo Then Exit Sub

message1 = "Do you wish to turn on debugging? " & vbCrLf
message1 = message1 & "  WARNING: this will slow down the run." & vbCrLf
If MsgBox(message1, vbYesNo + vbQuestion, "Debug?") = vbYes Then debug1 = True
    
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

wbMacro.Worksheets("Packages").Select
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
Range("tblPackages[ID]").Select       ' select all licenses there
On Error GoTo 0
If createFilesNow = True Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileOut = fso.createtextfile(ThisWorkbook.Path & "\json_packageinfo.js", True, True)
    fileOut.write "var json_packageinfo = [ {" & vbCrLf ' one-time header
End If
row = 1
For Each cell In Selection
'==================================================================================
' call get4windows here============================================================
    ' do a API 'get' for each package
    Dim theResponse As New DOMDocument  ' required for this to work: reference to Microsoft XML, v6.0
    Dim tempRange As Range
    Dim myNode As IXMLDOMNode
    Dim xmlAttribute As IXMLDOMAttribute
    Dim pathStr As String
    Dim pathLen, tempLen, textEnd As Long
    Dim theTag As String
    Dim tempStr, homeURL As String
    'Dim licenseID, licenseName, licenseText, countText As String
    Dim itemCount As Long
'    Set packageList = Selection
    cell.Activate
    packageID = cell.Value2
    ' todo: go get url from spreadsheet (to allow for private package = diff url
    'url = "https://audit.openlogic.com/packages/" & packageID
    url = cell.Offset(0, 6).Value2  ' this is home page URL
    If (url = "NOURL") Or (url = "") Then GoTo escapeOUT    ' no link here
    Set myRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    myRequest.Open "GET", url
    '    myRequest.SetCredentials "valerie@openlogicsoftware.com", "openlogic1", 0
    myRequest.SetCredentials login_id, pass, 0
    myRequest.SetRequestHeader "Content-Type", "application/vnd.openlogic.olexgovernance+xml"
    myRequest.SetRequestHeader "Accept", "application/vnd.openlogic.olexgovernance+xml"
    Debug.Print "get: " & packageID
    myRequest.Send      'pull trigger
    DoEvents
    Debug.Print Now
    theResponse.LoadXML myRequest.responseText
    theResponseString = myRequest.responseText
'!!!!!!!!!!!end of get4windows get call ===============================================================
    ' make MAC curl call here to get theResponseString
    
    If debug1 = True Then MsgBox theResponseString
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
'    packageDesc = Replace(packageDesc, vbCr, "\n")
'    packageDesc = Replace(packageDesc, vbCrLf, "\n")
     packageDesc = Replace(packageDesc, "   <![CDATA[", "")    ' unneeded XML data
     packageDesc = Replace(packageDesc, "]]>", "")    ' unneeded XML data
     'packageDesc = Replace(packageDesc, "href=", "href=\")
     cleanDesc = Application.WorksheetFunction.Clean(packageDesc)
     DoEvents
     packageDesc = Replace(cleanDesc, "&H0D", "")
     DoEvents
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
    
escapeOUT:  ' branch here if package not found
    If i = (pCount - 1) Then fileOut.write "} ];" & vbCrLf  'all done
End If
'******* write the json grid *********************************
With wbMacro.Sheets("json_packageinfo")
    .Range("A1").Offset(row, 0).Value = packageID
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
    Set theResponse = Nothing
If i = 5 Then       'offer bailout
    message1 = "Do you wish to quit now? " & vbCrLf
    If MsgBox(message1, vbYesNo + vbQuestion, "Bail Out Early?") = vbYes Then GoTo doneHere
End If

Next cell
'======================================================================================================
doneHere:
'fileOut.Close
wbMacro.Sheets("json_packageinfo").Range("E:E").WrapText = False
Range("A1").Select
'MsgBox ("Complete - see new json_packageinfo.js files in same directory")
MsgBox "Complete"
'=============================end dec11
Exit Sub 'cuz no error
errorOut:
    MsgBox ("Failed to find login creds")
    Exit Sub
    
errorOut2:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True  ' turn back on
    message1 = "Packages' tab not found. Aborting." & vbCrLf
    message1 = message1 & "Required: Run 'Create CORE json_tabs' first."
    MsgBox message1
End Sub
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


Function getLicIDTag2(tag, theResponse As String, positionX As Long) As String
Dim start1, end1, len1 As Long
'Dim tag2, endTag As String

'tag2 = "<" & tag
start1 = InStr(positionX, theResponse, tag)
start1 = start1 + Len(tag)
start1 = start1 + 7
end1 = InStr((start1), theResponse, ">")
end1 = end1 - 1
len1 = end1 - start1
getLicIDTag2 = Mid(theResponse, start1, len1)
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

